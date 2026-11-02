# -*- coding: ascii -*-
"""Tests for the gas review fixes of 2026-11-02.

Runs in plain CPython with the Revit assemblies stubbed (the stubs come from
_test_prv_zones). Covers:
  - the walk stays in gas piping: a gas appliance is a terminal, and a pipe or
    connector classified as a water, hydronic, sanitary or fire system is
    never followed or sized;
  - zero demand on a copper table at a long run takes the first size with a
    published capacity;
  - the table-only inverse diameter map (copper and CSST names);
  - revit_helpers.is_new_construction fails closed;
  - a rejected Parameter.Set() is a failure for pipes and fittings.
It does NOT exercise a real Revit model: the connectors, phases and
parameters are stand-ins. Live behaviour still needs a Revit run.

Run:  python -I _test_gas_traversal.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_prv_zones import _node, _edge, _finish, _Id   # installs the stubs

import gas_tables      # noqa: E402
import pipe_graph      # noqa: E402
import revit_helpers   # noqa: E402
import sizing_engine   # noqa: E402


# ---------------------------------------------------------------------------
# A tiny model: elements with typed classes and connector dicts
# ---------------------------------------------------------------------------
class _TypeName(object):
    def __init__(self, name):
        self.Name = name


class _Elem(object):
    def __init__(self, eid, kind, connectors, params=None, length=10.0):
        self.Id = _Id(eid)
        self.Location = None
        self.kind = kind                    # "Pipe" or "FamilyInstance"
        self.connectors = connectors
        self.params = params or {}
        self.length = length

    def GetType(self):
        return _TypeName(self.kind)


def _c(index, to, system="OtherPipe", direction="Bidirectional"):
    return {"connector_index": index, "direction": direction,
            "is_connected": to is not None, "connected_element_id": to,
            "connected_element_type": None, "origin_xyz": None,
            "system_type": system, "diameter_in": None,
            "connector_id": index}


class _Model(object):
    """Patches the revit_helpers reads pipe_graph makes, for one model."""

    def __init__(self, elements):
        self.elements = dict((e.Id.Value, e) for e in elements)

    def __enter__(self):
        self._saved = {}
        patches = {
            "get_connectors": lambda el: list(el.connectors),
            "get_parameter_value": lambda el, name: el.params.get(name),
            "get_pipe_length_feet": lambda el: el.length,
            "get_pipe_diameter_inches": lambda el: 0.5,
            "get_element_location": lambda el: None,
            "get_element_family_name": lambda el: el.params.get("family"),
        }
        for name, fn in patches.items():
            self._saved[name] = getattr(revit_helpers, name)
            setattr(revit_helpers, name, fn)
        self._saved_get = pipe_graph._get_element
        pipe_graph._get_element = lambda doc, eid: self.elements.get(eid)
        return self

    def __exit__(self, *args):
        for name, fn in self._saved.items():
            setattr(revit_helpers, name, fn)
        pipe_graph._get_element = self._saved_get


GAS_FIXTURE = {"IS_GAS_FIXTURE": True, "GAS_LOAD_MBH": 199.0,
               "FIXTURE_NAME": "WH-1", "family": "Water Heater"}


def _water_heater_model(heater_params=None):
    """meter 1 -> gas pipe 10 -> water heater 2
       water heater 2 CW -> cold pipe 20 -> tee 3 -> cold pipe 21 -> open
       water heater 2 HW -> hot pipe 30 -> open
    """
    meter = _Elem(1, "FamilyInstance",
                  [_c(0, None, direction="In"), _c(1, 10, direction="Out")],
                  {"family": "Gas Meter"})
    gas = _Elem(10, "Pipe", [_c(0, 1), _c(1, 2)])
    heater = _Elem(2, "FamilyInstance",
                   [_c(0, 10), _c(1, 20, "DomesticColdWater"),
                    _c(2, 30, "DomesticHotWater")],
                   heater_params if heater_params is not None
                   else dict(GAS_FIXTURE))
    cold = _Elem(20, "Pipe", [_c(0, 2, "DomesticColdWater"),
                              _c(1, 3, "DomesticColdWater")])
    tee = _Elem(3, "FamilyInstance",
                [_c(0, 20, "DomesticColdWater"), _c(1, 21, "DomesticColdWater"),
                 _c(2, None, "DomesticColdWater")], {"family": "Tee"})
    cold2 = _Elem(21, "Pipe", [_c(0, 3, "DomesticColdWater"),
                               _c(1, None, "DomesticColdWater")])
    hot = _Elem(30, "Pipe", [_c(0, 2, "DomesticHotWater"),
                             _c(1, None, "DomesticHotWater")])
    return [meter, gas, heater, cold, tee, cold2, hot]


class TestWalkStaysInGasPiping(unittest.TestCase):

    def _build(self, elements):
        with _Model(elements) as m:
            return pipe_graph.build_network(m.elements[1], None)

    def test_water_heater_is_a_terminal(self):
        g = self._build(_water_heater_model())
        self.assertEqual(sorted(g.edges.keys()), [10])
        self.assertNotIn(3, g.nodes)
        self.assertTrue(g.nodes[2].is_gas_fixture)
        self.assertEqual(g.disconnected, [])
        self.assertTrue(any("TERMINAL: appliance 2" in s
                            for s in g.traversal_log))

    def test_only_the_gas_pipe_is_sized(self):
        g = self._build(_water_heater_model())
        result = sizing_engine.size_system(
            g, "Schedule 40 Steel", 0.5, "402.4(2)",
            heat_content_btu_per_cf=1000.0)
        self.assertEqual(sorted(result["sizes"].keys()), [10])
        self.assertAlmostEqual(g.edges[10].cumulative_load_mbh, 199.0)

    def test_water_connector_on_a_fitting_is_not_followed(self):
        # The same heater WITHOUT IS_GAS_FIXTURE: it is walked as a fitting,
        # and its water connectors are still skipped by classification.
        params = {"family": "Water Heater"}
        g = self._build(_water_heater_model(params))
        self.assertEqual(sorted(g.edges.keys()), [10])
        skips = [s for s in g.traversal_log if s.startswith("SKIP: element 2")]
        self.assertEqual(len(skips), 2)

    def test_water_pipe_reached_directly_is_not_an_edge(self):
        # Gas tee 2 whose third connector is (wrongly) unclassified, joined to
        # a pipe that reads DomesticColdWater: the pipe is skipped.
        meter = _Elem(1, "FamilyInstance", [_c(0, 10, direction="Out")],
                      {"family": "Gas Meter"})
        gas = _Elem(10, "Pipe", [_c(0, 1), _c(1, 2)])
        tee = _Elem(2, "FamilyInstance", [_c(0, 10), _c(1, 11), _c(2, 20)],
                    {"family": "Tee"})
        gas2 = _Elem(11, "Pipe", [_c(0, 2), _c(1, 4)])
        rtu = _Elem(4, "FamilyInstance", [_c(0, 11)],
                    {"IS_GAS_FIXTURE": True, "GAS_LOAD_MBH": 100.0,
                     "FIXTURE_NAME": "RTU-1"})
        water = _Elem(20, "Pipe", [_c(0, 2, "DomesticColdWater"),
                                   _c(1, None, "DomesticColdWater")])
        g = self._build([meter, gas, tee, gas2, rtu, water])
        self.assertEqual(sorted(g.edges.keys()), [10, 11])
        self.assertNotIn(20, g.disconnected)
        self.assertTrue(any(s.startswith("SKIP: pipe 20")
                            for s in g.traversal_log))


class TestZeroDemandOnCopper(unittest.TestCase):
    """402.4(8) has no 3/4 K&L capacity from 550 ft. A zero load appliance
    at 600 ft used to abort the whole run."""

    def test_takes_the_first_published_size(self):
        g = pipe_graph.NetworkGraph()
        g.origin_id = 1
        _node(g, 1, "meter")
        _node(g, 2, "tee")
        _node(g, 3, "fixture", 50.0, "FURN-1")
        _node(g, 4, "fixture", 0.0, "SPARE")
        _edge(g, 10, 1, 2, 590.0)
        _edge(g, 11, 2, 3, 5.0)
        _edge(g, 12, 2, 4, 2.0)
        _finish(g)
        result = sizing_engine.size_system(
            g, "Semirigid Copper Tubing", 0.5, "402.4(8)",
            heat_content_btu_per_cf=1000.0)
        self.assertEqual(result["table_length_used_ft"], 600)
        self.assertEqual(result["sizes"][12], "1 K&L")
        detail = [d for d in result["segment_detail"] if d["pipe_id"] == 12][0]
        self.assertEqual(detail["capacity_mbh"],
                         gas_tables.get_capacity("402.4(8)", 595.0, "1 K&L"))


class TestTableOnlyInverseMap(unittest.TestCase):

    def test_copper_reads_as_copper(self):
        self.assertEqual(sizing_engine.nominal_for_diameter("402.4(8)", 1.0),
                         "1 K&L")
        self.assertEqual(sizing_engine.nominal_for_diameter("402.4(8)", 0.5),
                         "1/2 K&L")

    def test_csst_reads_as_csst_conservatively(self):
        # EHD-18 and EHD-19 are both 1/2 in.: the smaller one is kept.
        self.assertEqual(sizing_engine.nominal_for_diameter("402.4(15)", 0.5),
                         "EHD-18")

    def test_steel_unchanged(self):
        self.assertEqual(sizing_engine.nominal_for_diameter("402.4(2)", 1.0),
                         "1")

    def test_no_match(self):
        self.assertIsNone(sizing_engine.nominal_for_diameter("402.4(2)", 7.0))
        self.assertIsNone(sizing_engine.nominal_for_diameter("402.4(2)", None))


# ---------------------------------------------------------------------------
# Phase and parameter write stand-ins
# ---------------------------------------------------------------------------
class _Phases(object):
    def __init__(self, ids):
        self.ids = ids
        self.Size = len(ids)

    def get_Item(self, i):
        return _Elem(self.ids[i], "Phase", [])


class _Doc(object):
    def __init__(self, ids):
        self.Phases = _Phases(ids)


class _Phased(object):
    def __init__(self, eid, phase):
        self.Id = _Id(eid)
        if phase is not None:
            self.CreatedPhaseId = _Id(phase)


class TestIsNewConstruction(unittest.TestCase):

    def test_last_phase_is_new(self):
        self.assertTrue(revit_helpers.is_new_construction(
            _Phased(5, 200), _Doc([100, 200])))

    def test_earlier_phase_is_existing(self):
        self.assertFalse(revit_helpers.is_new_construction(
            _Phased(5, 100), _Doc([100, 200])))

    def test_unreadable_phase_fails_closed(self):
        self.assertFalse(revit_helpers.is_new_construction(
            _Phased(5, None), _Doc([100, 200])))
        self.assertFalse(revit_helpers.is_new_construction(
            _Phased(5, 200), _Doc([])))


class _Param(object):
    def __init__(self, accept):
        self.accept = accept
        self.IsReadOnly = False
        self.value = None

    def Set(self, value):
        self.value = value
        return self.accept


class _ParamElem(object):
    def __init__(self, eid, name, accept):
        self.Id = _Id(eid)
        self.param = _Param(accept)
        self.name = name

    def LookupParameter(self, name):
        return self.param if name == self.name else None

    def get_Parameter(self, bip):
        return None


class TestRejectedSetIsAFailure(unittest.TestCase):

    def setUp(self):
        revit_helpers.reset_pipe_diameter_approach()

    def test_pipe_set_rejected(self):
        ok, name = revit_helpers.set_pipe_diameter(
            _ParamElem(7, "Diameter", False), 1.0)
        self.assertFalse(ok)
        self.assertEqual(name, "FAILED")

    def test_pipe_set_accepted(self):
        pipe = _ParamElem(7, "Diameter", True)
        ok, name = revit_helpers.set_pipe_diameter(pipe, 1.0)
        self.assertTrue(ok)
        self.assertEqual(name, "LookupParameter")
        self.assertAlmostEqual(pipe.param.value, 1.0 / 12.0)

    def test_fitting_set_rejected(self):
        ok, name = revit_helpers.set_fitting_size(
            _ParamElem(8, "Nominal Radius", False), 1.0)
        self.assertFalse(ok)
        self.assertEqual(name, "FAILED")

    def test_fitting_set_accepted(self):
        fitting = _ParamElem(8, "Nominal Radius", True)
        ok, name = revit_helpers.set_fitting_size(fitting, 1.0)
        self.assertTrue(ok)
        self.assertEqual(name, "Nominal Radius")
        self.assertAlmostEqual(fitting.param.value, 0.5 / 12.0)


if __name__ == "__main__":
    unittest.main()
