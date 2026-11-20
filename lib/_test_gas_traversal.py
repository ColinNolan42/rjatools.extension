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
        # F1: the stop is recorded for the Size Gas warning
        self.assertEqual(g.non_gas_skipped, [(20, "DomesticColdWater")])

    def test_water_connector_on_a_fitting_is_recorded(self):
        # F1: the element side SKIP names the element that was not followed
        g = self._build(_water_heater_model({"family": "Water Heater"}))
        self.assertEqual(sorted(g.non_gas_skipped),
                         [(20, "DomesticColdWater"), (30, "DomesticHotWater")])

    def test_appliance_terminal_is_not_a_non_gas_stop(self):
        # A gas appliance's water connectors are by design (S21), no F1 line
        g = self._build(_water_heater_model())
        self.assertEqual(g.non_gas_skipped, [])


def _tee_model(cap_params, elbow_family=None):
    """meter 1 -> pipe 10 -> tee 2 -> pipe 11 -> RTU 4
                                   -> pipe 12 (-> elbow 5 -> pipe 13) -> cap 6
    """
    meter = _Elem(1, "FamilyInstance", [_c(0, 10, direction="Out")],
                  {"family": "Gas Meter"})
    gas = _Elem(10, "Pipe", [_c(0, 1), _c(1, 2)])
    tee = _Elem(2, "FamilyInstance", [_c(0, 10), _c(1, 11), _c(2, 12)],
                {"family": "Tee"})
    gas2 = _Elem(11, "Pipe", [_c(0, 2), _c(1, 4)])
    rtu = _Elem(4, "FamilyInstance", [_c(0, 11)],
                {"IS_GAS_FIXTURE": True, "GAS_LOAD_MBH": 100.0,
                 "FIXTURE_NAME": "RTU-1"})
    elements = [meter, gas, tee, gas2, rtu]
    if elbow_family:
        elements += [
            _Elem(12, "Pipe", [_c(0, 2), _c(1, 5)], length=30.0),
            _Elem(5, "FamilyInstance", [_c(0, 12), _c(1, 13)],
                  {"family": elbow_family}),
            _Elem(13, "Pipe", [_c(0, 5), _c(1, 6)], length=30.0),
        ]
    else:
        elements.append(_Elem(12, "Pipe", [_c(0, 2), _c(1, 6)]))
    elements.append(_Elem(6, "FamilyInstance", [_c(0, 13 if elbow_family
                                                   else 12)], cap_params))
    return elements


class TestBenchFixes(unittest.TestCase):
    """F2, F3 and F5 from the gas bench fix list (2026-11-18)."""

    def _build(self, elements):
        with _Model(elements) as m:
            return pipe_graph.build_network(m.elements[1], None)

    def test_cap_without_flag_is_listed(self):
        g = self._build(_tee_model({"family": "Cap - Generic"}))
        ends = pipe_graph.non_fixture_ends(g)
        self.assertEqual([n.element_id for n in ends], [6])

    def test_flagged_cap_is_not_listed(self):
        g = self._build(_tee_model({"IS_GAS_FIXTURE": True,
                                    "GAS_LOAD_MBH": 0.0,
                                    "FIXTURE_NAME": "FUTURE",
                                    "family": "RJA- Gas Equipment Cap"}))
        self.assertEqual(pipe_graph.non_fixture_ends(g), [])

    def _run_with_elbow(self, family):
        cap = {"IS_GAS_FIXTURE": True, "GAS_LOAD_MBH": 50.0,
               "FIXTURE_NAME": "UH-1"}
        return self._build(_tee_model(cap, elbow_family=family)).longest_run

    def test_elbow_detection_ignores_case(self):
        for family in ("Elbow - Generic", "ELBOW - Threaded",
                       "gas elbow", "RJA Long Radius elbow"):
            run = self._run_with_elbow(family)
            self.assertEqual(run["elbow_equiv_length_feet"], 5.0, family)
            # 70 ft pipe + 5 ft elbow + 3 ft for tee 2
            self.assertAlmostEqual(run["total_length_feet"], 78.0)

    def test_tee_into_the_branch_adds_three_feet(self):
        # Colin 2026-11-19: developed length = pipe + 5 ft per elbow + 3 ft
        # per tee. The run turns at tee 2 into pipe 12: 10 + 30 + 30 ft of
        # pipe, 1 elbow, 1 tee.
        cap = {"IS_GAS_FIXTURE": True, "GAS_LOAD_MBH": 50.0,
               "FIXTURE_NAME": "UH-1"}
        run = self._build(_tee_model(cap, elbow_family="Elbow - Generic")
                          ).longest_run
        self.assertEqual(run["farthest_fixture_name"], "UH-1")
        self.assertAlmostEqual(run["pipe_length_feet"], 70.0)
        self.assertEqual(run["elbow_count"], 1)
        self.assertEqual(run["tee_count"], 1)
        self.assertAlmostEqual(run["tee_equiv_length_feet"], 3.0)
        self.assertAlmostEqual(run["total_length_feet"], 78.0)

    def test_tee_straight_through_adds_three_feet(self):
        # The run goes on through tee 2 to RTU-1 (pipe 11 lengthened to 40
        # ft so it is the longest); the branch is only 10 + 10 ft.
        cap = {"IS_GAS_FIXTURE": True, "GAS_LOAD_MBH": 50.0,
               "FIXTURE_NAME": "UH-1"}
        els = _tee_model(cap)
        dict((e.Id.Value, e) for e in els)[11].length = 40.0
        run = self._build(els).longest_run
        self.assertEqual(run["farthest_fixture_name"], "RTU-1")
        self.assertAlmostEqual(run["pipe_length_feet"], 50.0)
        self.assertEqual(run["elbow_count"], 0)
        self.assertEqual(run["tee_count"], 1)
        self.assertAlmostEqual(run["total_length_feet"], 53.0)

    def test_two_tees_and_an_elbow_together(self):
        # meter 1 -10-> tee 2 -11-> tee 3 -14-> RTU 4
        #                  tee 2 -12-> cap 7 (short)
        #                  tee 3 -15-> elbow 5 -16-> UH-1 6 (longest)
        # 10 + 10 + 10 + 10 ft pipe + 5 (elbow) + 2 x 3 (tees) = 51 ft.
        cap = {"IS_GAS_FIXTURE": True, "GAS_LOAD_MBH": 50.0,
               "FIXTURE_NAME": "UH-1"}
        els = [
            _Elem(1, "FamilyInstance", [_c(0, 10, direction="Out")],
                  {"family": "Gas Meter"}),
            _Elem(10, "Pipe", [_c(0, 1), _c(1, 2)]),
            _Elem(2, "FamilyInstance", [_c(0, 10), _c(1, 11), _c(2, 12)],
                  {"family": "Tee"}),
            _Elem(11, "Pipe", [_c(0, 2), _c(1, 3)]),
            _Elem(3, "FamilyInstance", [_c(0, 11), _c(1, 14), _c(2, 15)],
                  {"family": "Tee"}),
            _Elem(14, "Pipe", [_c(0, 3), _c(1, 4)]),
            _Elem(4, "FamilyInstance", [_c(0, 14)],
                  {"IS_GAS_FIXTURE": True, "GAS_LOAD_MBH": 100.0,
                   "FIXTURE_NAME": "RTU-1"}),
            _Elem(15, "Pipe", [_c(0, 3), _c(1, 5)]),
            _Elem(5, "FamilyInstance", [_c(0, 15), _c(1, 16)],
                  {"family": "Elbow - Generic"}),
            _Elem(16, "Pipe", [_c(0, 5), _c(1, 6)]),
            _Elem(6, "FamilyInstance", [_c(0, 16)], cap),
            _Elem(12, "Pipe", [_c(0, 2), _c(1, 7)], length=2.0),
            _Elem(7, "FamilyInstance", [_c(0, 12)],
                  {"IS_GAS_FIXTURE": True, "GAS_LOAD_MBH": 10.0,
                   "FIXTURE_NAME": "UH-2"}),
        ]
        run = self._build(els).longest_run
        self.assertEqual(run["farthest_fixture_name"], "UH-1")
        self.assertAlmostEqual(run["pipe_length_feet"], 40.0)
        self.assertEqual((run["elbow_count"], run["tee_count"]), (1, 2))
        self.assertAlmostEqual(run["elbow_equiv_length_feet"], 5.0)
        self.assertAlmostEqual(run["tee_equiv_length_feet"], 6.0)
        self.assertAlmostEqual(run["total_length_feet"], 51.0)
        self.assertEqual(run["fitting_length_feet"], 0.0)

    def test_what_counts_as_a_tee(self):
        def node(kind, fixture=False, prv=False):
            n = pipe_graph.NetworkNode(5, None, kind)
            n.is_gas_fixture = fixture
            n.is_prv = prv
            return n
        self.assertTrue(pipe_graph.counts_as_tee(node("tee"), 5, 1))
        # the walk's own start, an appliance and a PRV never count
        self.assertFalse(pipe_graph.counts_as_tee(node("tee"), 5, 5))
        self.assertFalse(pipe_graph.counts_as_tee(node("tee", fixture=True), 5, 1))
        self.assertFalse(pipe_graph.counts_as_tee(node("tee", prv=True), 5, 1))
        self.assertFalse(pipe_graph.counts_as_tee(node("fitting"), 5, 1))
        self.assertFalse(pipe_graph.counts_as_tee(node("meter"), 5, 1))

    def test_not_an_elbow(self):
        run = self._run_with_elbow("Coupling - Generic")
        self.assertEqual(run["elbow_equiv_length_feet"], 0.0)

    def test_log_digest(self):
        g = self._build(_water_heater_model({"family": "Water Heater"}))
        g.log("SKIP: Element 99 already visited.")
        helper_log = [
            {"level": "INFO", "function": "f", "element_id": 1,
             "message": "found"},
            {"level": "WARNING", "function": "get_x", "element_id": 7,
             "message": "NOT FOUND"},
        ]
        helper_log += [
            # fitting 3 without the fixture params: routine, dropped
            {"level": "WARNING", "function": "get_parameter_value",
             "element_id": 3,
             "message": "Parameter 'IS_GAS_FIXTURE' NOT FOUND on element."},
        ]
        walk, helpers = pipe_graph.log_digest(g, helper_log)
        self.assertEqual(len([s for s in walk if s.startswith("SKIP")]), 2)
        self.assertFalse(any("already visited" in s for s in walk))
        self.assertEqual(helpers, ["WARNING: get_x element 7: NOT FOUND"])


class TestZeroDemandOnCopper(unittest.TestCase):
    """A zero load appliance on a long copper run takes the smallest size at
    or above the 3/4 in. firm minimum with a published capacity.

    This test was written when the copper columns were shifted two sizes and
    3/4 K&L read blank from 550 ft. Per the 2024 IFGC PDF, 402.4(8) publishes
    3/4 K&L at every row, so the answer is 3/4 K&L (2026-11-04). No 2024
    table leaves its first size at or above 3/4 in. blank while a larger one
    has a value, so the skip-a-blank fallback is kept only as a guard."""

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
        self.assertEqual(result["sizes"][12], "3/4 K&L")
        detail = [d for d in result["segment_detail"] if d["pipe_id"] == 12][0]
        self.assertEqual(detail["capacity_mbh"],
                         gas_tables.get_capacity("402.4(8)", 595.0, "3/4 K&L"))


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
