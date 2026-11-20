# -*- coding: ascii -*-
"""Regression tests for the gas traversal review findings (2026-11-20).

Findings n = 5, 11, 20, 21, 23, 26, 27, 28, 51, 55, 56, 57. Each test
class names its finding and fails on the code before the fix.

Runs in plain CPython with the Revit assemblies stubbed (the stubs come from
_test_prv_zones, the model stand-ins from _test_gas_traversal). It does NOT
exercise a real Revit model: connectors, phases, part types, parameter
specs and System Type names are stand-ins. Live behaviour still needs a
Revit run.

Run:  python -I _test_gas_traversal_review.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_prv_zones import _node, _edge, _finish, _Id   # installs the stubs
from _test_gas_traversal import _Elem, _c, _Model          # noqa: E402

import oneline_schematic   # noqa: E402
import pipe_graph          # noqa: E402
import revit_helpers       # noqa: E402
import shared_params       # noqa: E402
import tie_in              # noqa: E402


class _ReviewModel(_Model):
    """_Model plus the reads the fixes added: the pipe System Type name, the
    parameter spec and the created phase, all from el.params."""

    def __enter__(self):
        _Model.__enter__(self)
        extra = {
            "get_piping_system_type_name":
                lambda el: el.params.get("system_name", "Natural Gas"),
            "parameter_is_plain_number":
                lambda el, name: el.params.get("load_is_number"),
            "get_created_phase_id": lambda el: el.params.get("phase"),
        }
        # getattr with a default, so the suite also runs against the code
        # before the fix (where some readers did not exist) and fails there
        # on its assertions, not on the patching.
        self._added = []
        for name, fn in extra.items():
            if hasattr(revit_helpers, name):
                self._saved[name] = getattr(revit_helpers, name)
            else:
                self._added.append(name)
            setattr(revit_helpers, name, fn)
        return self

    def __exit__(self, *args):
        _Model.__exit__(self, *args)
        for name in self._added:
            delattr(revit_helpers, name)


def _build(elements, origin_id=1, upstream=None):
    with _ReviewModel(elements) as m:
        return pipe_graph.build_network(m.elements[origin_id], None, upstream)


def _meter(to):
    return _Elem(1, "FamilyInstance", [_c(0, None, direction="In"),
                                       _c(1, to, direction="Out")],
                 {"family": "Gas Meter"})


def _rtu(eid, to, mbh=100.0, name="RTU-1", **extra):
    params = {"IS_GAS_FIXTURE": True, "GAS_LOAD_MBH": mbh,
              "FIXTURE_NAME": name, "family": "RTU"}
    params.update(extra)
    return _Elem(eid, "FamilyInstance", [_c(0, to)], params)


# ---------------------------------------------------------------------------
# n=5  pipe joined straight to pipe
# ---------------------------------------------------------------------------
class TestPipeToPipeJoint(unittest.TestCase):

    def _els(self):
        # meter 1 -> pipe 10 -> pipe 11 -> RTU 4 (100 MBH)
        return [_meter(10),
                _Elem(10, "Pipe", [_c(0, 1), _c(1, 11)]),
                _Elem(11, "Pipe", [_c(0, 10), _c(1, 4)]),
                _rtu(4, 11)]

    def test_load_reaches_both_pipes(self):
        g = _build(self._els())
        self.assertAlmostEqual(g.edges[10].cumulative_load_mbh, 100.0)
        self.assertAlmostEqual(g.edges[11].cumulative_load_mbh, 100.0)
        self.assertAlmostEqual(g.nodes[1].cumulative_load_mbh, 100.0)

    def test_longest_run_passes_the_joint(self):
        run = _build(self._els()).longest_run
        self.assertEqual(run["farthest_fixture_name"], "RTU-1")
        self.assertAlmostEqual(run["total_length_feet"], 20.0)   # no fittings

    def test_joint_node_is_never_a_fitting_to_resize(self):
        g = _build(self._els())
        joint = g.nodes[10]
        self.assertEqual(joint.node_type, "joint")
        self.assertIsNone(joint.element)
        self.assertFalse(joint.is_elbow)
        self.assertEqual(pipe_graph.non_fixture_ends(g), [])

    def test_tie_in_start_joined_to_a_pipe(self):
        els = [_Elem(900, "Pipe", [_c(0, None), _c(1, 901)]),
               _Elem(901, "Pipe", [_c(0, 900), _c(1, 4)]),
               _rtu(4, 901)]
        upstream = {"status": "ok", "reason": "it is open",
                    "upstream_xyz": [0.0, 0.0, 0.0],
                    "upstream_neighbour_id": None,
                    "downstream_neighbour_ids": [901]}
        g = _build(els, origin_id=900, upstream=upstream)
        self.assertAlmostEqual(g.edges[900].cumulative_load_mbh, 100.0)
        self.assertAlmostEqual(g.edges[901].cumulative_load_mbh, 100.0)
        self.assertAlmostEqual(g.longest_run["total_length_feet"], 20.0)


# ---------------------------------------------------------------------------
# n=21  pipe with a tap on it
# ---------------------------------------------------------------------------
def _conn(index, to, xyz, ctype="End"):
    c = _c(index, to)
    c["origin_xyz"] = xyz
    c["connector_type"] = ctype
    return c


class TestTappedPipe(unittest.TestCase):

    def _els(self):
        # meter 1 -> pipe 10 -> RTU-1 4 (100 MBH) at its end
        #            pipe 10 tap 5 -> pipe 12 -> UH-1 6 (50 MBH)
        # The tap's Curve connector is listed FIRST, so the old code took it
        # and dropped RTU-1.
        return [_meter(10),
                _Elem(10, "Pipe", [_conn(0, 1, [0.0, 0.0, 0.0]),
                                   _conn(2, 5, [4.0, 0.0, 0.0], "Curve"),
                                   _conn(1, 4, [10.0, 0.0, 0.0])]),
                _rtu(4, 10),
                _Elem(5, "FamilyInstance", [_c(0, 10), _c(1, 12)],
                      {"family": "Tap - Threaded"}),
                _Elem(12, "Pipe", [_c(0, 5), _c(1, 6)], length=30.0),
                _rtu(6, 12, mbh=50.0, name="UH-1")]

    def test_both_branches_carry_load(self):
        g = _build(self._els())
        self.assertAlmostEqual(g.nodes[1].cumulative_load_mbh, 150.0)
        self.assertAlmostEqual(g.edges[10].cumulative_load_mbh, 150.0)
        self.assertEqual(sorted(n.fixture_name for n in g.nodes.values()
                                if n.is_gas_fixture), ["RTU-1", "UH-1"])
        self.assertEqual(g.disconnected, [])

    def test_tap_is_reported(self):
        g = _build(self._els())
        self.assertEqual(g.tapped_pipes, [10])
        self.assertTrue(any(s.startswith("WARNING: Pipe 10 has a tap")
                            for s in g.traversal_log))

    def test_tap_branch_is_charged_the_whole_pipe(self):
        run = _build(self._els()).longest_run
        self.assertEqual(run["farthest_fixture_name"], "UH-1")
        self.assertAlmostEqual(run["pipe_length_feet"], 40.0)

    def test_edge_ends_are_the_pipe_ends_not_the_tap(self):
        with _ReviewModel(self._els()) as m:
            e = pipe_graph.NetworkEdge(10, m.elements[10], 1, 4)
        self.assertEqual(e.start_xyz, [0.0, 0.0, 0.0])
        self.assertEqual(e.end_xyz, [10.0, 0.0, 0.0])


# ---------------------------------------------------------------------------
# n=11  a 0 MBH fixture never sets the developed length
# ---------------------------------------------------------------------------
class TestZeroLoadFixture(unittest.TestCase):

    def _els(self, far_mbh=0.0):
        # meter 1 -10-> tee 2 -11 (10 ft)-> RTU-1 4 (100 MBH)
        #                    tee 2 -12 (100 ft)-> WH-2 6 (far_mbh)
        return [_meter(10),
                _Elem(10, "Pipe", [_c(0, 1), _c(1, 2)]),
                _Elem(2, "FamilyInstance", [_c(0, 10), _c(1, 11), _c(2, 12)],
                      {"family": "Tee"}),
                _Elem(11, "Pipe", [_c(0, 2), _c(1, 4)]),
                _rtu(4, 11),
                _Elem(12, "Pipe", [_c(0, 2), _c(1, 6)], length=100.0),
                _rtu(6, 12, mbh=far_mbh, name="WH-2")]

    def test_zero_load_does_not_set_the_run(self):
        run = _build(self._els()).longest_run
        self.assertEqual(run["farthest_fixture_name"], "RTU-1")
        self.assertAlmostEqual(run["total_length_feet"], 23.0)

    def test_zero_load_is_a_data_error(self):
        g = _build(self._els())
        self.assertEqual([(d["element_id"], d["kind"]) for d in g.data_errors],
                         [(6, "zero_load")])
        # still a terminal: nothing past it walked, not a bare end
        self.assertTrue(g.nodes[6].is_gas_fixture)
        self.assertEqual(pipe_graph.non_fixture_ends(g), [])

    def test_loaded_far_fixture_still_sets_the_run(self):
        g = _build(self._els(far_mbh=10.0))
        self.assertEqual(g.longest_run["farthest_fixture_name"], "WH-2")
        self.assertEqual(g.data_errors, [])

    def test_only_zero_load_fixtures_leave_no_run(self):
        els = [_meter(10), _Elem(10, "Pipe", [_c(0, 1), _c(1, 6)]),
               _rtu(6, 10, mbh=0.0, name="WH-2")]
        self.assertAlmostEqual(_build(els).longest_run["total_length_feet"],
                               0.0)


# ---------------------------------------------------------------------------
# n=57  GAS_LOAD_MBH read by storage type, never a crash
# ---------------------------------------------------------------------------
class TestLoadParsing(unittest.TestCase):

    def test_parse(self):
        p = pipe_graph.parse_load_mbh
        self.assertEqual(p(64.8), (64.8, None))
        self.assertEqual(p(65), (65.0, None))
        self.assertEqual(p("64.8"), (64.8, None))
        self.assertEqual(p(" 64.8 MBH "), (64.8, None))
        self.assertEqual(p("1,200"), (1200.0, None))
        self.assertEqual(p(None), (0.0, None))
        for bad in ("abc", "", True, -5.0, "nan"):
            load, problem = p(bad)
            self.assertEqual(load, 0.0, bad)
            self.assertIsNotNone(problem, bad)

    def _one(self, **params):
        els = [_meter(10), _Elem(10, "Pipe", [_c(0, 1), _c(1, 4)]),
               _rtu(4, 10, **params)]
        return _build(els)

    def test_text_load_is_read(self):
        g = self._one(mbh="64.8 MBH")
        self.assertAlmostEqual(g.nodes[4].gas_load_mbh, 64.8)
        self.assertEqual(g.data_errors, [])

    def test_bad_text_is_a_data_error_not_a_crash(self):
        g = self._one(mbh="see schedule")
        self.assertEqual(g.nodes[4].gas_load_mbh, 0.0)
        self.assertEqual([d["kind"] for d in g.data_errors], ["bad_load"])

    def test_number_with_units_is_a_data_error(self):
        g = self._one(mbh=204000.0, load_is_number=False)
        self.assertEqual(g.nodes[4].gas_load_mbh, 0.0)
        self.assertEqual([d["kind"] for d in g.data_errors], ["bad_load"])

    def test_plain_number_is_read(self):
        g = self._one(mbh=64.8, load_is_number=True)
        self.assertAlmostEqual(g.nodes[4].gas_load_mbh, 64.8)


# ---------------------------------------------------------------------------
# n=28  fixture parameters on the Type
# ---------------------------------------------------------------------------
class _Symbol(object):
    def __init__(self, eid, params):
        self.Id = _Id(eid)
        self.params = params


class TestTypeParameters(unittest.TestCase):

    def test_type_bound_fixture_is_found(self):
        rtu = _Elem(4, "FamilyInstance", [_c(0, 10)], {"family": "RTU"})
        rtu.Symbol = _Symbol(40, {"IS_GAS_FIXTURE": True,
                                  "GAS_LOAD_MBH": 120.0,
                                  "FIXTURE_NAME": "RTU-2"})
        g = _build([_meter(10), _Elem(10, "Pipe", [_c(0, 1), _c(1, 4)]),
                    rtu])
        self.assertTrue(g.nodes[4].is_gas_fixture)
        self.assertAlmostEqual(g.edges[10].cumulative_load_mbh, 120.0)
        self.assertEqual(g.nodes[4].fixture_name, "RTU-2")

    def test_instance_value_wins(self):
        rtu = _rtu(4, 10, mbh=80.0, name="RTU-3")
        rtu.Symbol = _Symbol(40, {"GAS_LOAD_MBH": 999.0,
                                  "FIXTURE_NAME": "TYPE"})
        g = _build([_meter(10), _Elem(10, "Pipe", [_c(0, 1), _c(1, 4)]),
                    rtu])
        self.assertAlmostEqual(g.nodes[4].gas_load_mbh, 80.0)
        self.assertEqual(g.nodes[4].fixture_name, "RTU-3")


# ---------------------------------------------------------------------------
# n=20  multi connector unflagged appliance at a dead end
# ---------------------------------------------------------------------------
class TestUnflaggedApplianceDeadEnd(unittest.TestCase):

    def test_two_connector_dead_end_is_listed(self):
        # meter -> pipe 10 -> elbow 3 -> pipe 11 -> RTU 5 (gas + open
        # condensate connector, no IS_GAS_FIXTURE)
        els = [_meter(10), _Elem(10, "Pipe", [_c(0, 1), _c(1, 3)]),
               _Elem(3, "FamilyInstance", [_c(0, 10), _c(1, 11)],
                     {"family": "Elbow - Generic"}),
               _Elem(11, "Pipe", [_c(0, 3), _c(1, 5)]),
               _Elem(5, "FamilyInstance", [_c(0, 11), _c(1, None)],
                     {"family": "RTU"})]
        g = _build(els)
        self.assertEqual(g.nodes[5].node_type, "fitting")
        # the elbow mid run is not an end; the RTU is
        self.assertEqual([n.element_id for n in
                          pipe_graph.non_fixture_ends(g)], [5])

    def test_existing_stop_is_not_listed(self):
        g = pipe_graph.NetworkGraph()
        g.origin_id = 1
        _node(g, 1, "meter")
        n = _node(g, 2, "fitting")
        n.connectors = [{"connected_element_id": 77}]
        _edge(g, 10, 1, 2, 5.0)
        g.existing_connections = [77]
        self.assertEqual(pipe_graph.non_fixture_ends(g), [])


# ---------------------------------------------------------------------------
# n=23  mid-stream PRV judged on the most pipe on any branch
# ---------------------------------------------------------------------------
class TestMidstreamPrvPipeLength(unittest.TestCase):

    def test_straight_branch_makes_it_midstream(self):
        # PRV 3 -1 ft-> tee 8; branch A 8 ft with 3 elbows (23 ft developed
        # incl. the tee), branch B 11 ft straight (12 ft of pipe).
        g = pipe_graph.NetworkGraph()
        g.origin_id = 1
        _node(g, 1, "meter")
        _node(g, 3, "fitting", prv=True)
        _node(g, 8, "tee")
        for eid in (20, 21, 22):
            _node(g, eid, "fitting", elbow=True)
        _node(g, 4, "fixture", 100.0, "A")
        _node(g, 5, "fixture", 100.0, "B")
        _edge(g, 100, 1, 3, 20.0)
        _edge(g, 101, 3, 8, 1.0)
        _edge(g, 102, 8, 20, 2.0)
        _edge(g, 103, 20, 21, 2.0)
        _edge(g, 104, 21, 22, 2.0)
        _edge(g, 105, 22, 4, 2.0)
        _edge(g, 106, 8, 5, 11.0)
        _finish(g)
        self.assertAlmostEqual(g.nodes[3].downstream_pipe_ft, 12.0)
        self.assertEqual(g.midstream_prv_ids, [3])


# ---------------------------------------------------------------------------
# n=26  existing appliance fed by new piping on a tie-in run
# ---------------------------------------------------------------------------
class TestExistingApplianceOnTieIn(unittest.TestCase):

    def _run(self, appliance_params):
        els = [_Elem(900, "Pipe", [_c(0, None), _c(1, 4)], {"phase": 2}),
               _Elem(4, "FamilyInstance", [_c(0, 900)], appliance_params)]
        upstream = {"status": "ok", "reason": "it is open",
                    "upstream_xyz": None, "upstream_neighbour_id": None,
                    "downstream_neighbour_ids": [4]}
        return _build(els, origin_id=900, upstream=upstream)

    def test_existing_appliance_load_is_counted(self):
        g = self._run({"IS_GAS_FIXTURE": True, "GAS_LOAD_MBH": 40.0,
                       "FIXTURE_NAME": "WH-1", "phase": 1})
        self.assertAlmostEqual(g.edges[900].cumulative_load_mbh, 40.0)
        self.assertEqual(g.existing_appliances, [4])
        self.assertEqual(g.existing_connections, [])
        self.assertEqual(g.longest_run["farthest_fixture_name"], "WH-1")

    def test_existing_fitting_is_still_not_followed(self):
        g = self._run({"family": "Tee", "phase": 1})
        self.assertEqual(g.existing_connections, [4])
        self.assertNotIn(4, g.nodes)


# ---------------------------------------------------------------------------
# n=27  new tee on a modeled existing main
# ---------------------------------------------------------------------------
class _Category(object):
    Name = "Pipe Fittings"


class _Part(_Elem):
    Category = _Category()


class TestTeeOnExistingMain(unittest.TestCase):

    def _resolve(self, branch_phase):
        # picked pipe 900 (phase 2): end 0 -> new tee 50, end 1 -> pipe 901
        # tee 50 (phase 2): 900, existing pipes 60 and 61 (phase
        # branch_phase for 61, phase 1 for 60)
        els = {
            900: _Elem(900, "Pipe", [_c(0, 50), _c(1, 901)], {"phase": 2}),
            901: _Elem(901, "Pipe", [_c(0, 900), _c(1, None)], {"phase": 2}),
            50: _Part(50, "FamilyInstance", [_c(0, 900), _c(1, 60), _c(2, 61)],
                      {"family": "Tee", "phase": 2}),
            60: _Elem(60, "Pipe", [_c(0, 50)], {"phase": 1}),
            61: _Elem(61, "Pipe", [_c(0, 50)], {"phase": branch_phase}),
        }
        patches = {
            "get_connectors": lambda el: list(el.connectors),
            "get_element_by_id": lambda doc, i: els.get(i),
            "get_created_phase_id": lambda el: el.params.get("phase"),
            "get_parameter_value": lambda el, name: el.params.get(name),
        }
        saved = dict((k, getattr(revit_helpers, k)) for k in patches)
        try:
            for k, fn in patches.items():
                setattr(revit_helpers, k, fn)
            return tie_in.resolve_new_pipe_upstream(els[900], None)
        finally:
            for k, fn in saved.items():
                setattr(revit_helpers, k, fn)

    def test_tee_on_existing_main_is_the_upstream_end(self):
        r = self._resolve(branch_phase=1)
        self.assertEqual(r["status"], "ok")
        self.assertEqual(r["upstream_neighbour_id"], 50)
        self.assertEqual(r["downstream_neighbour_ids"], [901])
        self.assertIn("new tee on the existing main", r["reason"])

    def test_tee_with_new_branch_is_still_new(self):
        r = self._resolve(branch_phase=2)
        self.assertEqual(r["status"], "no_upstream")


# ---------------------------------------------------------------------------
# n=51  pressure reducing valve detected as a PRV
# ---------------------------------------------------------------------------
class TestPressureReducingValve(unittest.TestCase):

    def test_detected_in_the_walk_and_drawn_as_prv(self):
        els = [_meter(10), _Elem(10, "Pipe", [_c(0, 1), _c(1, 3)]),
               _Elem(3, "FamilyInstance", [_c(0, 10), _c(1, 11)],
                     {"family": "Gas Pressure Reducing Valve"}),
               _Elem(11, "Pipe", [_c(0, 3), _c(1, 4)], length=30.0),
               _rtu(4, 11)]
        g = _build(els)
        self.assertEqual(g.prv_ids, [3])
        self.assertEqual(g.midstream_prv_ids, [3])
        self.assertEqual(oneline_schematic._valve_kind(
            {"kind": "fitting", "family": "Gas Pressure Reducing Valve",
             "is_prv": g.nodes[3].is_prv}), "prv")


# ---------------------------------------------------------------------------
# n=55  gas only when the System Type name says gas
# ---------------------------------------------------------------------------
class TestGasSystemTypeName(unittest.TestCase):

    def _check(self, name, classification="OtherPipe"):
        pipe = _Elem(10, "Pipe", [_c(0, None, classification),
                                  _c(1, None, classification)],
                     {"system_name": name})
        with _ReviewModel([pipe]):
            return pipe_graph.check_gas_pipe(pipe)

    def test_compressed_air_pick_is_rejected(self):
        r = self._check("Compressed Air")
        self.assertFalse(r["ok"])
        self.assertIn("Compressed Air", r["reason"])

    def test_gas_names_pass(self):
        # Colin 2026-11-20: "gas" or "propane" anywhere, or "LP" as a word.
        for name in ("Natural Gas", "LP Gas", "LP", "Propane",
                     "GAS - MEDIUM PRESSURE", "lp-2psi", "PROPANE 11 IN WC"):
            self.assertTrue(self._check(name)["ok"], name)
            self.assertIsNone(pipe_graph.not_gas_system_name(name), name)

    def test_other_names_fail(self):
        for name in ("Condensate", "Pumped Discharge", "Help",
                     "Compressed Air", "Fuel Oil", "Helps"):
            r = self._check(name)
            self.assertFalse(r["ok"], name)
            self.assertIn("'{}'".format(name), r["reason"])
            self.assertIn("'gas' or 'propane' or the word 'LP'", r["reason"])
            self.assertIsNotNone(pipe_graph.not_gas_system_name(name), name)

    def test_unreadable_name_is_rejected_at_the_pick(self):
        self.assertFalse(self._check(None)["ok"])

    def test_walk_stops_at_condensate(self):
        # unflagged appliance 2 with gas in and an OtherPipe condensate pipe
        els = [_meter(10), _Elem(10, "Pipe", [_c(0, 1), _c(1, 2)]),
               _Elem(2, "FamilyInstance", [_c(0, 10), _c(1, 20)],
                     {"family": "Furnace"}),
               _Elem(20, "Pipe", [_c(0, 2), _c(1, None)],
                     {"system_name": "Condensate"})]
        g = _build(els)
        self.assertEqual(sorted(g.edges.keys()), [10])
        self.assertEqual(g.non_gas_skipped, [(20, "Condensate")])
        self.assertNotIn(20, g.disconnected)


# ---------------------------------------------------------------------------
# n=56  elbows by part type, name keywords as the fallback
# ---------------------------------------------------------------------------
class _MEPModel(object):
    def __init__(self, part):
        self.PartType = part


class _Fitting(object):
    def __init__(self, eid, part):
        self.Id = _Id(eid)
        self.MEPModel = _MEPModel(part)


class TestElbowDetection(unittest.TestCase):

    def test_names(self):
        for name in ("90 Ell - Threaded", "Pipe Bend", "Elbow - Generic",
                     "45 ELL", "ell-90"):
            self.assertTrue(revit_helpers.elbow_by_name(name), name)
        for name in ("Shell Coupling", "Coupling - Generic", "Tee", "",
                     None, "Bellows"):
            self.assertFalse(revit_helpers.elbow_by_name(name), name)

    def test_part_type_wins(self):
        self.assertTrue(revit_helpers.is_elbow(_Fitting(5, "Elbow")))
        self.assertFalse(revit_helpers.is_elbow(_Fitting(5, "Tee")))

    def test_ell_family_adds_five_feet_in_the_walk(self):
        els = [_meter(10), _Elem(10, "Pipe", [_c(0, 1), _c(1, 3)]),
               _Elem(3, "FamilyInstance", [_c(0, 10), _c(1, 11)],
                     {"family": "90 Ell - Threaded"}),
               _Elem(11, "Pipe", [_c(0, 3), _c(1, 4)]),
               _rtu(4, 11)]
        run = _build(els).longest_run
        self.assertEqual(run["elbow_count"], 1)
        self.assertAlmostEqual(run["total_length_feet"], 25.0)


# ---------------------------------------------------------------------------
# n=30  only piping connectors set the node type and are walked
# ---------------------------------------------------------------------------
def _dom(c, domain):
    c["domain"] = domain
    return c


class _Origin(object):
    X = Y = Z = 0.0


class _FakeConnector(object):
    def __init__(self, index, domain):
        self.Id = index
        self.ConnectorType = "End"
        self.Direction = "Bidirectional"
        self.IsConnected = False
        self.PipeSystemType = "OtherPipe"
        self.Radius = 0.0
        self.Origin = _Origin()
        self._domain = domain

    @property
    def Domain(self):
        if self._domain is None:
            raise AttributeError("no Domain")
        return self._domain


class _Manager(object):
    def __init__(self, connectors):
        self.Connectors = connectors


class _ConnElem(object):
    def __init__(self, eid, connectors):
        self.Id = _Id(eid)
        self.ConnectorManager = _Manager(connectors)


class TestNonPipingConnectors(unittest.TestCase):

    def _valve_model(self, third_domain):
        # meter 1 -> pipe 10 -> Ansul valve 3 -> pipe 11 -> RTU 4, the valve
        # has a third, unconnected connector of domain third_domain
        return [_meter(10), _Elem(10, "Pipe", [_c(0, 1), _c(1, 3)]),
                _Elem(3, "FamilyInstance",
                      [_c(0, 10), _c(1, 11), _dom(_c(2, None), third_domain)],
                      {"family": "RJA - Gas Shutoff Valve (Ansul)"}),
                _Elem(11, "Pipe", [_c(0, 3), _c(1, 4)]),
                _rtu(4, 11)]

    def test_electrical_connector_does_not_make_a_tee(self):
        g = _build(self._valve_model("DomainElectrical"))
        node = g.nodes[3]
        self.assertEqual(node.node_type, "fitting")
        self.assertEqual(node.connector_count, 2)
        self.assertEqual(len(node.connectors), 2)
        # no false 3 ft tee equivalent: 20 ft of pipe only
        self.assertAlmostEqual(g.longest_run["total_length_feet"], 20.0)
        self.assertEqual(g.longest_run["tee_count"], 0)
        self.assertEqual(pipe_graph.non_fixture_ends(g), [])
        self.assertEqual(oneline_schematic._valve_kind(
            {"kind": node.node_type, "family": node.family_name,
             "is_prv": node.is_prv}), "ansul")

    def test_unreadable_domain_still_counts(self):
        g = _build(self._valve_model(None))
        self.assertEqual(g.nodes[3].node_type, "tee")
        self.assertEqual(g.nodes[3].connector_count, 3)

    def test_piping_connectors_filter(self):
        cons = [_dom(_c(i, None), d) for i, d in enumerate(
            ("DomainPiping", "DomainElectrical", None, "DomainHvac",
             "DomainUndefined", "DomainCableTrayConduit"))]
        kept = revit_helpers.piping_connectors(cons, 3)
        self.assertEqual([c["connector_index"] for c in kept], [0, 2, 4])

    def test_get_connectors_reads_the_domain(self):
        el = _ConnElem(3, [_FakeConnector(0, "DomainPiping"),
                           _FakeConnector(1, "DomainElectrical"),
                           _FakeConnector(2, None)])
        cons = revit_helpers.get_connectors(el)
        self.assertEqual([c["domain"] for c in cons],
                         ["DomainPiping", "DomainElectrical", None])
        self.assertEqual(len(revit_helpers.piping_connectors(cons)), 2)


# ---------------------------------------------------------------------------
# 2026-11-20 cleanup, item F: one accessory classifier
# ---------------------------------------------------------------------------
class TestAccessoryKind(unittest.TestCase):

    def test_classifier_order_and_words(self):
        kind = shared_params.accessory_kind
        self.assertEqual(kind("Gas Pressure Reducing Valve"), "prv")
        self.assertEqual(kind("Maxitrol Regulator"), "prv")
        self.assertEqual(kind("RJA - Gas Shutoff Valve (Ansul)"), "ansul")
        self.assertEqual(kind("Ball Valve - Threaded"), "iso")
        self.assertIsNone(kind("Tee - Threaded"))
        self.assertIsNone(kind("Reducing Tee"))
        self.assertIsNone(kind(None))
        self.assertIsNone(kind(""))
        for kw in shared_params.PRV_FAMILY_KEYWORDS:
            self.assertEqual(kind("X " + kw.upper()), "prv", kw)

    def test_keywords_text_lists_every_prv_word(self):
        text = shared_params.prv_keywords_text()
        for kw in shared_params.PRV_FAMILY_KEYWORDS:
            self.assertIn("'{}'".format(kw), text)
        self.assertIn(" or ", text)

    def test_walk_and_one_line_use_it(self):
        here = os.path.dirname(os.path.abspath(__file__))
        for name in ("pipe_graph.py", "oneline_schematic.py"):
            src = open(os.path.join(here, name)).read()
            self.assertIn("shared_params.accessory_kind(", src, name)
            self.assertNotIn("in shared_params.PRV_FAMILY_KEYWORDS", src,
                             name)
            self.assertNotIn("in shared_params.ISOLATION_VALVE_KEYWORDS",
                             src, name)
            self.assertNotIn("in shared_params.ANSUL_VALVE_KEYWORDS", src,
                             name)

    def test_valve_kind_follows_the_classifier(self):
        vk = oneline_schematic._valve_kind
        self.assertEqual(vk({"kind": "fitting", "family": "Ball Valve"}),
                         "iso")
        self.assertEqual(vk({"kind": "fitting", "family": "Ansul Valve"}),
                         "ansul")
        # a PRV name the walk did not flag gets no invented symbol
        self.assertIsNone(vk({"kind": "fitting", "family": "Gas Regulator"}))
        self.assertEqual(vk({"kind": "fitting", "family": "Gas Regulator",
                             "is_prv": True}), "prv")
        self.assertIsNone(vk({"kind": "fitting", "family": "Ball Valve",
                              "is_elbow": True}))


# ---------------------------------------------------------------------------
# 2026-11-20 cleanup, item G: piping connectors only in the gas paths
# ---------------------------------------------------------------------------
class _DirConnector(_FakeConnector):
    def __init__(self, index, domain, direction="Out", connected=True):
        _FakeConnector.__init__(self, index, domain)
        self.IsConnected = connected
        self.Direction = direction
        self.CoordinateSystem = type("CS", (object,), {"BasisZ": None})()


class TestPipingConnectorsOnly(unittest.TestCase):

    def test_get_piping_connectors(self):
        el = _ConnElem(3, [_FakeConnector(0, "DomainPiping"),
                           _FakeConnector(1, "DomainElectrical"),
                           _FakeConnector(2, None)])
        self.assertEqual(
            [c["connector_index"]
             for c in revit_helpers.get_piping_connectors(el)], [0, 2])

    def test_meter_validation_ignores_an_electrical_out(self):
        el = _ConnElem(1, [])
        saved = revit_helpers.get_connectors
        try:
            revit_helpers.get_connectors = lambda e: [
                {"connector_index": 0, "direction": "In",
                 "is_connected": False, "domain": "DomainPiping",
                 "connected_element_id": None},
                {"connector_index": 1, "direction": "Out",
                 "is_connected": True, "domain": "DomainElectrical",
                 "connected_element_id": 9}]
            v = revit_helpers.validate_selected_element(el)
            self.assertFalse(v["is_valid"])
            self.assertEqual(v["connector_summary"], ["In"])
            revit_helpers.get_connectors = lambda e: [
                {"connector_index": 0, "direction": "Out",
                 "is_connected": True, "domain": "DomainPiping",
                 "connected_element_id": 9},
                {"connector_index": 1, "direction": "In",
                 "is_connected": False, "domain": "DomainElectrical",
                 "connected_element_id": None}]
            v = revit_helpers.validate_selected_element(el)
            self.assertTrue(v["is_valid"])
            self.assertEqual(v["connector_summary"], ["Out"])
        finally:
            revit_helpers.get_connectors = saved

    def test_is_piping_connector(self):
        self.assertTrue(revit_helpers._is_piping_connector(
            _FakeConnector(0, "DomainPiping")))
        self.assertFalse(revit_helpers._is_piping_connector(
            _FakeConnector(0, "DomainElectrical")))
        self.assertTrue(revit_helpers._is_piping_connector(
            _FakeConnector(0, None)))        # unreadable counts as piping

    def _reconnect(self, piping_only):
        """reconnect_open_ends on a valve whose only open connector is
        electrical. Outline is the first Revit call made for an open
        connector, so it records whether the connector was searched."""
        db = sys.modules["Autodesk.Revit.DB"]
        reached = []
        names = ("BoundingBoxIntersectsFilter", "Line", "Outline", "XYZ")
        saved = dict((n, getattr(db, n, None)) for n in names)

        def _outline(*a):
            reached.append(a)
            raise RuntimeError("searched")
        try:
            for n in names:
                setattr(db, n, lambda *a: None)
            db.Outline = _outline
            el = _ConnElem(3, [_DirConnector(0, "DomainPiping"),
                               _DirConnector(1, "DomainElectrical",
                                             connected=False)])
            try:
                joined = revit_helpers.reconnect_open_ends(
                    None, el, system_type_ids=[7], piping_only=piping_only)
            except RuntimeError:
                joined = None
        finally:
            for n, v in saved.items():
                if v is None:
                    delattr(db, n)
                else:
                    setattr(db, n, v)
        return joined, reached

    def test_reconnect_skips_an_electrical_connector_for_gas(self):
        joined, reached = self._reconnect(True)
        self.assertEqual(joined, 0)
        self.assertEqual(reached, [])

    def test_reconnect_unchanged_for_water(self):
        # the default (water) still searches every open connector
        _joined, reached = self._reconnect(False)
        self.assertEqual(len(reached), 1)

    def test_only_size_gas_asks_for_piping_only(self):
        here = os.path.dirname(os.path.abspath(__file__))
        water = open(os.path.join(here, "water_sizing_engine.py")).read()
        self.assertNotIn("piping_only", water)
        # headless branch: the Size Gas pipeline lives in lib/gas_run.py
        size_gas = open(os.path.join(here, "gas_run.py")).read()
        self.assertIn("piping_only=True", size_gas)
        rg = open(os.path.join(here, "report_generator.py")).read()
        self.assertIn("revit_helpers.get_piping_connectors(origin_element)",
                      rg)


if __name__ == "__main__":
    unittest.main()
