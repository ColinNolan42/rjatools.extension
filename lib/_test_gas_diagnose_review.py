# -*- coding: ascii -*-
"""Regression tests for the 2026-11-20 gas review fixes on the Diagnose side
(report_generator and the Diagnose gas branch): n=10/11 (0 MBH data
errors), n=33/41 (a pipe pick is the tie-in pipe, checked as Size Gas checks
it), n=36 (fitting counts), n=42 (nested PRVs agree with ready_for_sizing),
n=62 (zone runs), n=63 (markdown escaping), plus the walk's joints and
tapped pipes.

Plain CPython with the Revit assemblies stubbed. The Diagnose pushbutton
itself cannot run here (it calls main() and needs pyRevit), so its wiring is
checked by reading its source; the logic it calls lives in report_generator
and is tested directly. Nothing here proves behavior inside Revit.

Run:  python _test_gas_diagnose_review.py
"""

import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from _test_prv_zones import _install_revit_stubs   # noqa: E402,F401

import pipe_graph          # noqa: E402
import report_generator as rg   # noqa: E402
import revit_helpers       # noqa: E402
import tie_in              # noqa: E402

DIAGNOSE = os.path.join(HERE, "..", "Developer Tools.tab", "Performance.panel",
                        "Diagnose.pushbutton", "script.py")


class _Node(object):
    def __init__(self, eid, node_type, family="Family", fixture=False,
                 mbh=0.0, name=None, midstream=False, downstream_ft=0.0):
        self.element_id = eid
        self.node_type = node_type
        self.family_name = family
        self.is_gas_fixture = fixture
        self.gas_load_mbh = mbh
        self.fixture_name = name
        self.location_xyz = None
        self.connector_count = 2
        self.cumulative_load_mbh = mbh
        self.is_midstream_prv = midstream
        self.downstream_pipe_ft = downstream_ft
        self.connectors = []


class _Graph(object):
    def __init__(self, nodes=(), tie_in_origin=False):
        self.nodes = dict((n.element_id, n) for n in nodes)
        self.edges = {}
        self.disconnected = []
        self.longest_run = {"total_length_feet": 100.0,
                            "pipe_length_feet": 90.0, "elbow_count": 2,
                            "elbow_equiv_length_feet": 10.0,
                            "tee_count": 0, "tee_equiv_length_feet": 0.0,
                            "farthest_fixture_name": "RTU-1",
                            "farthest_fixture_id": 10,
                            "path_element_ids": [1, 10]}
        self.nested_prv_ids = []
        self.zone_runs = {}
        self.midstream_prv_ids = []
        self.origin_id = 1
        self.origin_is_tie_in = tie_in_origin
        self.data_errors = []
        self.tapped_pipes = []
        self.existing_appliances = []


def _meter_conns(out=True, connected=True):
    return [{"direction": "Out" if out else "Bidirectional",
             "is_connected": connected}]


def _loaded(*extra):
    nodes = [_Node(1, "meter"),
             _Node(10, "fixture", fixture=True, mbh=100.0, name="RTU-1")]
    return list(nodes) + list(extra)


class ZeroLoadDataErrors(unittest.TestCase):
    """n=10/11: every 0 MBH fixture is listed as a data error (id and name);
    a total of 0 MBH stays a blocking error."""

    def test_walk_data_errors_are_listed_with_id_and_name(self):
        g = _Graph(_loaded(_Node(11, "fixture", fixture=True, mbh=0.0,
                                 name="WH-1"),
                           _Node(12, "fixture", fixture=True, mbh=0.0,
                                 name="GUH-2")))
        g.data_errors = [
            {"element_id": 11, "name": "WH-1", "kind": "zero_load",
             "message": "GAS_LOAD_MBH is 0"},
            {"element_id": 12, "name": "GUH-2", "kind": "bad_load",
             "message": "GAS_LOAD_MBH text 'abc' is not a number"}]
        v = rg.gas_validation(g, _meter_conns())
        self.assertEqual(len(v["data_errors"]), 2)
        self.assertIn("WH-1 (element 11)", v["data_errors"][0])
        self.assertIn("GUH-2 (element 12)", v["data_errors"][1])
        self.assertIn("not a number", v["data_errors"][1])
        # the walk skips them, so a loaded system still sizes
        self.assertTrue(v["ready_for_sizing"])
        self.assertEqual(v["error_messages"], [])

    def test_walk_record_is_the_one_source(self):
        # item B (2026-11-20): pipe_graph records every 0 MBH fixture in
        # graph.data_errors; the report does not test the load again
        g = _Graph(_loaded(_Node(11, "fixture", fixture=True, mbh=0.0,
                                 name="WH-1")))
        v = rg.gas_validation(g, _meter_conns())
        self.assertEqual(v["data_errors"], [])
        del g.data_errors
        self.assertEqual(rg.gas_validation(g, _meter_conns())["data_errors"],
                         [])
        self.assertFalse(hasattr(rg, "zero_load_fixtures"))
        g = _Graph(_loaded(_Node(11, "fixture", fixture=True, mbh=0.0,
                                 name="WH-1")))
        g.data_errors = [{"element_id": 11, "name": "WH-1",
                          "kind": "zero_load", "message": "GAS_LOAD_MBH is 0"}]
        self.assertEqual(rg.gas_validation(g, _meter_conns())["data_errors"],
                         ["WH-1 (element 11): GAS_LOAD_MBH is 0."])

    def test_total_zero_stays_blocking(self):
        g = _Graph([_Node(1, "meter"),
                    _Node(11, "fixture", fixture=True, mbh=0.0, name="A")])
        g.data_errors = [{"element_id": 11, "name": "A", "kind": "zero_load",
                          "message": "GAS_LOAD_MBH is 0"}]
        v = rg.gas_validation(g, _meter_conns())
        self.assertFalse(v["ready_for_sizing"])
        self.assertIn("Total load is 0 MBH.", v["error_messages"])
        self.assertEqual(len(v["data_errors"]), 1)

    def test_paste_block_lists_data_errors(self):
        g = _Graph(_loaded(_Node(11, "fixture", fixture=True, mbh=0.0,
                                 name="WH-1")))
        g.data_errors = [{"element_id": 11, "name": "WH-1",
                          "kind": "zero_load", "message": "GAS_LOAD_MBH is 0"}]
        text = _format(g)
        self.assertIn("DATA ERROR  WH-1 (element 11)", text)
        self.assertIn("DATA-ERROR-load=0", text)


class _Elem(object):
    class _Id(object):
        IntegerValue = 5

    Id = _Id()


def _format(graph, **kw):
    saved = (revit_helpers.get_connectors, revit_helpers.get_element_location,
             revit_helpers.get_element_family_name, revit_helpers.eid_int)
    revit_helpers.get_connectors = lambda e: _meter_conns()
    revit_helpers.get_element_location = lambda e: None
    revit_helpers.get_element_family_name = lambda e: "Pipe Types"
    revit_helpers.eid_int = lambda i: 5
    try:
        return rg.format_diagnostic_output(graph, _Elem(), **kw)
    finally:
        (revit_helpers.get_connectors, revit_helpers.get_element_location,
         revit_helpers.get_element_family_name, revit_helpers.eid_int) = saved


class TieInPipePick(unittest.TestCase):
    """n=33/41: a pipe pick runs check_gas_pipe and
    resolve_new_pipe_upstream (imported), skips the meter checks and is
    labelled the tie-in pipe."""

    def setUp(self):
        self.calls = []
        self._check = pipe_graph.check_gas_pipe
        self._resolve = tie_in.resolve_new_pipe_upstream
        self._is_pipe = revit_helpers.is_pipe
        self._validate = revit_helpers.validate_selected_element
        revit_helpers.is_pipe = lambda e: True

    def tearDown(self):
        pipe_graph.check_gas_pipe = self._check
        tie_in.resolve_new_pipe_upstream = self._resolve
        revit_helpers.is_pipe = self._is_pipe
        revit_helpers.validate_selected_element = self._validate

    def _patch(self, gas_ok, upstream):
        calls = self.calls

        def _check(pipe):
            calls.append("check")
            return {"ok": gas_ok, "reason": "the picked pipe is on "
                    "'Hydronic Supply', which is not a gas piping system"}

        def _resolve(pipe, doc, system=None):
            calls.append("resolve")
            return upstream
        pipe_graph.check_gas_pipe = _check
        tie_in.resolve_new_pipe_upstream = _resolve

    def test_non_gas_pipe_is_rejected_like_size_gas(self):
        self._patch(False, {"status": "ok"})
        res = tie_in.check_gas_start(object(), None)
        self.assertFalse(res["ok"])
        self.assertIn("This is not a gas pipe", res["message"])
        self.assertIn(tie_in.GAS_START_HINT, res["message"])
        self.assertIn("Hydronic Supply", res["message"])
        self.assertEqual(self.calls, ["check"])

    def test_ambiguous_upstream_is_rejected(self):
        self._patch(True, {"status": "ambiguous", "reason": "more than one "
                           "end is open"})
        res = tie_in.check_gas_start(object(), None)
        self.assertFalse(res["ok"])
        self.assertEqual(res["kind"], "tie_in")
        self.assertIn("more than one end is open", res["message"])
        self.assertIn("Pick the first new gas pipe", res["message"])

    def test_ok_passes_upstream_and_tee_on_existing_reason(self):
        reason = tie_in.pick_upstream_end(
            [{"connector_index": 0, "kind": tie_in.END_TEE_ON_EXISTING},
             {"connector_index": 1, "kind": tie_in.END_NEW}])["reason"]
        up = {"status": "ok", "reason": reason}
        self._patch(True, up)
        res = tie_in.check_gas_start(object(), None)
        self.assertTrue(res["ok"])
        self.assertEqual(res["kind"], "tie_in")
        self.assertIs(res["upstream"], up)
        self.assertIn("new tee on the existing main", res["message"])
        self.assertEqual(self.calls, ["check", "resolve"])

    def test_meter_pick_is_validated_as_the_meter(self):
        # item A: anything not a pipe is checked as the gas meter
        revit_helpers.is_pipe = lambda e: False
        revit_helpers.validate_selected_element = lambda e: {
            "is_valid": False, "reason": "no Out connector."}
        res = tie_in.check_gas_start(object(), None)
        self.assertFalse(res["ok"])
        self.assertEqual(res["kind"], "meter")
        self.assertIn(tie_in.GAS_START_HINT, res["message"])
        self.assertIn("no Out connector.", res["message"])
        revit_helpers.validate_selected_element = lambda e: {
            "is_valid": True, "reason": ""}
        res = tie_in.check_gas_start(object(), None)
        self.assertEqual((res["ok"], res["kind"], res["upstream"]),
                         (True, "meter", None))
        self.assertEqual(self.calls, [])

    def test_connect_into_existing_takes_only_a_pipe(self):
        revit_helpers.is_pipe = lambda e: False
        res = tie_in.check_gas_start(object(), None, pipe_only=True)
        self.assertFalse(res["ok"])
        self.assertIsNone(res["kind"])
        self.assertIn("Connect into Existing was chosen", res["message"])

    def test_meter_checks_skipped_on_a_tie_in(self):
        g = _Graph(_loaded(), tie_in_origin=True)
        v = rg.gas_validation(g, None)
        names = [c["check"] for c in v["checks"]]
        self.assertFalse([n for n in names if "Meter" in n])
        self.assertTrue(v["ready_for_sizing"])
        # a meter pick with no Out connector still fails, as Size Gas does
        v = rg.gas_validation(_Graph(_loaded()), _meter_conns(out=False))
        self.assertFalse(v["ready_for_sizing"])

    def test_paste_block_labels_the_tie_in_pipe(self):
        g = _Graph(_loaded(), tie_in_origin=True)
        text = _format(g, tie_in_check={
            "ok": True, "message": "New gas piping connection. Upstream end "
            "found because it is open."})
        self.assertIn("TIE-IN PIPE", text)
        self.assertNotIn("Meter validation", text)
        self.assertNotIn("Meter has Out connector", text)
        self.assertIn("Tie-in pipe validation: PASS", text)
        self.assertIn("ready_for_sizing: True", text)
        text = _format(_Graph(_loaded()))
        self.assertIn("Meter validation: PASS", text)


class NestedPrvAgreement(unittest.TestCase):
    """n=42: nested mid-stream PRVs block only with Mid Stream PRV ticked,
    and ready_for_sizing matches the summary's errors."""

    def _graph(self):
        g = _Graph(_loaded())
        g.nested_prv_ids = [77]
        return g

    def test_diagnose_default_is_a_warning(self):
        v = rg.gas_validation(self._graph(), _meter_conns())
        self.assertTrue(v["ready_for_sizing"])
        self.assertEqual(v["error_messages"], [])
        self.assertTrue([w for w in v["warning_messages"] if "77" in w])

    def test_ticked_is_an_error(self):
        v = rg.gas_validation(self._graph(), _meter_conns(),
                              mid_stream_prv=True)
        self.assertFalse(v["ready_for_sizing"])
        self.assertTrue([e for e in v["error_messages"] if "77" in e])

    def test_ready_always_matches_the_summary_errors(self):
        for g in (self._graph(), _Graph([_Node(1, "meter")]),
                  _Graph(_loaded())):
            for tick in (None, True):
                v = rg.gas_validation(g, _meter_conns(), tick)
                self.assertEqual(v["ready_for_sizing"],
                                 not v["error_messages"])
                text = _format(g, mid_stream_prv=tick)
                self.assertIn("ready_for_sizing: {}".format(
                    v["ready_for_sizing"]), text)


class FittingCounts(unittest.TestCase):
    """n=36 and joints: an accessory is counted once, never subtracted from
    another count; joints (no element) are not fittings."""

    @staticmethod
    def _kind(n):
        return "PRV" if "Regulator" in n.family_name else None

    def test_tee_type_regulator_is_not_double_counted(self):
        g = _Graph([_Node(1, "meter"),
                    _Node(2, "tee", family="Gas Regulator (vent)"),
                    _Node(3, "fitting", family="Elbow"),
                    _Node(4, "tee", family="Tee")])
        c = rg.gas_summary_counts(g, self._kind)
        self.assertEqual((c["tees"], c["fittings_other"], c["prv_count"]),
                         (1, 1, 1))

    def test_never_negative(self):
        g = _Graph([_Node(1, "meter"),
                    _Node(2, "tee", family="Gas Regulator (vent)")])
        c = rg.gas_summary_counts(g, self._kind)
        self.assertEqual((c["tees"], c["fittings_other"]), (0, 0))

    def test_joints_are_not_fittings(self):
        g = _Graph(_loaded(_Node(30, "joint", family="PIPE JOINT"),
                           _Node(3, "fitting", family="Elbow")))
        c = rg.gas_summary_counts(g, self._kind)
        self.assertEqual(c["fittings_other"], 1)
        text = _format(g)
        self.assertIn("=== FITTINGS (1 found) ===", text)
        # listed in the network graph (debug), never as a fitting
        self.assertNotIn("[30]  joint  connectors", text)
        self.assertIn("node [30]  joint", text)


class TappedAndExisting(unittest.TestCase):

    def test_tapped_pipe_is_a_warning(self):
        g = _Graph(_loaded())
        g.tapped_pipes = [501]
        v = rg.gas_validation(g, _meter_conns())
        self.assertTrue(v["ready_for_sizing"])
        self.assertTrue([w for w in v["warning_messages"] if "501" in w])
        self.assertIn("TAPPED PIPES", _format(g))

    def test_existing_appliances_listed(self):
        g = _Graph(_loaded(_Node(40, "fixture", fixture=True, mbh=50.0,
                                 name="WH-EX")))
        g.existing_appliances = [40]
        self.assertEqual(rg.existing_appliance_rows(g), [(40, "WH-EX")])
        self.assertIn("[40]  WH-EX", _format(g))


class ZoneRuns(unittest.TestCase):
    """n=62: each zone's developed length is shown with a mid-stream PRV."""

    def test_zone_rows(self):
        g = _Graph(_loaded())
        self.assertEqual(rg.zone_run_rows(g), [])
        g.midstream_prv_ids = [9]
        g.zone_runs = {1: {"total_length_feet": 60.0,
                           "farthest_fixture_name": "PRV-1"},
                       9: {"total_length_feet": 180.0,
                           "farthest_fixture_name": "RTU-2"}}
        rows = rg.zone_run_rows(g)
        self.assertEqual(len(rows), 2)
        self.assertIn("Meter zone", rows[0][0])
        self.assertIn("After PRV 9", rows[1][0])
        # item J: Diagnose does not know the box, so the rows say when
        for label, _run in rows:
            self.assertIn("(used when Mid Stream PRV is ticked)", label)
            self.assertNotIn("(Mid Stream PRV ticked)", label)
        self.assertEqual(rg.zone_run_text(rows[1][1]), "180.0 ft to RTU-2")
        text = _format(g)
        self.assertIn("60.0 ft to PRV-1", text)
        self.assertIn("180.0 ft to RTU-2", text)


class MarkdownEscape(unittest.TestCase):
    """n=63: model text cannot break or hide a print_md table cell."""

    def test_md_cell(self):
        self.assertEqual(rg.md_cell("WH-1 <TANK>"), "WH-1 &lt;TANK&gt;")
        cell = rg.md_cell("Valve | Gas *main* _x_")
        self.assertNotIn("|", cell)
        self.assertNotIn("*", cell)
        self.assertNotIn("_", cell)
        self.assertEqual(rg.md_cell(None), "")
        self.assertEqual(rg.md_cell(12), "12")
        self.assertEqual(rg.md_cell("a\nb"), "a b")


class DiagnoseWiring(unittest.TestCase):
    """The Diagnose gas branch calls the shared, tested functions."""

    @classmethod
    def setUpClass(cls):
        f = open(DIAGNOSE, "rb")
        try:
            cls.src = f.read().decode("utf-8")
        finally:
            f.close()
        start = cls.src.index("def _gas_check_pipe_pick(")
        cls.gas = cls.src[start:cls.src.index("def diagnose_water(")]

    def test_uses_the_shared_checks(self):
        # item F: Diagnose counts valves with the one classifier
        self.assertIn("shared_params.accessory_kind(", self.src)
        self.assertNotIn("_GAS_ISO_KW", self.src)
        for call in ("tie_in.check_gas_start(",
                     "pipe_graph.total_connected_mbh(graph)",
                     "revit_helpers.get_piping_connectors(element)",
                     "report_generator.gas_summary_counts(",
                     "report_generator.gas_validation(",
                     "report_generator.zone_run_rows(",
                     "report_generator.existing_appliance_rows(",
                     "pipe_graph.build_network(element, doc, upstream)",
                     "tie_in_check=tie_in_check"):
            self.assertIn(call, self.gas)

    def test_old_bugs_are_gone(self):
        self.assertNotIn("len(fittings) - len(accessories)", self.gas)
        self.assertNotIn("if graph.nested_prv_ids:\n        errors.append",
                         self.gas.replace("\r\n", "\n"))
        self.assertNotIn("'| {} | {:.1f} | {} |'.format(\n"
                         "                n.fixture_name,",
                         self.gas.replace("\r\n", "\n"))
        self.assertIn("report_generator.md_cell(n.fixture_name)", self.gas)

    def test_failed_pick_is_not_called_a_cancel(self):
        self.assertIn("ui_helpers.last_pick_error()", self.src)


if __name__ == "__main__":
    unittest.main()
