# -*- coding: ascii -*-
"""Regression tests for the gas sizing review fixes (2026-11-20).

One test class per behavior, numbered as in the review brief:
  1  fitting_target_size never sizes a new tee below the existing main
  2  each Size Gas write step checks its Commit status; a step that did not
     commit reports zero successes and the later steps do not run
  3  Size Gas stops with the zero load message at 0 MBH total
  4  a meter joined straight to a mid-stream PRV sizes
  5  propane CFH Table 402.4(37) uses the propane heat content
  6  the meter and PRVs are never resized (fitting_resize_exempt)
  7  the Table 402.4(13) 150 CFH check counts the existing load
  8  choose_sized_type keeps the current Type, then the same name stem
  9  402.4(35) and 402.4(37) labels say regulator to building
  10 reconnect_candidate_reason rejects existing and other system pipes
  11 a pipe leading only to 0 MBH data error fixtures is left unsized
  12 a 'joint' node is never resized or looked up in the fitting step

Plain CPython with the Revit assemblies stubbed (_test_prv_zones). The Size
Gas pushbutton script is exec'd with its final main() call removed, pyRevit
and ui_helpers stubbed, so its helper functions run against stand-in graphs.
Revit Transaction, TransactionStatus and the failure API are stand-ins too.
main() itself is NOT run: its wiring (items 3 and 7) is checked by reading
its source. Nothing here proves behavior inside Revit.

Run:  python -I _test_gas_sizing_review.py
"""

import os
import re
import sys
import types
import unittest

try:
    from unittest import mock
except ImportError:     # pragma: no cover
    import mock

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import _test_prv_zones as Z    # noqa: E402  (installs the Revit stubs)
from _test_prv_zones import _node, _edge, _finish, _Id   # noqa: E402

import gas_tables      # noqa: E402
import pipe_graph      # noqa: E402
import revit_helpers   # noqa: E402
import sizing_engine   # noqa: E402
import shared_params   # noqa: E402

SIZE_GAS = os.path.normpath(os.path.join(
    HERE, "..", "Developer Tools.tab", "Sizing.panel", "Size Gas.pushbutton",
    "script.py"))

PROPANE_CFH = "402.4(37)"     # PE CTS, undiluted propane, CFH
PROPANE_2PSI = "402.4(36)"    # PE, propane, 2 psi, MBH
PROPANE_MBH = "402.4(28)"     # Sched 40 steel, propane, MBH
COPPER_HOUSE_LINE = "402.4(13)"
ONELINE = os.path.normpath(os.path.join(
    HERE, "..", "Developer Tools.tab", "Designer Tools.panel",
    "One-Line.pushbutton", "script.py"))


def _need(owner, name):
    """owner.name, or a test FAILURE (not an error) when the code under test
    does not have it, so the suite reports cleanly against older code."""
    if isinstance(owner, dict):
        value = owner.get(name)
    else:
        value = getattr(owner, name, None)
    if value is None:
        raise AssertionError("{} is not defined in the code under test".format(
            name))
    return value


# ---------------------------------------------------------------------------
# Size Gas script loader
# ---------------------------------------------------------------------------
_SCRIPT_NS = []


def _size_gas():
    """The Size Gas script's globals, exec'd once with main() not called."""
    if _SCRIPT_NS:
        return _SCRIPT_NS[0]
    with open(SIZE_GAS) as f:
        src = f.read()
    src = re.sub(r"^main\(\)\s*$", "pass", src, flags=re.M)

    pyrevit = sys.modules["pyrevit"]
    script = types.ModuleType("pyrevit.script")
    script.get_output = lambda: mock.MagicMock()
    script.get_logger = lambda: mock.MagicMock()
    had_script = hasattr(pyrevit, "script")
    pyrevit.script = script

    ui = types.ModuleType("ui_helpers")
    ui.rja_alert = lambda *a, **k: None
    saved_ui = sys.modules.get("ui_helpers")
    sys.modules["ui_helpers"] = ui

    # An older script imported Transaction at the top; any stand-in will do
    # for the load (the write tests install _TransactionStubs themselves).
    db = sys.modules["Autodesk.Revit.DB"]
    had_tx = hasattr(db, "Transaction")
    if not had_tx:
        db.Transaction = _Tx

    revit = mock.MagicMock()
    ns = {"__name__": "size_gas_script", "__file__": SIZE_GAS,
          "__revit__": revit}
    try:
        exec(compile(src, SIZE_GAS, "exec"), ns)
    finally:
        if not had_tx:
            del db.Transaction
        if saved_ui is None:
            del sys.modules["ui_helpers"]
        else:
            sys.modules["ui_helpers"] = saved_ui
        if not had_script:
            del pyrevit.script
    _SCRIPT_NS.append(ns)
    return ns


def _main_source():
    with open(SIZE_GAS) as f:
        src = f.read()
    m = re.search(r"^def main\(\):.*?(?=^def |\Z)", src, re.S | re.M)
    return m.group(0)


# ---------------------------------------------------------------------------
# Revit Transaction stand-ins
# ---------------------------------------------------------------------------
class _TxStatus(object):
    Committed = "Committed"
    RolledBack = "RolledBack"


class _Options(object):
    def SetFailuresPreprocessor(self, handler):
        self.handler = handler

    def SetClearAfterRollback(self, value):
        self.clear = value


class _Tx(object):
    """A Transaction whose Commit returns the status set for its name
    (matched by substring), Committed otherwise."""
    statuses = {}
    made = []

    def __init__(self, doc, name):
        self.name = name
        self._started = self._ended = False
        _Tx.made.append(name)

    def GetFailureHandlingOptions(self):
        return _Options()

    def SetFailureHandlingOptions(self, options):
        pass

    def Start(self):
        self._started = True

    def Commit(self):
        self._ended = True
        for key, status in _Tx.statuses.items():
            if key in self.name:
                return status
        return _TxStatus.Committed

    def RollBack(self):
        self._ended = True

    def HasStarted(self):
        return self._started

    def HasEnded(self):
        return self._ended


class _Doc(object):
    def Regenerate(self):
        pass


class _TransactionStubs(object):
    """Adds Transaction and the failure API to the Autodesk.Revit.DB stub."""
    NAMES = {"Transaction": _Tx, "TransactionStatus": _TxStatus,
             "IFailuresPreprocessor": object,
             "FailureProcessingResult": types.SimpleNamespace(
                 Continue="Continue", ProceedWithRollBack="RollBack"),
             "FailureSeverity": types.SimpleNamespace(Warning="Warning")}

    def __enter__(self):
        self.db = sys.modules["Autodesk.Revit.DB"]
        self.saved = {}
        for name, value in self.NAMES.items():
            self.saved[name] = getattr(self.db, name, None)
            setattr(self.db, name, value)
        _Tx.statuses = {}
        _Tx.made = []
        cache = getattr(revit_helpers, "_PREPROCESSOR_CLASS", None)
        if cache is not None:
            del cache[:]
        return self

    def __exit__(self, *args):
        for name, value in self.saved.items():
            if value is None:
                delattr(self.db, name)
            else:
                setattr(self.db, name, value)
        cache = getattr(revit_helpers, "_PREPROCESSOR_CLASS", None)
        if cache is not None:
            del cache[:]


# ---------------------------------------------------------------------------
# Graph helpers
# ---------------------------------------------------------------------------
class _E(object):
    def __init__(self, eid, a, b, load):
        self.element_id, self.from_node_id, self.to_node_id = eid, a, b
        self.cumulative_load_mbh = load


class _G(object):
    def __init__(self, edges, children=None):
        self.edges = dict((e.element_id, e) for e in edges)
        self.node_children = children or {}


class _Stub(object):
    """An element stand-in for fittings and pipes (no IsValidObject flag, so
    revit_helpers.is_valid_element reads it as live)."""
    def __init__(self, eid):
        self.Id = _Id(eid)
        self.Location = None


def _with_pipes(graph):
    for eid, e in graph.edges.items():
        e.pipe = _Stub(eid)
    return graph


# ===========================================================================
# 1  fitting_target_size: a new tee never smaller than the existing main
# ===========================================================================
class T01FittingTargetSizeExistingMain(unittest.TestCase):
    """New tee 2 cut into an existing 2 in. main (pipes 101 in, 103 out) with
    a new 1 in. branch 102."""

    def _graph(self):
        return _G([_E(101, 1, 2, 500.0), _E(103, 2, 9, 400.0),
                   _E(102, 2, 4, 100.0)])

    def test_new_tee_on_existing_main_targets_the_main(self):
        got = sizing_engine.fitting_target_size(
            2, self._graph(), {102: "1"}, set(),
            existing_sizes={101: "2", 103: "2"})
        self.assertEqual(got, ("2", False))

    def test_unreadable_existing_size_raises(self):
        with self.assertRaises(ValueError) as ctx:
            sizing_engine.fitting_target_size(
                2, self._graph(), {102: "1"}, set(),
                existing_sizes={101: None, 103: "2"})
        self.assertIn("101", str(ctx.exception))

    def test_existing_pipe_alone_never_sets_a_size(self):
        got = sizing_engine.fitting_target_size(
            2, self._graph(), {}, set(),
            existing_sizes={101: "2", 103: "2"})
        self.assertEqual(got, (None, False))


# ===========================================================================
# 2  Transactions: the Commit status is checked, later steps are skipped
# ===========================================================================
class T02WriteStepsCheckCommit(unittest.TestCase):
    """meter 1 -> tee 2 -> (PRV 3 as a plain fitting) -> RTU-1 4,
    tee 2 -> GUH-1 5. Pipes 101, 102 are distribution pipes; 103 and 104
    are fixture stubs (queued for STEP 9)."""

    def setUp(self):
        self.ns = _size_gas()
        self.tx = _TransactionStubs().__enter__()
        self.graph = _with_pipes(Z._build(with_prv=False))
        self.doc = _Doc()
        self.result = sizing_engine.size_system(
            self.graph, "Schedule 40 Steel", 0.5, Z.DOWN,
            heat_content_btu_per_cf=840.0)
        self.pipe_writes = []
        self.fit_writes = []
        self.patches = [
            mock.patch.object(revit_helpers, "is_new_construction",
                              lambda el, doc: True),
            mock.patch.object(revit_helpers, "set_pipe_diameter",
                              self._set_pipe),
            mock.patch.object(revit_helpers, "set_fitting_size",
                              self._set_fit),
            mock.patch.object(revit_helpers, "is_transition",
                              lambda el: False),
        ]
        for p in self.patches:
            p.start()

    def tearDown(self):
        for p in self.patches:
            p.stop()
        self.tx.__exit__(None, None, None)

    def _set_pipe(self, pipe, inches):
        self.pipe_writes.append(revit_helpers.eid_int(pipe.Id))
        return True, "RBS_PIPE_DIAMETER_PARAM"

    def _set_fit(self, element, inches, piping_only=False):
        # Size Gas asks for piping connectors only on a Type swap (item G)
        self.assertTrue(piping_only)
        self.fit_writes.append(revit_helpers.eid_int(element.Id))
        return True, "Nominal Radius"

    def _steps(self):
        ns = self.ns
        return (_need(ns, "_write_pipe_sizes"),
                _need(ns, "_write_fitting_sizes"),
                _need(ns, "_write_stub_sizes"))

    def test_committed_run_reports_successes(self):
        pipes, fits, stubs = self._steps()
        w = pipes(self.doc, self.graph, self.result)
        self.assertTrue(w["committed"])
        self.assertEqual(w["success_count"], 2)
        self.assertEqual(len(w["skipped_stubs"]), 2)

    def test_pipe_step_rolled_back_reports_zero_and_skips_the_rest(self):
        pipes, fits, stubs = self._steps()
        _Tx.statuses = {"Size Gas Pipes": _TxStatus.RolledBack}
        w = pipes(self.doc, self.graph, self.result)
        self.assertFalse(w["committed"])
        self.assertEqual(w["success_count"], 0)
        self.assertEqual(w["fail_count"], 2)    # both attempted writes
        self.assertTrue(w["error"])

        fit, fit_committed = fits(self.doc, self.graph, w["new_sizes"],
                                  set(), {}, w["committed"])
        self.assertFalse(fit_committed)
        self.assertEqual(fit["resized"], 0)
        self.assertEqual(self.fit_writes, [])
        self.assertFalse(any("Fittings" in n for n in _Tx.made))

        before = list(self.pipe_writes)
        st = stubs(self.doc, self.graph, w["skipped_stubs"], fit_committed)
        self.assertEqual(st["success"], 0)
        self.assertEqual(st["fail"], len(w["skipped_stubs"]))
        self.assertEqual(self.pipe_writes, before)     # no stub was written
        self.assertFalse(any("Stub" in n for n in _Tx.made))

    def test_fitting_step_rolled_back_reports_zero_and_skips_stubs(self):
        pipes, fits, stubs = self._steps()
        _Tx.statuses = {"Fittings": _TxStatus.RolledBack}
        w = pipes(self.doc, self.graph, self.result)
        self.assertTrue(w["committed"])
        fit, fit_committed = fits(self.doc, self.graph, w["new_sizes"],
                                  set(), {}, w["committed"])
        self.assertTrue(self.fit_writes)     # the tee was written ...
        self.assertFalse(fit_committed)      # ... then rolled back
        self.assertEqual(fit["resized"], 0)
        st = stubs(self.doc, self.graph, w["skipped_stubs"], fit_committed)
        self.assertEqual(st["success"], 0)
        self.assertEqual(st["fail"], 2)
        self.assertFalse(any("Stub" in n for n in _Tx.made))

    def test_stub_step_rolled_back_counts_every_stub_failed(self):
        pipes, fits, stubs = self._steps()
        _Tx.statuses = {"Stub": _TxStatus.RolledBack}
        w = pipes(self.doc, self.graph, self.result)
        fit, fit_committed = fits(self.doc, self.graph, w["new_sizes"],
                                  set(), {}, w["committed"])
        self.assertTrue(fit_committed)
        st = stubs(self.doc, self.graph, w["skipped_stubs"], fit_committed)
        self.assertEqual(st["success"], 0)
        self.assertEqual(st["fail"], len(w["skipped_stubs"]))
        self.assertEqual(st["fail"], 2)
        self.assertTrue(st["rolled_back"])

        info = _need(self.ns, "_summary_info")(
            self.result, None, w, st, fit, [_need(self.ns, "STUB_STEP")])
        self.assertTrue(info["stub_rolled_back"])
        self.assertEqual(info["stub_success"], 0)
        self.assertEqual(info["stub_fail"], 2)


# ===========================================================================
# 3  Size Gas stops at a 0 MBH total load
# ===========================================================================
def _load_graph(*loads, **kw):
    """A stand-in graph with one gas fixture per load (and one plain
    fitting), for pipe_graph.total_connected_mbh and the stop messages."""
    nodes = {99: types.SimpleNamespace(is_gas_fixture=False, gas_load_mbh=0)}
    for i, v in enumerate(loads):
        nodes[i + 1] = types.SimpleNamespace(is_gas_fixture=True,
                                             gas_load_mbh=v)
    return types.SimpleNamespace(nodes=nodes, existing=kw.get("existing"))


class T03ZeroLoadStop(unittest.TestCase):

    def test_total_connected_mbh(self):
        fn = _need(pipe_graph, "total_connected_mbh")
        self.assertEqual(fn(_load_graph(0.0, None, 64.8, 35.2)), 100.0)
        self.assertIsInstance(fn(_load_graph()), float)
        self.assertEqual(fn(_load_graph()), 0.0)

    def test_zero_total_gives_the_stop_message(self):
        msg = _need(pipe_graph, "zero_load_stop_message")(
            _load_graph(0.0, None, 0))
        self.assertIsNotNone(msg)
        self.assertIn("0 MBH", msg)
        self.assertIn("3 gas fixture", msg)
        self.assertIn("Nothing was changed", msg)

    def test_any_load_gives_no_message(self):
        fn = _need(pipe_graph, "zero_load_stop_message")
        self.assertIsNone(fn(_load_graph(0.0, 64.8)))

    def test_script_keeps_no_copy(self):
        ns = _size_gas()
        for name in ("_zero_load_message", "_shown_heat_content"):
            self.assertNotIn(name, ns)
        src = _main_source()
        self.assertNotIn("gas_load_mbh <= 0", src)
        self.assertNotIn("no_load=", src)

    def test_main_stops_before_sizing(self):
        src = _main_source()
        at = src.find("pipe_graph.zero_load_stop_message(graph)")
        self.assertGreater(at, 0, "main() does not call "
                                  "zero_load_stop_message")
        sizing = src.find("sizing_engine.size_system(")
        self.assertGreater(sizing, at)
        block = src[at:src.find("\n\n", at)]
        self.assertIn("_alert(zero_load", block)
        self.assertIn("return", block)


# ===========================================================================
# 4  Meter joined straight to a mid-stream PRV
# ===========================================================================
class T04MeterStraightToPrv(unittest.TestCase):
    """meter 1 joined to PRV 3 with no pipe; PRV 3 -> pipe 103 (20 ft) ->
    tee 7 -> pipe 104 (15 ft) -> RTU-1 4, tee 7 -> pipe 105 -> UH-1 5."""

    def _graph(self):
        g = pipe_graph.NetworkGraph()
        g.origin_id = 1
        _node(g, 1, "meter")
        _node(g, 3, "fitting", prv=True)
        _node(g, 7, "tee")
        _node(g, 4, "fixture", 800.0, "RTU-1")
        _node(g, 5, "fixture", 200.0, "UH-1")
        g.node_children[1] = [3]
        _edge(g, 103, 3, 7, 20.0)
        _edge(g, 104, 7, 4, 15.0)
        _edge(g, 105, 7, 5, 5.0)
        _finish(g)
        return g

    def test_sizes_without_an_error(self):
        g = self._graph()
        self.assertEqual(g.midstream_prv_ids, [3])
        try:
            r = sizing_engine.size_system(
                g, "Schedule 40 Steel", 2.0, Z.UP, downstream_table_id=Z.DOWN)
        except ValueError as e:
            self.fail("size_system raised: {}".format(e))
        self.assertEqual(sorted(r["sizes"]), [103, 104, 105])
        for pid in (103, 104, 105):
            self.assertEqual(r["edge_context"][pid]["table_id"], Z.DOWN)
        # 20 + 15 pipe + 3 ft tee to RTU-1 on the regulated side.
        self.assertEqual(r["sizes"][103],
                         Z._expected_size(Z.DOWN, 38.0, 1000.0))
        self.assertEqual(r["zones"][0]["table_id"], Z.UP)
        self.assertEqual(r["zones"][0]["run_ft"], 0.0)
        self.assertEqual(r["zones"][1]["table_id"], Z.DOWN)


# ===========================================================================
# 5  Propane CFH Table 402.4(37) uses the propane heat content
# ===========================================================================
class T05PropaneHeatContent(unittest.TestCase):

    def test_402_4_37_converts_150_mbh_to_60_cfh(self):
        value, units = sizing_engine.demand_in_table_units(
            150.0, PROPANE_CFH, 840.0)
        self.assertEqual(units, "CFH")
        self.assertAlmostEqual(value, 60.0)
        value, units = sizing_engine.demand_in_table_units(
            150.0, PROPANE_CFH, 840.0, 2500.0)
        self.assertAlmostEqual(value, 60.0)
        self.assertEqual(_need(sizing_engine, "heat_content_for_table")(
            PROPANE_CFH, 840.0), 2500.0)

    def test_402_4_37_after_a_prv_uses_propane(self):
        # meter 1 -> pipe 101 (30 ft) -> PRV 3 -> pipe 103 (20 ft) -> 150 MBH
        g = pipe_graph.NetworkGraph()
        g.origin_id = 1
        _node(g, 1, "meter")
        _node(g, 3, "fitting", prv=True)
        _node(g, 4, "fixture", 150.0, "FURN-1")
        _edge(g, 101, 1, 3, 30.0)
        _edge(g, 103, 3, 4, 20.0)
        _finish(g)
        r = sizing_engine.size_system(
            g, "Polyethylene Plastic Pipe", 2.0, PROPANE_2PSI,
            heat_content_btu_per_cf=840.0, downstream_table_id=PROPANE_CFH)
        detail = dict((d["pipe_id"], d) for d in r["segment_detail"])
        self.assertEqual(detail[103]["units"], "CFH")
        self.assertAlmostEqual(detail[103]["demand_cfh"], 60.0)
        # The meter side is an MBH table: the load is compared directly.
        self.assertEqual(detail[101]["units"], "MBH")
        self.assertAlmostEqual(detail[101]["demand_cfh"], 150.0)
        zones = dict((z["table_id"], z) for z in r["zones"])
        self.assertEqual(zones[PROPANE_CFH].get("heat_content_btu_per_cf"),
                         2500.0)
        self.assertIsNone(zones[PROPANE_2PSI].get("heat_content_btu_per_cf",
                                                  "missing"))

    def test_mbh_table_uses_no_heat_content(self):
        fn = _need(sizing_engine, "heat_content_for_table")
        self.assertIsNone(fn(PROPANE_MBH, 840.0, 2500.0))
        self.assertEqual(sizing_engine.demand_in_table_units(
            150.0, PROPANE_MBH, 840.0, 2500.0), (150.0, "MBH"))

    def test_natural_gas_cfh_table_uses_the_natural_value(self):
        fn = _need(sizing_engine, "heat_content_for_table")
        self.assertEqual(fn(Z.DOWN, 840.0, 2500.0), 840.0)
        value, units = sizing_engine.demand_in_table_units(
            150.0, Z.DOWN, 840.0, 2500.0)
        self.assertEqual(units, "CFH")
        self.assertAlmostEqual(value, 150000.0 / 840.0)


# ===========================================================================
# 6  The meter and every PRV are never resized
# ===========================================================================
class T06MeterAndPrvNeverResized(unittest.TestCase):

    def test_fitting_resize_exempt(self):
        fn = _need(sizing_engine, "fitting_resize_exempt")
        g = Z._build()
        self.assertTrue(fn(g.nodes[1], g))      # meter
        self.assertTrue(fn(g.nodes[3], g))      # mid-stream PRV
        self.assertIsNone(fn(g.nodes[2], g))    # tee
        equip = Z._build(prv_to_fixture_ft=4.0)  # equipment PRV
        self.assertTrue(fn(equip.nodes[3], equip))
        self.assertIsNone(fn(None, g))

    def test_resize_fittings_never_passes_them_to_set_fitting_size(self):
        ns = _size_gas()
        g = _with_pipes(Z._build())
        # The meter read as a 2 connector fitting, as a meter family with no
        # meter classification can be.
        g.nodes[1].node_type = "fitting"
        new_sizes = dict((eid, "2") for eid in g.edges)
        calls = []

        def _set_fit(element, inches, piping_only=False):
            calls.append(revit_helpers.eid_int(element.Id))
            return True, "Nominal Radius"

        with mock.patch.object(revit_helpers, "is_new_construction",
                               lambda el, doc: True), \
                mock.patch.object(revit_helpers, "is_transition",
                                  lambda el: False), \
                mock.patch.object(revit_helpers, "set_fitting_size",
                                  _set_fit):
            counts = ns["_resize_fittings"](g, new_sizes, _Doc(), set())
        self.assertIn(2, calls)          # the tee is still resized
        self.assertNotIn(3, calls)       # PRV
        self.assertNotIn(1, calls)       # meter
        self.assertEqual(counts.get("prv"), 2)


# ===========================================================================
# 7  Table 402.4(13) 150 CFH limit counts the existing load
# ===========================================================================
class T07HouseLineLimitWithExistingLoad(unittest.TestCase):

    def test_existing_plus_new_over_150_cfh_warns(self):
        fn = _need(_size_gas(), "_table_limit_check")
        graph = _load_graph(60.0, existing={"load_mbh": 100.0})
        warnings = fn(COPPER_HOUSE_LINE, graph, 840.0, 2500.0)
        self.assertEqual(len(warnings), 1)
        self.assertIn("150 CFH", warnings[0])
        self.assertIn("190 CFH (160 MBH)", warnings[0])

    def test_new_load_alone_is_within_the_limit(self):
        fn = _need(_size_gas(), "_table_limit_check")
        self.assertEqual(fn(COPPER_HOUSE_LINE, _load_graph(60.0),
                            840.0, 2500.0), [])

    def test_main_passes_the_graph(self):
        found = re.search(r"_table_limit_check\(\s*table_id,\s*graph,",
                          _main_source())
        self.assertTrue(found, "main() does not call _table_limit_check with "
                               "the graph")


# ===========================================================================
# 8  choose_sized_type
# ===========================================================================
class T08ChooseSizedType(unittest.TestCase):

    def test_keeps_the_current_type_at_the_size(self):
        fn = _need(revit_helpers, "choose_sized_type")
        cands = [("thr", 1.5, '1-1/2" Threaded'), ("cur", 1.5, '1-1/2" Press')]
        self.assertEqual(fn(("cur", 1.5, '1-1/2" Press'), cands, 1.5), "cur")

    def test_prefers_the_same_name_apart_from_size(self):
        fn = _need(revit_helpers, "choose_sized_type")
        cands = [("thr", 1.5, '1-1/2" Threaded'),
                 ("prs", 1.5, '1-1/2" Press'),
                 ("one", 1.0, '1" Press')]
        self.assertEqual(fn(("one", 1.0, '1" Press'), cands, 1.5), "prs")
        ball = [("a", 0.75, '0 3/4"'), ("b", 2.0, '2"'), ("c", 2.0, '2" Alt')]
        self.assertEqual(fn(("x", 1.0, '1"'), ball, 2.0), "b")

    def test_first_match_when_no_stem_matches_and_none_when_no_size(self):
        fn = _need(revit_helpers, "choose_sized_type")
        cands = [("thr", 1.5, '1-1/2" Threaded'), ("wld", 1.5, '1-1/2" Weld')]
        self.assertEqual(fn(("one", 1.0, '1" Press'), cands, 1.5), "thr")
        self.assertIsNone(fn(("one", 1.0, '1" Press'), cands, 2.0))


# ===========================================================================
# 9  Propane PE table labels
# ===========================================================================
class T09PropaneLabels(unittest.TestCase):

    def test_labels_say_regulator_to_building(self):
        labels = dict((o["table_id"], o["label"])
                      for o in gas_tables.TABLE_OPTIONS)
        for tid in ("402.4(35)", "402.4(37)"):
            self.assertIn("regulator to building", labels[tid], tid)
            self.assertNotIn("regulator to appliance", labels[tid], tid)


# ===========================================================================
# 10 reconnect_candidate_reason
# ===========================================================================
class _Phases(object):
    def __init__(self, ids):
        self._ids = ids
        self.Size = len(ids)

    def get_Item(self, i):
        return types.SimpleNamespace(Id=_Id(self._ids[i]))


class _PhaseDoc(object):
    def __init__(self):
        self.Phases = _Phases([100, 200])     # 200 = New Construction


class _Pipe(object):
    def __init__(self, eid, phase, system_id):
        self.Id = _Id(eid)
        self.CreatedPhaseId = _Id(phase)
        self._system = system_id

    def get_Parameter(self, bip):
        system = self._system
        return types.SimpleNamespace(AsElementId=lambda: _Id(system))


class T10ReconnectCandidateReason(unittest.TestCase):

    def setUp(self):
        bip = revit_helpers.BuiltInParameter
        self._added = not hasattr(bip, "RBS_PIPING_SYSTEM_TYPE_PARAM")
        if self._added:
            bip.RBS_PIPING_SYSTEM_TYPE_PARAM = "RBS_PIPING_SYSTEM_TYPE_PARAM"

    def tearDown(self):
        if self._added:
            del revit_helpers.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM

    def test_new_pipe_on_the_valve_system_may_move(self):
        fn = _need(revit_helpers, "reconnect_candidate_reason")
        self.assertIsNone(fn(_PhaseDoc(), _Pipe(10, 200, 500), set([500])))

    def test_existing_phase_pipe_is_rejected(self):
        fn = _need(revit_helpers, "reconnect_candidate_reason")
        why = fn(_PhaseDoc(), _Pipe(11, 100, 500), set([500]))
        self.assertIsNotNone(why)
        self.assertIn("existing", why)

    def test_pipe_on_another_system_type_is_rejected(self):
        fn = _need(revit_helpers, "reconnect_candidate_reason")
        why = fn(_PhaseDoc(), _Pipe(12, 200, 600), set([500]))
        self.assertIsNotNone(why)
        self.assertIn("another piping system", why)

    def test_no_readable_valve_system_rejects(self):
        fn = _need(revit_helpers, "reconnect_candidate_reason")
        self.assertIsNotNone(fn(_PhaseDoc(), _Pipe(13, 200, 500), set()))


# ===========================================================================
# 11 Pipes leading only to 0 MBH data error fixtures are left unsized
# ===========================================================================
class T11ZeroLoadDataErrorUnsized(unittest.TestCase):
    """meter 1 -> pipe 101 -> tee 2 -> pipe 102 -> PRV-less fitting 3 ->
    pipe 103 -> RTU-1 4 (800 MBH); tee 2 -> pipe 104 -> GUH-1 5 (0 MBH,
    a data error)."""

    def _graph(self):
        g = pipe_graph.NetworkGraph()
        g.origin_id = 1
        _node(g, 1, "meter")
        _node(g, 2, "tee")
        _node(g, 3, "fitting")
        _node(g, 4, "fixture", 800.0, "RTU-1")
        _node(g, 5, "fixture", 0.0, "GUH-1")
        _edge(g, 101, 1, 2, 50.0)
        _edge(g, 102, 2, 3, 30.0)
        _edge(g, 103, 3, 4, 20.0)
        _edge(g, 104, 2, 5, 40.0)
        g.data_errors = [{"element_id": 5, "name": "GUH-1", "kind": "zero",
                          "message": "GAS_LOAD_MBH is 0"}]
        _finish(g)
        return g

    def _size(self, g):
        return sizing_engine.size_system(
            g, "Schedule 40 Steel", 0.5, Z.DOWN, heat_content_btu_per_cf=840.0)

    def test_pipe_is_left_unsized_and_listed(self):
        g = self._graph()
        r = self._size(g)
        self.assertNotIn(104, r["sizes"])
        self.assertEqual(r.get("unsized_zero_load"),
                         [{"pipe_id": 104, "fixture_ids": [5],
                           "table_id": Z.DOWN}])
        # The rest of the system still sizes.
        for pid in (101, 102, 103):
            self.assertIn(pid, r["sizes"])

    def test_sizing_row_is_data_error(self):
        g = self._graph()
        rows = dict((row["pipe_id"], row) for row in
                    sizing_engine.sizing_rows(self._size(g), g))
        self.assertEqual(rows[104]["status"], "DATA ERROR")
        self.assertEqual(rows[104]["size"], "")
        self.assertIn("5", rows[104]["note"])
        self.assertEqual(rows[101]["status"], "OK")

    def test_zero_load_fixture_not_flagged_still_gets_minimum(self):
        # Only a recorded data error leaves a pipe unsized.
        g = self._graph()
        g.data_errors = [{"element_id": 999}]
        r = self._size(g)
        self.assertIn(104, r["sizes"])
        self.assertEqual(r.get("unsized_zero_load"), [])


# ===========================================================================
# 12 A 'joint' node is never resized or looked up in the fitting step
# ===========================================================================
class T12JointNeverResized(unittest.TestCase):
    """meter 1 -> pipe 10 -> joint 10 -> pipe 11 -> RTU-1 4. The joint is
    keyed by pipe 10's id and has no element (pipe_graph._add_joint)."""

    def _graph(self):
        g = pipe_graph.NetworkGraph()
        g.origin_id = 1
        _node(g, 1, "meter")
        _node(g, 4, "fixture", 100.0, "RTU-1")
        joint = pipe_graph.NetworkNode(10, None, "joint")
        joint.element = None
        joint.family_name = "PIPE JOINT"
        g.add_node(joint)
        _edge(g, 10, 1, 10, 10.0)
        _edge(g, 11, 10, 4, 10.0)
        _finish(g)
        return _with_pipes(g)

    def _run(self, call):
        g = self._graph()
        looked_up, resized, transition_reads = [], [], []

        def _current(doc, eid):
            looked_up.append(eid)
            return None

        def _set_fit(element, inches, piping_only=False):
            resized.append(revit_helpers.eid_int(element.Id))
            return True, "Nominal Radius"

        def _is_transition(el):
            transition_reads.append(el)
            return False

        with _TransactionStubs(), \
                mock.patch.object(revit_helpers, "current_element", _current,
                                  create=True), \
                mock.patch.object(revit_helpers, "set_fitting_size",
                                  _set_fit), \
                mock.patch.object(revit_helpers, "is_new_construction",
                                  lambda el, doc: True), \
                mock.patch.object(revit_helpers, "is_transition",
                                  _is_transition):
            counts = call(g, {10: "1", 11: "1"})
        self.assertNotIn(10, looked_up)
        self.assertNotIn(10, resized)
        self.assertNotIn(None, transition_reads)
        self.assertEqual(counts["no_element"], 0)
        self.assertEqual(counts["failures"], [])
        return g

    def test_resize_fittings_skips_the_joint(self):
        ns = _size_gas()
        self._run(lambda g, sizes: ns["_resize_fittings"](
            g, sizes, _Doc(), set()))

    def test_fitting_step_skips_the_joint(self):
        step = _need(_size_gas(), "_write_fitting_sizes")
        self._run(lambda g, sizes: step(_Doc(), g, sizes, set(), {}, True)[0])

    def test_transition_ids_skip_the_joint(self):
        ns = _size_gas()
        g = self._graph()
        with mock.patch.object(revit_helpers, "is_transition",
                               lambda el: self.fail("read {}".format(el))
                               if el is None else False):
            self.assertNotIn(10, ns["_transition_ids"](g))


# ===========================================================================
# 13  Size Gas write fixes (2026-11-20 cleanup, item H)
# ===========================================================================
class T13WriteFixes(unittest.TestCase):
    """meter 1 -> tee 2 -> (fitting 3) -> RTU-1 4, tee 2 -> GUH-1 5.
    Pipe 104 feeds GUH-1 (5)."""

    def setUp(self):
        self.ns = _size_gas()
        self.tx = _TransactionStubs().__enter__()
        self.graph = _with_pipes(Z._build(with_prv=False))
        self.result = sizing_engine.size_system(
            self.graph, "Schedule 40 Steel", 0.5, Z.DOWN,
            heat_content_btu_per_cf=840.0)
        self.patches = [
            mock.patch.object(revit_helpers, "is_new_construction",
                              lambda el, doc: True),
            mock.patch.object(revit_helpers, "set_pipe_diameter",
                              lambda pipe, inches: (True, "RBS")),
        ]
        for p in self.patches:
            p.start()

    def tearDown(self):
        for p in self.patches:
            p.stop()
        self.tx.__exit__(None, None, None)

    def test_existing_appliance_stub_not_in_new_sizes(self):
        self.graph.existing_appliances = [5]
        w = self.ns["_write_pipe_sizes"](_Doc(), self.graph, self.result)
        self.assertNotIn(104, w["new_sizes"])      # kept as drawn
        self.assertIn(103, w["new_sizes"])         # new stub, written later
        self.assertIn(101, w["new_sizes"])         # distribution pipe
        self.assertEqual([s["pipe_id"] for s in w["skipped_stubs"]], [103])
        self.assertEqual(w["existing_skip_count"], 1)

    def test_existing_pipe_read_on_its_own_zone_table(self):
        g = _with_pipes(Z._build())               # mid stream PRV 3
        result = sizing_engine.size_system(
            g, "Schedule 40 Steel", 2.0, Z.UP, heat_content_btu_per_cf=840.0,
            downstream_table_id=Z.DOWN)
        # an open ended existing pipe off the PRV side
        open_edge = _edge(g, 199, 3, None, 5.0)
        open_edge.pipe = _Stub(199)
        open_edge.run_key = 3
        read = []

        def _nominal(pipe, table_id):
            read.append((revit_helpers.eid_int(pipe.Id), table_id))
            return "1"
        with mock.patch.object(revit_helpers, "is_new_construction",
                               lambda el, doc: False), \
                mock.patch.dict(self.ns, {"_get_pipe_nominal_size": _nominal}):
            sizes = self.ns["_existing_pipe_sizes"](g, {}, result, _Doc())
        tables = dict(read)
        self.assertEqual(tables[103], Z.DOWN)      # after the PRV
        self.assertEqual(tables[101], Z.UP)        # meter side
        self.assertNotIn(199, tables)              # open end: not read
        self.assertNotIn(199, sizes)
        self.assertIsNone(sizing_engine.edge_table_id(result, open_edge))

    def test_rolled_back_pipe_step_names_no_approach(self):
        _Tx.statuses = {"Size Gas Pipes": _TxStatus.RolledBack}
        self.ns["_confirmed_approach"][0] = "RBS_PIPE_DIAMETER_PARAM"
        try:
            w = self.ns["_write_pipe_sizes"](_Doc(), self.graph, self.result)
            st = {"success": 0, "fail": 0, "rows": [], "rolled_back": False}
            info = self.ns["_summary_info"](
                self.result, None, w, st, self.ns["_new_fit_counts"](),
                [self.ns["PIPE_STEP"]])
        finally:
            self.ns["_confirmed_approach"][0] = None
        self.assertIsNone(info["approach"])
        import gas_console
        rows = dict(gas_console._summary_rows(info))
        self.assertEqual(rows["API approach used"],
                         "None, the pipe sizes were rolled back")

    def test_stub_rollback_keeps_only_the_rollback_row(self):
        _Tx.statuses = {"Stub": _TxStatus.RolledBack}
        stubs = [{"pipe_id": 103, "recommended_size": "1",
                  "fixture_name": "RTU-1", "demand_mbh": 800.0},
                 {"pipe_id": 104, "recommended_size": "BAD",
                  "fixture_name": "GUH-1", "demand_mbh": 400.0}]
        st = self.ns["_write_stub_sizes"](_Doc(), self.graph, stubs, True)
        self.assertTrue(st["rolled_back"])
        self.assertEqual(len(st["rows"]), 1)
        self.assertEqual(st["rows"][0]["id"], "-")
        self.assertIn("rolled back", st["rows"][0]["reason"])


# ===========================================================================
# 14  One stop message for a ticked Mid Stream PRV with none found (E, F)
# ===========================================================================
class T14MidStreamPrvStop(unittest.TestCase):

    def test_no_message_when_box_off_or_prv_found(self):
        fn = pipe_graph.midstream_prv_stop_message
        self.assertIsNone(fn(Z._build(with_prv=False), False))
        self.assertIsNone(fn(Z._build(), True))

    def test_no_regulator_names_every_keyword(self):
        msg = pipe_graph.midstream_prv_stop_message(
            Z._build(with_prv=False), True)
        self.assertIn("Mid Stream PRV is checked", msg)
        for kw in shared_params.PRV_FAMILY_KEYWORDS:
            self.assertIn("'{}'".format(kw), msg)

    def test_equipment_prv_only(self):
        msg = pipe_graph.midstream_prv_stop_message(
            Z._build(prv_to_fixture_ft=4.0), True)
        self.assertIn("1 regulator(s) were found", msg)
        self.assertIn("treated as equipment PRVs", msg)

    def test_size_gas_and_oneline_both_stop_on_it(self):
        self.assertIn("pipe_graph.midstream_prv_stop_message(graph, "
                      "mid_stream_prv)", _main_source())
        with open(ONELINE) as f:
            src = f.read()
        at = src.find("pipe_graph.midstream_prv_stop_message(")
        self.assertGreater(at, 0)
        block = src[at:src.find("prv_notes = []", at)]
        self.assertIn("_alert(prv_stop", block)
        self.assertIn("return", block)
        # the old behavior (drop the table and draw) is gone
        self.assertNotIn("the downstream table was not", src)
        self.assertNotIn("'regulating' in its", src)
        self.assertNotIn("'regulating' in its", _main_source())


# ===========================================================================
# 15  One gas start check in Size Gas and One-Line (item A)
# ===========================================================================
class T15OneGasStartCheck(unittest.TestCase):

    def test_both_buttons_call_check_gas_start_and_nothing_else(self):
        with open(ONELINE) as f:
            oneline = f.read()
        for src in (_main_source(), oneline):
            self.assertIn("tie_in.check_gas_start(", src)
            self.assertIn("pipe_only=bool(connect_existing)", src)
            for old in ("pipe_graph.check_gas_pipe(",
                        "tie_in.resolve_new_pipe_upstream(",
                        "revit_helpers.validate_selected_element("):
                self.assertNotIn(old, src)


if __name__ == "__main__":
    unittest.main()
