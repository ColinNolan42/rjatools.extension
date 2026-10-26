# -*- coding: ascii -*-
# CPython tests for water_flagging (the numbered problem list, one row per
# run) and water_system_errors (the model-wide error list). Revit is stubbed
# out as in the other _test_water_*.py files.

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_water_graph import (          # noqa: E402
    _install_revit_stubs, make_graph, add_node, COLD)

import water_graph            # noqa: E402
import water_performance as wp  # noqa: E402
import water_flagging         # noqa: E402
import water_system_errors    # noqa: E402
import water_checks           # noqa: E402


def result(pid, label, required=None, drawn=1.0, reason=""):
    r = wp.PipeResult(pid)
    r.label = label
    r.drawn_in = drawn
    r.required_nominal = required
    r.reason = reason
    r.system = "cold"
    return r


def chain(*kinds):
    """RPZ then a chain of nodes, ids 2, 3, 4 ... in the order given."""
    g = make_graph()
    g.origin_id = 1
    add_node(g, 1, water_graph.KIND_ORIGIN)
    parent = 1
    for index, kind in enumerate(kinds):
        nid = index + 2
        add_node(g, nid, kind, parent=parent)
        parent = nid
    return g


P = water_graph.KIND_PIPE
F = water_graph.KIND_FITTING


class TestGrayRuns(unittest.TestCase):

    def test_a_gray_run_is_one_row(self):
        g = chain(P, F, P, P)           # pipes 2, 4, 5 with an elbow between
        results = dict((i, result(i, wp.GRAY,
                                  reason="no fixture demand downstream"))
                       for i in (2, 4, 5))
        items = water_flagging.flag_items(g, results)
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0].result.element_id, 2)
        self.assertEqual(items[0].dependents, [4, 5])
        cells = water_flagging.row_cells(items[0])
        self.assertEqual(cells[1], "GRAY")
        self.assertIn("No demand downstream", cells[2])
        self.assertIn("2 more pipe(s)", cells[2])

    def test_separate_gray_runs_are_separate_rows(self):
        g = chain(P, P)
        add_node(g, 9, P, parent=1)
        results = {2: result(2, wp.GRAY), 3: result(3, wp.GREEN),
                   9: result(9, wp.GRAY)}
        items = water_flagging.flag_items(g, results)
        self.assertEqual([i.result.element_id for i in items], [2, 9])


class TestSizeRuns(unittest.TestCase):

    def test_red_run_reports_the_start_and_the_size_change(self):
        g = chain(P, P, P)
        results = {2: result(2, wp.RED, "1-1/4"),
                   3: result(3, wp.RED, "1-1/4"),
                   4: result(4, wp.RED, "1")}
        items = water_flagging.flag_items(g, results)
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0].steps, [(4, "1")])
        text = water_flagging.row_cells(items[0])[2]
        self.assertTrue(text.startswith("Undersized, 2 more pipe(s)"))
        self.assertIn("1 in at pipe 4", text)

    def test_different_statuses_are_not_grouped(self):
        g = chain(P, P)
        results = {2: result(2, wp.RED, "1-1/4"),
                   3: result(3, wp.PURPLE, "3/4")}
        items = water_flagging.flag_items(g, results)
        self.assertEqual([i.result.label for i in items],
                         [wp.RED, wp.PURPLE])
        self.assertEqual([i.number for i in items], [1, 2])

    def test_row_cells_carry_the_sizes(self):
        g = chain(P)
        r = result(2, wp.RED, "1-1/4", drawn=1.0)
        r.gpm = 24.7
        r.fps = 9.6
        cells = water_flagging.row_cells(water_flagging.flag_items(
            g, {2: r})[0])
        self.assertEqual(cells[4:], ["1 in", "1-1/4 in", "24.7", "9.6"])


class TestSystemErrors(unittest.TestCase):

    def _checks(self, *findings):
        return {"findings": list(findings)}

    def test_errors_first_with_short_text(self):
        g = chain(P)
        warn = water_checks.Finding(water_checks.SEVERITY_WARNING, "w",
                                    "Long message. More.", short="Short warn.")
        err = water_checks.Finding(water_checks.SEVERITY_ERROR, "e",
                                   "Another. More.")
        info = water_checks.Finding(water_checks.SEVERITY_INFO, "i", "Info.")
        sizing = {"segments": []}
        out = water_system_errors.collect(
            g, sizing, None, self._checks(warn, err, info))
        self.assertEqual(out[0][0], water_system_errors.ERROR)
        self.assertEqual(out[0][1], "Another.")
        self.assertEqual(out[1], (water_system_errors.WARNING, "Short warn."))
        self.assertEqual(len(out), 2)

    def test_no_circuit_setters_is_an_error(self):
        g = chain(P)
        recirc = {"legs": [{}, {}], "unrooted_pipe_ids": []}
        out = water_system_errors.collect(
            g, {"segments": []}, recirc, self._checks())
        self.assertTrue(any("No circuit setters" in t for _s, t in out))

    def test_summary_line_counts(self):
        entries = [(water_system_errors.ERROR, "a"),
                   (water_system_errors.WARNING, "b")]
        self.assertIn("1 error(s) and 1 warning(s)",
                      water_system_errors.summary_line(entries))
        self.assertEqual(water_system_errors.summary_line([]),
                         "No system errors found.")


if __name__ == "__main__":
    unittest.main()
