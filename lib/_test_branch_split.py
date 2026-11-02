# -*- coding: ascii -*-
"""Tests for the single terminal unit run check (hvac_graph.single_terminal_runs).

RJA rule (Colin, 2026-11-02): a duct with one terminal downstream is a
BRANCH. A root unit connection that serves only one terminal makes its whole
run one branch, which is nearly always a model set up problem or a rare
point of use unit, so the tools always judge (Duct Performance) or size
(Size Duct) that run as a main and print one short NOTE (Colin, 2026-11-02,
always on, no dialog option). VAV boxes serving one diffuser are
normal and are not reported. The pure block of hvac_graph.py is run on its
own, since hvac_graph itself imports the Revit API.

Run:  python _test_branch_split.py
"""

import io
import os
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))


def _load():
    src = io.open(os.path.join(HERE, "hvac_graph.py"), encoding="utf-8").read()
    start = src.index("def single_terminal_runs(")
    end = src.index("def solid_fill_pattern_id(")
    ns = {}
    exec(src[start:end], ns)
    return ns["single_terminal_runs"], ns["single_terminal_text"]


RUNS, TEXT = _load()


def downstream(root, children, terminals):
    """Terminals at or below each node (plain post order, no side filter)."""
    out = {}

    def walk(n):
        acc = set([n]) if n in terminals else set()
        for c in children.get(n, []):
            acc |= walk(c)
        out[n] = acc
        return acc
    walk(root)
    return out


def runs(root, children, terminals, equipment):
    return RUNS(root, children, downstream(root, children, terminals), equipment)


class TestSingleTerminalRuns(unittest.TestCase):

    def test_unit_feeding_one_diffuser_is_reported_whole_run(self):
        # unit 1 -> duct 2 -> elbow 3 -> duct 4 -> flex 5 -> diffuser 6
        ch = {1: [2], 2: [3], 3: [4], 4: [5], 5: [6]}
        self.assertEqual(runs(1, ch, {6}, {1}), {2: {2, 3, 4, 5, 6}})

    def test_main_with_two_taps_is_not_reported(self):
        ch = {1: [2], 2: [3, 4], 3: [5], 4: [6], 5: [7], 6: [8]}
        self.assertEqual(runs(1, ch, {7, 8}, {1}), {})

    def test_one_supply_and_one_return_both_reported(self):
        # Each connection counts on its own, though the unit sees 2 terminals.
        ch = {1: [2, 10], 2: [3], 3: [4], 10: [11], 11: [12]}
        self.assertEqual(runs(1, ch, {4, 12}, {1}), {2: {2, 3, 4}, 10: {10, 11, 12}})

    def test_only_the_single_terminal_connection_is_reported(self):
        # Supply main with two taps, return with one grille.
        ch = {1: [2, 10], 2: [3, 4], 3: [5], 4: [6], 10: [11]}
        self.assertEqual(runs(1, ch, {5, 6, 11}, {1}), {10: {10, 11}})

    def test_vav_box_to_one_diffuser_is_not_reported(self):
        # unit 1 -> MP 2 -> boxes 3 and 13, each to one diffuser.
        ch = {1: [2], 2: [3, 13], 3: [4], 4: [5], 13: [14], 14: [15]}
        self.assertEqual(runs(1, ch, {5, 15}, {1, 3, 13}), {})

    def test_single_box_unit_run_stops_at_the_box(self):
        # unit 1 -> MP 2 -> box 3 -> duct 4 -> diffuser 5: the unit's run is
        # reported, but only up to the box.
        ch = {1: [2], 2: [3], 3: [4], 4: [5]}
        self.assertEqual(runs(1, ch, {5}, {1, 3}), {2: {2}})

    def test_picked_duct_root_is_not_reported(self):
        ch = {2: [3], 3: [4]}
        self.assertEqual(runs(2, ch, {4}, set()), {})

    def test_cycle_does_not_hang(self):
        ch = {1: [2], 2: [3], 3: [2, 4]}
        ds = {1: {4}, 2: {4}, 3: {4}, 4: {4}}
        self.assertEqual(RUNS(1, ch, ds, {1}), {2: {2, 3, 4}})


class NoteText(unittest.TestCase):
    def test_note_says_judged_as_a_main(self):
        t = TEXT(101, 202)
        self.assertTrue(t.startswith("NOTE: unit 101 connection at element 202"))
        self.assertIn("judged as a main", t)
        self.assertNotIn("WARNING", t)


if __name__ == "__main__":
    unittest.main()
