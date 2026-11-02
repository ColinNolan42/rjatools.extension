# -*- coding: ascii -*-
"""Tests for the main / branch rule (hvac_graph.split_above_ids).

RJA rule (Colin, 2026-11-02): a duct with one terminal downstream is a
BRANCH only when the run has split above it, at ductwork (never equipment)
with two or more terminals downstream. A unit feeding a single terminal
never splits, so its whole run stays a MAIN and keeps the velocity and
friction check. The pure block of hvac_graph.py is run on its own, since
hvac_graph itself imports the Revit API.

Run:  python _test_branch_split.py
"""

import io
import os
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))


def _load():
    src = io.open(os.path.join(HERE, "hvac_graph.py"), encoding="utf-8").read()
    start = src.index("def split_above_ids(")
    end = src.index("def solid_fill_pattern_id(")
    ns = {}
    exec(src[start:end], ns)
    return ns["split_above_ids"]


SPLIT_ABOVE = _load()


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


def branches(root, children, terminals, equipment, ducts):
    """The ducts the tool calls BRANCH: one terminal below and a split above."""
    ds = downstream(root, children, terminals)
    above = SPLIT_ABOVE(root, children, ds, equipment)
    return set(d for d in ducts if len(ds.get(d, ())) == 1 and d in above)


class TestBranchSplit(unittest.TestCase):

    def test_unit_feeding_one_diffuser_is_all_main(self):
        # unit 1 -> duct 2 -> elbow 3 -> duct 4 -> flex 5 -> diffuser 6
        ch = {1: [2], 2: [3], 3: [4], 4: [5], 5: [6]}
        self.assertEqual(branches(1, ch, {6}, {1}, {2, 4, 5}), set())

    def test_main_with_two_taps_has_two_branches(self):
        # unit 1 -> main 2, taps 3 and 4 -> branch ducts 5, 6 -> diffusers 7, 8
        ch = {1: [2], 2: [3, 4], 3: [5], 4: [6], 5: [7], 6: [8]}
        self.assertEqual(branches(1, ch, {7, 8}, {1}, {2, 5, 6}), {5, 6})

    def test_unit_with_one_supply_and_one_return_is_all_main(self):
        # The root unit's count mixes supply and return (2), but equipment
        # is never the split: both single runs stay mains.
        ch = {1: [2, 10], 2: [3], 3: [4], 10: [11], 11: [12]}
        self.assertEqual(branches(1, ch, {4, 12}, {1}, {2, 3, 10, 11}), set())

    def test_picked_duct_root_counts_as_the_split(self):
        # Picked supply duct 2 (no unit) with two taps.
        ch = {2: [3, 4], 3: [5], 4: [6], 5: [7], 6: [8]}
        self.assertEqual(branches(2, ch, {7, 8}, set(), {2, 5, 6}), {5, 6})

    def test_vav_box_to_one_diffuser_stays_main(self):
        # unit 1 -> MP duct 2 -> box 3 -> duct 4 -> diffuser 5
        ch = {1: [2], 2: [3], 3: [4], 4: [5]}
        self.assertEqual(branches(1, ch, {5}, {1, 3}, {2, 4}), set())

    def test_box_outlet_split_in_ductwork_makes_branches(self):
        # unit 1 -> duct 2 -> box 3 -> duct 4 -> tee 5 -> ducts 6, 7 -> 8, 9
        ch = {1: [2], 2: [3], 3: [4], 4: [5], 5: [6, 7], 6: [8], 7: [9]}
        self.assertEqual(branches(1, ch, {8, 9}, {1, 3}, {2, 4, 6, 7}), {6, 7})

    def test_split_after_a_long_single_run(self):
        # The run splits only at the end; the ducts before it are mains,
        # the two after it branches.
        ch = {1: [2], 2: [3], 3: [4], 4: [5, 6], 5: [7], 6: [8]}
        self.assertEqual(branches(1, ch, {7, 8}, {1}, {2, 3, 5, 6}), {5, 6})

    def test_cycle_does_not_hang(self):
        ch = {1: [2], 2: [3], 3: [2, 4], 4: [5], 5: [6]}
        ds = {1: {6}, 2: {6}, 3: {6}, 4: {6}, 5: {6}, 6: {6}}
        self.assertEqual(SPLIT_ABOVE(1, ch, ds, {1}), set())


if __name__ == "__main__":
    unittest.main()
