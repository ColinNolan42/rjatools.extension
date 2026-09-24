# -*- coding: ascii -*-
"""Tests for water_graph's demand accounting and tree logic.

Runs in plain CPython by stubbing the Revit assemblies, so the WSFU
arithmetic can be checked without opening Revit. This does NOT test the
traversal against a real model - that needs a live Revit session, and the
rigor level of any such test must be stated separately.

Run:  python _test_water_graph.py
"""

import os
import sys
import types
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


# ---------------------------------------------------------------------------
# Stub the Revit API so revit_helpers / water_graph import in CPython.
# ---------------------------------------------------------------------------
def _install_revit_stubs():
    autodesk = types.ModuleType("Autodesk")
    revit = types.ModuleType("Autodesk.Revit")
    db = types.ModuleType("Autodesk.Revit.DB")

    class _Enum(object):
        def __init__(self, name):
            self.name = name

        def __str__(self):
            return self.name

    for name in ("BuiltInParameter", "ElementId", "FilteredElementCollector",
                 "BuiltInCategory", "FlowDirectionType", "MEPSystem",
                 "SpecTypeId",
                 # used by water_drafting
                 "XYZ", "Line", "TextNote", "TextNoteOptions", "TextNoteType",
                 "ViewDrafting", "ViewFamily", "ViewFamilyType", "ViewSheet",
                 "Viewport"):
        setattr(db, name, _Enum(name))

    autodesk.Revit = revit
    revit.DB = db
    sys.modules["Autodesk"] = autodesk
    sys.modules["Autodesk.Revit"] = revit
    sys.modules["Autodesk.Revit.DB"] = db

    pyrevit = types.ModuleType("pyrevit")

    class _HostApp(object):
        version = "2024"

    pyrevit.HOST_APP = _HostApp()
    sys.modules["pyrevit"] = pyrevit


_install_revit_stubs()

import shared_params          # noqa: E402
import water_graph            # noqa: E402


# ---------------------------------------------------------------------------
# Graph builder that bypasses Revit entirely.
# ---------------------------------------------------------------------------
COLD = shared_params.SYSTEM_DOMESTIC_COLD_WATER
HOT = shared_params.SYSTEM_DOMESTIC_HOT_WATER


def make_graph():
    return water_graph.WaterGraph()


def add_node(graph, nid, kind, parent=None, system=COLD, **kw):
    node = water_graph.WaterNode.__new__(water_graph.WaterNode)
    node.element_id = nid
    node.element = None
    node.kind = kind
    node.family_name = "stub"
    node.location_xyz = None
    node.connectors = []
    node.connector_count = 0
    node.parent_by_system = {}
    node.children_by_system = {}
    node.system = system
    node.length_feet = kw.get("length_feet", 0.0)
    node.diameter_inches = 0.0
    node.system_type_id = None
    node.system_type_name = None
    node.is_water_fixture = (kind == water_graph.KIND_FIXTURE)
    node.fixture_name = kw.get("fixture_name", "")
    node.cw_wsfu = kw.get("cw", 0.0)
    node.hw_wsfu = kw.get("hw", 0.0)
    node.total_wsfu = kw.get("total", 0.0)
    node.hwr_active = False
    node.has_cw = True
    node.has_hw = True
    node.served_fixture_ids = kw.get("served", [])
    node.demand_wsfu = 0.0
    node.assigned_size = None
    graph.add(node)
    # Mirror what _make_node does during a real traversal, so the report sees
    # the same fixture register the engine does.
    if kind == water_graph.KIND_FIXTURE and nid not in graph.fixture_ids:
        graph.fixture_ids.append(nid)
    if parent is not None:
        link(graph, parent, nid, system)
    return node


def link(graph, parent_id, child_id, system):
    """Attach child to parent on one system's tree."""
    graph.nodes[child_id].parent_by_system[system] = parent_id
    graph.nodes[parent_id].children_by_system.setdefault(system, []).append(child_id)


class TestDescendants(unittest.TestCase):

    def test_subtree_collected(self):
        g = make_graph()
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1)
        add_node(g, 3, water_graph.KIND_FITTING, parent=2)
        add_node(g, 4, water_graph.KIND_PIPE, parent=3)
        add_node(g, 5, water_graph.KIND_PIPE, parent=3)
        self.assertEqual(sorted(g.descendants(1, COLD)), [2, 3, 4, 5])
        self.assertEqual(sorted(g.descendants(3, COLD)), [4, 5])
        self.assertEqual(g.descendants(4, COLD), [])
        # The hot tree is empty here, and must not borrow the cold tree's links.
        self.assertEqual(g.descendants(1, HOT), [])


class TestColdDemandUsesTotal(unittest.TestCase):
    """One public lavatory (cold 1.5, hot 1.5, total 2.0) fed through a heater.

        origin -> main(2) -> tee(3) --> coldBranch(4) -> lav(5)
                                   \\-> heaterBranch(6) -> heater(7)
        heater(7) -> hotPipe(8) -> lav(5)   [hot tree, lav served by heater]
    """

    def build(self):
        g = make_graph()
        g.origin_id = 1
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1)
        add_node(g, 3, water_graph.KIND_FITTING, parent=2)
        add_node(g, 4, water_graph.KIND_PIPE, parent=3)
        add_node(g, 5, water_graph.KIND_FIXTURE, parent=4,
                 cw=1.5, hw=1.5, total=2.0, fixture_name="L-1")
        add_node(g, 6, water_graph.KIND_PIPE, parent=3)
        add_node(g, 7, water_graph.KIND_HEATER, parent=6, served=[5])
        add_node(g, 8, water_graph.KIND_PIPE, parent=7, system=HOT)
        # The lavatory sits on BOTH trees: its cold connector hangs off the
        # cold branch, its hot connector off the heater's hot pipe. This is the
        # case that a single shared visited set gets wrong.
        link(g, 8, 5, HOT)
        return g

    def test_main_carries_total_once(self):
        g = self.build()
        water_graph._assign_demand(g)
        # The main must see the lavatory exactly once, at TOTAL. If the hot
        # load were added again through the heater it would read 3.5.
        self.assertAlmostEqual(g.nodes[2].demand_wsfu, 2.0)

    def test_cold_branch_to_fixture(self):
        g = self.build()
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[4].demand_wsfu, 2.0)

    def test_cold_branch_to_heater_carries_hot_not_total(self):
        g = self.build()
        water_graph._assign_demand(g)
        # This branch feeds only the heater, so it carries the lavatory's HOT
        # load (1.5), not its TOTAL (2.0). It would be zero without the heater
        # walk, and overstated if it used TOTAL.
        self.assertAlmostEqual(g.nodes[6].demand_wsfu, 1.5)

    def test_hot_pipe_uses_hot(self):
        g = self.build()
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[8].demand_wsfu, 1.5)


class TestMixedFixtures(unittest.TestCase):
    """Cold-only fixtures must not be inflated, hot-only must still count.

        origin -> main(2) -> tee(3) -> wcPipe(4) -> WC(5)      cold only
                                    -> htrPipe(6) -> heater(7) serves DW(9)
                                    heater(7) -> hotPipe(8) -> DW(9)  hot only
    """

    def build(self):
        g = make_graph()
        g.origin_id = 1
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1)
        add_node(g, 3, water_graph.KIND_FITTING, parent=2)
        add_node(g, 4, water_graph.KIND_PIPE, parent=3)
        # public water closet, flush tank: 5.0 cold, no hot, 5.0 total
        add_node(g, 5, water_graph.KIND_FIXTURE, parent=4,
                 cw=5.0, hw=0.0, total=5.0, fixture_name="WC-1")
        add_node(g, 6, water_graph.KIND_PIPE, parent=3)
        add_node(g, 7, water_graph.KIND_HEATER, parent=6, served=[9])
        add_node(g, 8, water_graph.KIND_PIPE, parent=7, system=HOT)
        # dishwashing machine: no cold, 1.4 hot, 1.4 total
        add_node(g, 9, water_graph.KIND_FIXTURE, parent=8, system=HOT,
                 cw=0.0, hw=1.4, total=1.4, fixture_name="DW-1")
        return g

    def test_main_sums_both(self):
        g = self.build()
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[2].demand_wsfu, 6.4)

    def test_cold_only_branch(self):
        g = self.build()
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[4].demand_wsfu, 5.0)

    def test_heater_branch_carries_hot_load(self):
        g = self.build()
        water_graph._assign_demand(g)
        # The dishwasher has no cold side, so its HOT and TOTAL are both 1.4
        # and this branch reads the same either way. Kept as a regression guard.
        self.assertAlmostEqual(g.nodes[6].demand_wsfu, 1.4)

    def test_hot_pipe(self):
        g = self.build()
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[8].demand_wsfu, 1.4)


class TestNoHeater(unittest.TestCase):

    def test_cold_only_system_still_sizes(self):
        g = make_graph()
        g.origin_id = 1
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1)
        add_node(g, 3, water_graph.KIND_FIXTURE, parent=2,
                 cw=5.0, hw=0.0, total=5.0)
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[2].demand_wsfu, 5.0)


class TestTwoFixturesShareAHeater(unittest.TestCase):

    def test_heater_branch_counts_each_fixture_once(self):
        g = make_graph()
        g.origin_id = 1
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1)
        add_node(g, 3, water_graph.KIND_HEATER, parent=2, served=[4, 5])
        add_node(g, 4, water_graph.KIND_FIXTURE, parent=3, system=HOT,
                 cw=1.5, hw=1.5, total=2.0)
        add_node(g, 5, water_graph.KIND_FIXTURE, parent=3, system=HOT,
                 cw=1.0, hw=1.0, total=1.4)
        water_graph._assign_demand(g)
        # Pipe 2 feeds only the heater. Both fixtures are reached through it,
        # each counted once at HOT: 1.5 + 1.0. Not their totals (3.4), and not
        # doubled.
        self.assertAlmostEqual(g.nodes[2].demand_wsfu, 2.5)


class TestReturnSystemDetection(unittest.TestCase):
    """Hot supply and hot recirculation share one CLASSIFICATION but not one
    System Type. These check the return is worked out from the model."""

    def _hot_pipe(self, g, nid, parent, type_id, type_name):
        node = add_node(g, nid, water_graph.KIND_PIPE, parent=parent,
                        system=HOT)
        node.system_type_id = type_id
        node.system_type_name = type_name
        g.system_types[type_id] = type_name
        return node

    def test_pump_decides_even_when_the_name_says_nothing(self):
        """The strongest signal is topological, so a project that names its
        systems 'HW-1' and 'HW-2' still gets the right answer."""
        g = make_graph()
        add_node(g, 1, water_graph.KIND_HEATER, system=HOT)
        self._hot_pipe(g, 2, 1, 100, "HW-1")
        self._hot_pipe(g, 3, 1, 200, "HW-2")
        add_node(g, 4, water_graph.KIND_PUMP, parent=3, system=HOT)
        g.pump_ids.append(4)

        found = water_graph.detect_return_system_types(g)
        self.assertEqual(found["detected"], set([200]))
        self.assertTrue(found["certain"])
        picked = [c for c in found["candidates"] if c["id"] == 200][0]
        self.assertIn("pump", " ".join(picked["reasons"]))

    def test_pump_reached_through_a_fitting(self):
        """A pump connects through fittings, so the walk has to step past
        them to find the pipe that carries the System Type."""
        g = make_graph()
        add_node(g, 1, water_graph.KIND_HEATER, system=HOT)
        self._hot_pipe(g, 2, 1, 100, "Domestic Hot Water")
        add_node(g, 3, water_graph.KIND_FITTING, parent=1, system=HOT)
        self._hot_pipe(g, 4, 3, 200, "Domestic Hot Water Recirculation")
        add_node(g, 5, water_graph.KIND_FITTING, parent=4, system=HOT)
        add_node(g, 6, water_graph.KIND_PUMP, parent=5, system=HOT)
        g.pump_ids.append(6)

        found = water_graph.detect_return_system_types(g)
        self.assertEqual(found["detected"], set([200]))

    def test_name_decides_when_there_is_no_pump(self):
        g = make_graph()
        add_node(g, 1, water_graph.KIND_HEATER, system=HOT)
        self._hot_pipe(g, 2, 1, 100, "Domestic Hot Water")
        self._hot_pipe(g, 3, 1, 200, "Domestic Hot Water Recirculation")

        found = water_graph.detect_return_system_types(g)
        self.assertEqual(found["detected"], set([200]))
        # Nothing topological spoke, so the tool must not claim certainty.
        self.assertFalse(found["certain"])

    def test_pump_beats_a_misleading_name(self):
        """If the names and the pump disagree, topology wins."""
        g = make_graph()
        add_node(g, 1, water_graph.KIND_HEATER, system=HOT)
        self._hot_pipe(g, 2, 1, 100, "HWR Main")     # named like a return
        self._hot_pipe(g, 3, 1, 200, "Hot Water")
        add_node(g, 4, water_graph.KIND_PUMP, parent=3, system=HOT)
        g.pump_ids.append(4)

        found = water_graph.detect_return_system_types(g)
        self.assertTrue(found["certain"])
        self.assertIn(200, found["detected"])

    def test_single_hot_system_is_never_called_a_return(self):
        """One hot System Type and no pump means a job with no recirculation.
        Marking it a return would leave the hot water unsized."""
        g = make_graph()
        add_node(g, 1, water_graph.KIND_HEATER, system=HOT)
        self._hot_pipe(g, 2, 1, 100, "Domestic Hot Water")
        self._hot_pipe(g, 3, 1, 100, "Domestic Hot Water")

        found = water_graph.detect_return_system_types(g)
        self.assertEqual(found["detected"], set())

    def test_minority_guess_is_flagged_as_unconfirmed(self):
        g = make_graph()
        add_node(g, 1, water_graph.KIND_HEATER, system=HOT)
        for nid in range(2, 12):
            self._hot_pipe(g, nid, 1, 100, "HW A")
        self._hot_pipe(g, 20, 1, 200, "HW B")

        found = water_graph.detect_return_system_types(g)
        self.assertEqual(found["detected"], set([200]))
        self.assertFalse(found["certain"])
        picked = [c for c in found["candidates"] if c["id"] == 200][0]
        self.assertIn("NOTHING CONFIRMED THIS", " ".join(picked["reasons"]))

    def test_no_hot_piping_detects_nothing(self):
        g = make_graph()
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1)
        found = water_graph.detect_return_system_types(g)
        self.assertEqual(found["detected"], set())
        self.assertEqual(found["candidates"], [])

    def test_cold_piping_is_never_a_return_candidate(self):
        g = make_graph()
        add_node(g, 1, water_graph.KIND_ORIGIN)
        cold = add_node(g, 2, water_graph.KIND_PIPE, parent=1, system=COLD)
        cold.system_type_id = 900
        cold.system_type_name = "Domestic Cold Water Recirculation"
        found = water_graph.detect_return_system_types(g)
        self.assertEqual(found["candidates"], [])


def _run():
    suite = unittest.TestLoader().loadTestsFromModule(sys.modules[__name__])
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    print("=" * 70)
    print("SUMMARY: {} run, {} passed, {} failed, {} errors".format(
        result.testsRun,
        result.testsRun - len(result.failures) - len(result.errors),
        len(result.failures), len(result.errors)))
    print("=" * 70)
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    sys.exit(_run())
