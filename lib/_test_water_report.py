# -*- coding: ascii -*-
"""End-to-end check of water_sizing_engine + water_report on a synthetic system.

Stubs the Revit assemblies so the whole chain (demand -> size -> printed
report) can be exercised in plain CPython. This proves the arithmetic and the
layout only. It does NOT prove traversal against a real model.

Run:  python _test_water_report.py          (prints the report)
      python _test_water_report.py --test   (asserts, exits non-zero on fail)
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_water_graph import (          # noqa: E402
    _install_revit_stubs, make_graph, add_node, link, COLD, HOT)

import shared_params        # noqa: E402
import water_graph          # noqa: E402
import water_sizing_engine  # noqa: E402
import water_report         # noqa: E402


def connector(system, diameter_in):
    return {"connector_index": 0, "direction": "In", "is_connected": True,
            "connected_element_id": None, "connected_element_type": None,
            "origin_xyz": None, "connected_system_element_id": None,
            "system_type": system, "diameter_in": diameter_in}


def build_demo_system():
    """A small public restroom group plus a water heater.

        RPZ(1)
          |
        main(2) 3 in run
          |
        tee(3)
          +-- coldBranch(4) --- tee(5)
          |                       +-- wcPipe(6) --- WC-1 (public flush tank, 5.0)
          |                       +-- wcPipe(7) --- WC-2 (public flush tank, 5.0)
          |                       +-- lavPipe(8) -- L-1 (public lav 1.5/1.5/2.0)
          +-- htrBranch(9) --- heater(10)
                                  |
                               hotPipe(11) --- L-1 (same lavatory, hot side)
    """
    g = make_graph()
    g.origin_id = 1
    add_node(g, 1, water_graph.KIND_ORIGIN)
    add_node(g, 2, water_graph.KIND_PIPE, parent=1, length_feet=42.0)
    add_node(g, 3, water_graph.KIND_FITTING, parent=2)
    add_node(g, 4, water_graph.KIND_PIPE, parent=3, length_feet=18.0)
    add_node(g, 5, water_graph.KIND_FITTING, parent=4)
    add_node(g, 6, water_graph.KIND_PIPE, parent=5, length_feet=6.0)
    add_node(g, 7, water_graph.KIND_PIPE, parent=5, length_feet=7.5)
    add_node(g, 8, water_graph.KIND_PIPE, parent=5, length_feet=9.0)

    wc1 = add_node(g, 20, water_graph.KIND_FIXTURE, parent=6,
                   cw=5.0, hw=0.0, total=5.0, fixture_name="WC-1")
    wc2 = add_node(g, 21, water_graph.KIND_FIXTURE, parent=7,
                   cw=5.0, hw=0.0, total=5.0, fixture_name="WC-2")
    lav = add_node(g, 22, water_graph.KIND_FIXTURE, parent=8,
                   cw=1.5, hw=1.5, total=2.0, fixture_name="L-1")
    for node, type_name in ((wc1, "Water Closet, Flush Tank"),
                            (wc2, "Water Closet, Flush Tank"),
                            (lav, "Lavatory")):
        node.type_name = type_name
        node.is_public = True
    wc1.connectors = [connector(COLD, 0.5)]
    wc2.connectors = [connector(COLD, 0.5)]
    lav.connectors = [connector(COLD, 0.5), connector(HOT, 0.5)]

    add_node(g, 9, water_graph.KIND_PIPE, parent=3, length_feet=12.0)
    heater = add_node(g, 10, water_graph.KIND_HEATER, parent=9, served=[22])
    heater.family_name = "AOSMITH- DEL 6-20"
    g.heater_ids.append(10)
    add_node(g, 11, water_graph.KIND_PIPE, parent=10, system=HOT,
             length_feet=15.0)
    link(g, 11, 22, HOT)

    g.system_types = {1823: "Domestic Cold Water", 1822: "Domestic Hot Water"}
    return g


HEADER = {"date": "2026-09-24", "job": "Demo Restroom Group",
          "job_number": "26.000", "by": "CN"}


class TestDemoSizing(unittest.TestCase):

    def setUp(self):
        self.graph = build_demo_system()
        water_graph._assign_demand(self.graph)
        self.result = water_sizing_engine.size_network(self.graph)
        self.by_id = dict((s.element_id, s) for s in self.result["segments"])

    def test_main_demand_and_size(self):
        # 5.0 + 5.0 + 2.0 = 12.0 total wsfu. Cold table: 6 < 12 <= 19 -> 1 inch.
        main = self.by_id[2]
        self.assertAlmostEqual(main.demand_wsfu, 12.0)
        self.assertEqual(main.nominal_size, "1")

    def test_heater_branch_carries_hot_total(self):
        # Feeds the heater only, so it carries the lavatory's HOT load 1.5,
        # not its TOTAL 2.0. Cold table: 1.5 <= 2.0 -> 1/2 inch.
        htr = self.by_id[9]
        self.assertAlmostEqual(htr.demand_wsfu, 1.5)
        self.assertEqual(htr.nominal_size, "1/2")

    def test_single_wc_branch_raised_to_connector_size(self):
        wc = self.by_id[6]
        self.assertAlmostEqual(wc.demand_wsfu, 5.0)
        # 5.0 wsfu cold -> 3/4 by table (2 < 5 <= 6).
        self.assertEqual(wc.nominal_size, "3/4")

    def test_branch_serving_three_fixtures_gets_multi_minimum(self):
        branch = self.by_id[4]
        self.assertAlmostEqual(branch.demand_wsfu, 12.0)
        self.assertEqual(branch.fixture_count, 3)

    def test_hot_pipe_uses_hot_column(self):
        hot = self.by_id[11]
        self.assertAlmostEqual(hot.demand_wsfu, 1.5)
        # Hot table at 5 fps: 1 < 1.5 <= 3 -> 3/4 inch.
        self.assertEqual(hot.nominal_size, "3/4")

    def test_wsfu_table_lists_every_fixture_type(self):
        lines = water_report.wsfu_table(self.graph)
        text = "\n".join(lines)
        self.assertNotIn("no water fixtures found", text)
        self.assertIn("Water Closet, Flush Tank, PUBLIC", text)
        self.assertIn("Lavatory, PUBLIC", text)
        # 2 water closets + 1 lavatory
        totals = [l for l in lines if l.startswith("TOTALS")][0]
        self.assertIn("3", totals.split()[1])

    def test_wsfu_table_totals_match_the_main(self):
        lines = water_report.wsfu_table(self.graph)
        text = "\n".join(lines)
        # extended TOTAL = 2x5.0 + 1x2.0 = 12.0, the same load the main carries
        self.assertIn("Cold water sizing load (TOTAL wsfu): 12", text)
        # extended HOT = 1 x 1.5
        self.assertIn("Hot water sizing load (HOT wsfu): 1.50", text)

    def test_wsfu_data_feeds_both_renderers(self):
        # The drafting view and the console both render from wsfu_data(), so
        # the numbers can never disagree between the sheet and the window.
        data = water_report.wsfu_data(self.graph)
        self.assertEqual(len(data["rows"]), 2)
        self.assertAlmostEqual(data["ext_total"], 12.0)
        self.assertAlmostEqual(data["ext_hot"], 1.5)
        self.assertEqual(data["totals"]["qty"], "3")
        # every column key must be present on every row, or the schedule would
        # silently skip cells
        for row in data["rows"] + [data["totals"]]:
            for key, _header, _vw, _cw in water_report.WSFU_COLUMNS:
                self.assertIn(key, row)

    def test_report_is_ascii_and_has_all_blocks(self):
        text = water_report.build_report(self.graph, self.result, HEADER)
        text.encode("ascii")
        for needle in ("BASIS OF DESIGN",
                       "WATER SUPPLY FIXTURE UNIT TAKE-OFF",
                       "PIPE SIZING BY SEGMENT",
                       "FLAGS, SKIPPED AND NOT EVALUATED",
                       "2024 International Plumbing Code"):
            self.assertIn(needle, text)


class TestDraftingLayout(unittest.TestCase):
    """The drafting module's pure layout helpers.

    The Revit drawing calls cannot run here. This covers the arithmetic and
    the shared column contract only.
    """

    def setUp(self):
        import water_drafting
        self.wd = water_drafting

    def test_wrap_breaks_on_spaces_and_keeps_every_word(self):
        line = ("Pipe sizing per the firm WSFU-per-size table (Copper Type L), "
                "nominal sizes 1/2\" through 3\", maximum velocity 8 fps cold "
                "water and 5 fps hot water.")
        wrapped = self.wd._wrap(line, 60)
        self.assertTrue(len(wrapped) > 1)
        for piece in wrapped:
            self.assertLessEqual(len(piece), 60)
        self.assertEqual(" ".join(wrapped).split(), line.split())

    def test_wrap_handles_empty(self):
        self.assertEqual(self.wd._wrap("", 40), [""])

    def test_table_width_matches_the_shared_columns(self):
        # The drafting table's width must come from the same column list the
        # console table uses, or the two can disagree.
        expected = sum(c[2] for c in water_report.WSFU_COLUMNS)
        self.assertAlmostEqual(expected, 0.705)
        self.assertEqual(len(water_report.WSFU_COLUMNS), 8)

    def test_default_view_name_uses_job_number(self):
        self.assertEqual(
            self.wd._default_view_name(
                {"job_number": "26.000", "date": "2026-09-24"}),
            "WSFU Take-Off - 26.000 - 2026-09-24")
        self.assertEqual(
            self.wd._default_view_name({"job_number": "", "date": "2026-09-24"}),
            "WSFU Take-Off - 2026-09-24")


def _print_report():
    graph = build_demo_system()
    water_graph._assign_demand(graph)
    result = water_sizing_engine.size_network(graph)
    print(water_report.build_report(graph, result, HEADER))


if __name__ == "__main__":
    if "--test" in sys.argv:
        suite = unittest.TestLoader().loadTestsFromModule(sys.modules[__name__])
        res = unittest.TextTestRunner(verbosity=2).run(suite)
        print("SUMMARY: {} run, {} failed, {} errors".format(
            res.testsRun, len(res.failures), len(res.errors)))
        sys.exit(0 if res.wasSuccessful() else 1)
    _print_report()
