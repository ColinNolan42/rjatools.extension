# -*- coding: ascii -*-
"""Tests for ifgc_gas_sizing_tables.json against the 2024 IFGC PDF.

Every cell in PDF_CELLS was read from 2024_IFGC_1.pdf, Tables 402.4(1) to
(37) (PDF pages 132 to 177), not from the JSON: the first, middle and last
column of the first, 100 ft (or middle) and last length row of each table.
The 10 ft rows of 402.4(9), (25) and (37) and the 2,000 ft row of (36) were
also checked on the page images. A later edit to the JSON that moves a
value, drops a column or changes a unit fails here.

Also covers the propane units (402.4(25) to (36) publish MBH, every other
table CFH) and sizing on an MBH table, where heat content shall have no
effect (Colin, 2026-11-04).

Run:  python -I _test_gas_tables.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_prv_zones import _node, _edge, _finish   # installs the stubs

import gas_tables      # noqa: E402
import pipe_graph      # noqa: E402
import sizing_engine   # noqa: E402


# table id: (number of size columns, [(length row, column index, value)])
# None is "NA" in the published table.
PDF_CELLS = {
    "402.4(1)": (14, [(10, 0, 131), (10, 7, 8580), (10, 13, 303000), (100, 0, 38), (100, 7, 2470), (100, 13, 87200), (2000, 0, None), (2000, 7, 488), (2000, 13, 17200)]),
    "402.4(2)": (14, [(10, 0, 172), (10, 7, 11300), (10, 13, 399000), (100, 0, 50), (100, 7, 3260), (100, 13, 115000), (2000, 0, None), (2000, 7, 644), (2000, 13, 22700)]),
    "402.4(3)": (9, [(10, 0, 454), (10, 4, 5500), (10, 8, 60800), (100, 0, 131), (100, 4, 1580), (100, 8, 17500), (2000, 0, 26), (2000, 4, 313), (2000, 8, 3460)]),
    "402.4(4)": (9, [(10, 0, 660), (10, 4, 8000), (10, 8, 88500), (100, 0, 190), (100, 4, 2300), (100, 8, 25500), (2000, 0, 38), (2000, 4, 455), (2000, 8, 5040)]),
    "402.4(5)": (9, [(10, 0, 1510), (10, 4, 17100), (10, 8, 189000), (100, 0, 462), (100, 4, 5260), (100, 8, 58200), (2000, 0, 91), (2000, 4, 1040), (2000, 8, 11500)]),
    "402.4(6)": (9, [(10, 0, 2350), (10, 4, 28500), (10, 8, 316000), (100, 0, 677), (100, 4, 8200), (100, 8, 90800), (2000, 0, 134), (2000, 4, 1620), (2000, 8, 18000)]),
    "402.4(7)": (9, [(10, 0, 3190), (10, 4, 36200), (10, 8, 401000), (100, 0, 979), (100, 4, 11100), (100, 8, 123000), (2000, 0, 181), (2000, 4, 2200), (2000, 8, 24300)]),
    "402.4(8)": (9, [(10, 0, 20), (10, 4, 210), (10, 8, 2650), (100, 0, None), (100, 4, 60), (100, 8, 761), (2000, 0, None), (2000, 4, 12), (2000, 8, 151)]),
    "402.4(9)": (9, [(10, 0, 27), (10, 4, 276), (10, 8, 3490), (100, 0, None), (100, 4, 79), (100, 8, 1000), (2000, 0, None), (2000, 4, 16), (2000, 8, 199)]),
    "402.4(10)": (9, [(10, 0, 39), (10, 4, 402), (10, 8, 5080), (100, 0, 11), (100, 4, 116), (100, 8, 1460), (2000, 0, None), (2000, 4, 23), (2000, 8, 289)]),
    "402.4(11)": (9, [(10, 0, 190), (10, 4, 1970), (10, 8, 24900), (100, 0, 55), (100, 4, 568), (100, 8, 7170), (2000, 0, 11), (2000, 4, 112), (2000, 8, 1420)]),
    "402.4(12)": (9, [(10, 0, 245), (10, 4, 2550), (10, 8, 32200), (100, 0, 71), (100, 4, 735), (100, 8, 9280), (2000, 0, 14), (2000, 4, 145), (2000, 8, 1830)]),
    "402.4(13)": (9, [(10, 0, 303), (10, 4, 3150), (10, 8, 39800), (100, 0, 87), (100, 4, 907), (100, 8, 11500), (2000, 0, 17), (2000, 4, 179), (2000, 8, 2270)]),
    "402.4(14)": (9, [(10, 0, 511), (10, 4, 5320), (10, 8, 67100), (100, 0, 147), (100, 4, 1530), (100, 8, 19300), (2000, 0, 29), (2000, 4, 302), (2000, 8, 3820)]),
    "402.4(15)": (14, [(5, 0, 46), (5, 7, 546), (5, 13, 4140), (100, 0, 9), (100, 7, 118), (100, 13, 933), (300, 0, 5), (300, 7, 67), (300, 13, 540)]),
    "402.4(16)": (14, [(5, 0, 120), (5, 7, 1370), (5, 13, 10100), (100, 0, 24), (100, 7, 295), (100, 13, 2280), (300, 0, 13), (300, 7, 168), (300, 13, 1320)]),
    "402.4(17)": (14, [(5, 0, 173), (5, 7, 1950), (5, 13, 14260), (100, 0, 35), (100, 7, 421), (100, 13, 3210), (300, 0, 19), (300, 7, 240), (300, 13, 1860)]),
    "402.4(18)": (14, [(10, 0, 270), (10, 7, 2990), (10, 13, 21600), (100, 0, 79), (100, 7, 920), (100, 13, 6880), (500, 0, 34), (500, 7, 404), (500, 13, 3090)]),
    "402.4(19)": (14, [(10, 0, 523), (10, 7, 5660), (10, 13, 40400), (100, 0, 154), (100, 7, 1740), (100, 13, 12800), (500, 0, 66), (500, 7, 766), (500, 13, 5760)]),
    "402.4(20)": (8, [(10, 0, 153), (10, 4, 1440), (10, 7, 13900), (100, 0, 44), (100, 4, 415), (100, 7, 3990), (500, 0, 18), (500, 4, 174), (500, 7, 1670)]),
    "402.4(21)": (8, [(10, 0, 201), (10, 4, 1900), (10, 7, 18260), (100, 0, 58), (100, 4, 547), (100, 7, 5250), (500, 0, 24), (500, 4, 229), (500, 7, 2200)]),
    "402.4(22)": (8, [(10, 0, 1860), (10, 4, 17600), (10, 7, 169000), (100, 0, 535), (100, 4, 5050), (100, 7, 48600), (2000, 0, 106), (2000, 4, 1000), (2000, 7, 9600)]),
    "402.4(23)": (2, [(10, 0, 54), (10, 1, 372), (100, 0, 16), (100, 1, 107), (500, 0, None), (500, 1, 45)]),
    "402.4(24)": (2, [(10, 0, 72), (10, 1, 490), (100, 0, 21), (100, 1, 141), (500, 0, None), (500, 1, 59)]),
    "402.4(25)": (9, [(10, 0, 3320), (10, 4, 40300), (10, 8, 446000), (100, 0, 956), (100, 4, 11600), (100, 8, 128000), (2000, 0, 189), (2000, 4, 2290), (2000, 8, 25400)]),
    "402.4(26)": (9, [(10, 0, 5890), (10, 4, 71300), (10, 8, 789000), (100, 0, 1690), (100, 4, 20500), (100, 8, 227000), (2000, 0, 335), (2000, 4, 4060), (2000, 8, 44900)]),
    "402.4(27)": (9, [(10, 0, 2680), (10, 4, 32400), (10, 8, 359000), (100, 0, 770), (100, 4, 9330), (100, 8, 103000), (2000, 0, 152), (2000, 4, 1840), (2000, 8, 20400)]),
    "402.4(28)": (9, [(10, 0, 291), (10, 4, 3520), (10, 8, 39000), (100, 0, 84), (100, 4, 1010), (100, 8, 11200), (2000, 0, 17), (2000, 4, 200), (2000, 8, 2220)]),
    "402.4(29)": (9, [(10, 0, 513), (10, 4, 5330), (10, 8, 67400), (100, 0, 147), (100, 4, 1540), (100, 8, 19400), (2000, 0, 29), (2000, 4, 304), (2000, 8, 3830)]),
    "402.4(30)": (9, [(10, 0, 45), (10, 4, 467), (10, 8, 5890), (100, 0, 13), (100, 4, 134), (100, 8, 1700), (2000, 0, None), (2000, 4, 27), (2000, 8, 335)]),
    "402.4(31)": (9, [(10, 0, 413), (10, 4, 4300), (10, 8, 54200), (100, 0, 119), (100, 4, 1240), (100, 8, 15600), (2000, 0, 23), (2000, 4, 244), (2000, 8, 3090)]),
    "402.4(32)": (14, [(5, 0, 72), (5, 7, 863), (5, 13, 6550), (100, 0, 14), (100, 7, 186), (100, 13, 1480), (300, 0, 8), (300, 7, 107), (300, 13, 854)]),
    "402.4(33)": (14, [(10, 0, 426), (10, 7, 4720), (10, 13, 34200), (100, 0, 124), (100, 7, 1450), (100, 13, 10900), (500, 0, 53), (500, 7, 638), (500, 13, 4880)]),
    "402.4(34)": (14, [(10, 0, 826), (10, 7, 8950), (10, 13, 63800), (100, 0, 243), (100, 7, 2760), (100, 13, 20300), (500, 0, 104), (500, 7, 1210), (500, 13, 9110)]),
    "402.4(35)": (8, [(10, 0, 340), (10, 4, 3210), (10, 7, 30900), (100, 0, 98), (100, 4, 924), (100, 7, 8900), (500, 0, 41), (500, 4, 387), (500, 7, 3720)]),
    "402.4(36)": (8, [(10, 0, 3130), (10, 4, 29500), (10, 7, 284000), (100, 0, 899), (100, 4, 8500), (100, 7, 81700), (2000, 0, 178), (2000, 4, 1680), (2000, 7, 16200)]),
    "402.4(37)": (2, [(10, 0, 121), (10, 1, 828), (100, 0, 35), (100, 1, 238), (500, 0, 15), (500, 1, 100)]),
}

COPPER_SIZES = ["1/4 K&L", "3/8 K&L", "1/2 K&L", "5/8 K&L", "3/4 K&L",
                "1 K&L", "1-1/4 K&L", "1-1/2 K&L", "2 K&L"]
COPPER_IDS = [0.305, 0.402, 0.527, 0.652, 0.745, 0.995, 1.245, 1.481, 1.959]
COPPER_TABLES = ["402.4(%d)" % n for n in (8, 9, 10, 11, 12, 13, 14,
                                            29, 30, 31)]
MBH_TABLES = ["402.4(%d)" % n for n in range(25, 37)]


def _row_lengths(table_id):
    return sorted(int(k) for k in gas_tables.get_table(table_id)["data"])


class TestCellsAgainstPdf(unittest.TestCase):

    def test_all_37_tables_present(self):
        self.assertEqual(len(PDF_CELLS), 37)
        for tid in PDF_CELLS:
            gas_tables.get_table(tid)

    def test_column_counts(self):
        for tid, (ncols, _cells) in PDF_CELLS.items():
            t = gas_tables.get_table(tid)
            self.assertEqual(len(t["pipe_sizes_nominal"]), ncols, tid)
            # CSST tables carry EHD flow designations instead of an ID list.
            ids = t.get("pipe_sizes_actual_id", t.get("ehd_flow_designations"))
            self.assertEqual(len(ids), ncols, tid)
            for length, row in t["data"].items():
                self.assertEqual(len(row), ncols, "%s row %s" % (tid, length))

    def test_cells(self):
        for tid, (_ncols, cells) in PDF_CELLS.items():
            data = gas_tables.get_table(tid)["data"]
            for length, col, value in cells:
                self.assertEqual(data[str(length)][col], value,
                                 "%s %d ft column %d" % (tid, length, col))


class TestSizeLabels(unittest.TestCase):

    def test_copper_labels_and_ids(self):
        for tid in COPPER_TABLES:
            t = gas_tables.get_table(tid)
            self.assertEqual(t["pipe_sizes_nominal"], COPPER_SIZES, tid)
            self.assertEqual(t["pipe_sizes_actual_id"], COPPER_IDS, tid)
            self.assertEqual(t["pipe_sizes_acr"][:7],
                             ["3/8", "1/2", "5/8", "3/4", "7/8", "1-1/8",
                              "1-3/8"], tid)
            self.assertEqual(t["pipe_sizes_acr"][7:], [None, None], tid)

    def test_no_acr_size_name(self):
        self.assertNotIn("1-5/8 ACR", sizing_engine.NOMINAL_TO_INCHES)
        for tid in COPPER_TABLES:
            self.assertNotIn("1-5/8 ACR", gas_tables.list_pipe_sizes(tid))

    def test_cts_tubing(self):
        for tid, second in (("402.4(23)", "3/4 CTS"), ("402.4(24)", "3/4 CTS"),
                            ("402.4(37)", "1 CTS")):
            t = gas_tables.get_table(tid)
            self.assertEqual(t["pipe_sizes_nominal"], ["1/2 CTS", second], tid)
            self.assertEqual(t["pipe_sizes_designation"], ["SDR 7", "SDR 11"])
            self.assertEqual(t["pipe_sizes_actual_id"], [0.445, 0.927], tid)

    def test_no_phantom_sizes(self):
        for tid in ("402.4(3)", "402.4(25)", "402.4(26)", "402.4(27)"):
            self.assertEqual(gas_tables.list_pipe_sizes(tid)[-1], "4", tid)

    def test_every_size_has_inches(self):
        for tid in PDF_CELLS:
            for s in gas_tables.list_pipe_sizes(tid):
                self.assertIn(s, sizing_engine.NOMINAL_TO_INCHES,
                              "%s size %s" % (tid, s))

    # Cells the 2024 IFGC prints out of sequence, kept exactly as printed.
    # 402.4(36), 400 ft, 3 in.: the page image (PDF page 175, printed 4-52)
    # reads 12,000 MBH between 21,500 at 350 ft and 18,700 at 450 ft. Very
    # likely a misprint; the printed value is the lower, conservative one.
    PRINTED_OUT_OF_SEQUENCE = {("402.4(36)", 400, 6): 12000}

    def test_capacities_fall_with_length(self):
        for (tid, length, col), value in self.PRINTED_OUT_OF_SEQUENCE.items():
            self.assertEqual(
                gas_tables.get_table(tid)["data"][str(length)][col], value)
        for tid in PDF_CELLS:
            data = gas_tables.get_table(tid)["data"]
            rows = _row_lengths(tid)
            for col in range(len(data[str(rows[0])])):
                vals = [data[str(r)][col] for r in rows
                        if data[str(r)][col] is not None
                        and (tid, r, col) not in self.PRINTED_OUT_OF_SEQUENCE]
                self.assertEqual(vals, sorted(vals, reverse=True),
                                 "%s column %d" % (tid, col))


class TestUnits(unittest.TestCase):

    def test_units(self):
        for tid in PDF_CELLS:
            expected = "MBH" if tid in MBH_TABLES else "CFH"
            self.assertEqual(gas_tables.table_units(tid), expected, tid)

    def test_demand_in_table_units(self):
        self.assertEqual(sizing_engine.demand_in_table_units(
            100.0, "402.4(28)", 840.0), (100.0, "MBH"))
        v, u = sizing_engine.demand_in_table_units(84.0, "402.4(2)", 840.0)
        self.assertEqual(u, "CFH")
        self.assertAlmostEqual(v, 100.0)


def _two_fixture_graph(load_a, load_b, run_ft):
    g = pipe_graph.NetworkGraph()
    g.origin_id = 1
    _node(g, 1, "meter")
    _node(g, 2, "tee")
    _node(g, 3, "fixture", load_a, "A")
    _node(g, 4, "fixture", load_b, "B")
    _edge(g, 10, 1, 2, run_ft - 10.0)
    _edge(g, 11, 2, 3, 10.0)
    _edge(g, 12, 2, 4, 5.0)
    _finish(g)
    return g


class TestSizingOnMbhTable(unittest.TestCase):
    """402.4(28) (propane, Sch 40, 11.0 in. w.c., 0.5 in. w.c. drop), 100 ft
    row per the PDF: 1/2 84, 3/4 175, 1 330, 1-1/4 677 MBH."""

    def _size(self, heat):
        g = _two_fixture_graph(150.0, 60.0, 100.0)
        return g, sizing_engine.size_system(
            g, "Schedule 40 Steel", 0.7, "402.4(28)",
            heat_content_btu_per_cf=heat)

    def test_heat_content_has_no_effect(self):
        _g, a = self._size(840.0)
        _g, b = self._size(2516.0)
        _g, c = self._size(1000.0)
        self.assertEqual(a["sizes"], b["sizes"])
        self.assertEqual(a["sizes"], c["sizes"])

    def test_sizes_compare_mbh(self):
        g, r = self._size(840.0)
        # Main carries 210 MBH: 3/4 holds 175, 1 holds 330.
        self.assertEqual(r["sizes"][10], "1")
        # A 150 MBH: 3/4 (175). B 60 MBH: 1/2 (84) upsized to the minimum.
        self.assertEqual(r["sizes"][11], "3/4")
        self.assertEqual(r["sizes"][12], "3/4")
        d = [x for x in r["segment_detail"] if x["pipe_id"] == 10][0]
        self.assertEqual(d["units"], "MBH")
        self.assertEqual(d["demand_cfh"], 210.0)
        out = sizing_engine.format_sizing_output(r, g)
        self.assertIn("Capacity in:   MBH", out)
        self.assertIn("Heat Content:  not used", out)
        self.assertIn("(cap 330 MBH)", out)


class TestCopperSizing(unittest.TestCase):
    """Gas bench G1 loads on 402.4(10) at the 100 ft row, per the PDF:
    3/4 K&L 116, 1 K&L 247, 1-1/4 K&L 445 CFH. Heat content 840."""

    def test_g1_style_answers(self):
        g = _two_fixture_graph(100.0, 80.0, 99.45)
        r = sizing_engine.size_system(
            g, "Semirigid Copper Tubing", 0.7, "402.4(10)",
            heat_content_btu_per_cf=840.0)
        self.assertEqual(r["table_length_used_ft"], 100)
        self.assertEqual(r["sizes"][10], "1 K&L")        # 180 MBH, 214 CFH
        self.assertEqual(r["sizes"][11], "1 K&L")        # 100 MBH, 119 CFH
        self.assertEqual(r["sizes"][12], "3/4 K&L")      # 80 MBH, 95 CFH
        d = [x for x in r["segment_detail"] if x["pipe_id"] == 10][0]
        self.assertEqual(d["units"], "CFH")
        self.assertAlmostEqual(d["demand_cfh"], 214.3, places=1)

    def test_two_inch_column_restored(self):
        self.assertEqual(gas_tables.get_capacity("402.4(9)", 10, "2 K&L"),
                         3490)


if __name__ == "__main__":
    unittest.main(verbosity=1)
