# -*- coding: ascii -*-
"""Tests for gas_labels.py (One-Line rebuild step 1, 2026-11-18).

Label rules (Colin): 3"G, 25' with no space before G and no material;
cumulative MBH on line 2; a FUTURE cap shows its planned MBH;
notes inlet pressure and drop from the table JSON, not the 0.7 placeholder.

Run:  python -I _test_gas_labels.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import gas_labels as gl   # noqa: E402


class PipeLabel(unittest.TestCase):

    def test_standard_forms(self):
        self.assertEqual(gl.pipe_label("3", 25.2), '3"G, 25\'')
        self.assertEqual(gl.pipe_label("1-1/4", 6.4), '1-1/4"G, 6\'')
        self.assertEqual(gl.pipe_label("2-1/2", 11.0), '2-1/2"G, 11\'')
        self.assertEqual(gl.pipe_label("1/2", 3.0), '1/2"G, 3\'')

    def test_no_ft_suffix_and_no_space_before_g(self):
        s = gl.pipe_label("3", 25)
        self.assertNotIn("FT", s)
        self.assertNotIn(' G', s)

    def test_material_stripped(self):
        self.assertEqual(gl.pipe_label("1-1/4 K&L", 6.4), '1-1/4"G, 6\'')
        self.assertEqual(gl.pipe_label("2 K&L", 40), '2"G, 40\'')

    def test_inch_mark_in_input(self):
        self.assertEqual(gl.pipe_label('2-1/2"', 10), '2-1/2"G, 10\'')

    def test_csst(self):
        self.assertEqual(gl.pipe_label("EHD-13", 12), "EHD-13G, 12'")

    def test_rounding_half_up(self):
        self.assertEqual(gl.pipe_label("1", 5.5), '1"G, 6\'')
        self.assertEqual(gl.pipe_label("1", 6.5), '1"G, 7\'')
        self.assertEqual(gl.pipe_label("1", 6.49), '1"G, 6\'')

    def test_short_segment_never_zero(self):
        self.assertEqual(gl.pipe_label("1", 0.3), '1"G, 1\'')

    def test_no_size(self):
        self.assertEqual(gl.pipe_label(None, 25.2), "25'")
        self.assertEqual(gl.pipe_label("", 25.2), "25'")


class MbhAndFixtures(unittest.TestCase):

    def test_mbh(self):
        self.assertEqual(gl.mbh_text(1615), "1615 MBH")
        self.assertEqual(gl.mbh_text(1615.0), "1615 MBH")
        self.assertEqual(gl.mbh_text(64.8), "64.8 MBH")
        self.assertEqual(gl.mbh_text(119.98), "120 MBH")
        self.assertEqual(gl.mbh_text(None), "0 MBH")

    def test_segment_label(self):
        self.assertEqual(gl.segment_label("3", 11, 1615),
                         ['3"G, 11\'', "1615 MBH"])

    def test_fixture_label(self):
        self.assertEqual(gl.fixture_label("RTU-1", 64.8),
                         ["RTU-1", "64.8 MBH"])
        self.assertEqual(gl.fixture_label("  ", 90), ["UNNAMED", "90 MBH"])

    def test_cap_labels(self):
        # Colin 2026-11-20: a FUTURE cap carries its planned MBH; at 0 it
        # is a data error the gate refuses, never a FUTURE CONNECTION label
        self.assertEqual(gl.cap_label(None, 0), ["FUTURE", "0 MBH"])
        self.assertEqual(gl.cap_label("UNNAMED", 0), ["FUTURE", "0 MBH"])
        self.assertFalse(hasattr(gl, "FUTURE_CONNECTION"))
        self.assertEqual(gl.cap_label("FUTURE", 120), ["FUTURE", "120 MBH"])
        self.assertEqual(gl.cap_label("", 120), ["FUTURE", "120 MBH"])
        self.assertEqual(gl.cap_label("RTU-9", 0), ["RTU-9", "0 MBH"])

    def test_is_future_name(self):
        for n in (None, "", "unnamed", "Future", "FUTURE RTU"):
            self.assertTrue(gl.is_future_name(n), n)
        self.assertFalse(gl.is_future_name("RTU-1"))


class Notes(unittest.TestCase):

    def test_g3c_notes(self):
        lines = gl.notes_lines("402.4(2)", 1814, 117.0)
        self.assertEqual(lines[1], "GAS PIPING SIZED FOR LESS THAN 2 PSI")
        self.assertEqual(lines[2],
                         "MAX PRESSURE LOSS OF 0.5 IN. W.C. PER IFGC TABLE 402.4(2)")
        self.assertEqual(lines[3], "PIPING MATERIAL: SCHEDULE 40 METALLIC PIPE")
        self.assertEqual(lines[4], "TOTAL CONNECTED LOAD: 1814 MBH")
        self.assertEqual(lines[5], "TOTAL DEVELOPED LENGTH: 117'")
        self.assertEqual(len(lines), len(gl.NOTES_TEMPLATE))

    def test_no_placeholder_pressure(self):
        for tid in ("402.4(1)", "402.4(2)", "402.4(3)", "402.4(8)"):
            text = " ".join(gl.notes_lines(tid, 100, 50))
            self.assertNotIn("0.7", text, tid)
            self.assertNotIn("[X]", text, tid)

    def test_two_psi_copper(self):
        lines = gl.notes_lines("402.4(12)", 630, 200)
        self.assertEqual(lines[1], "GAS PIPING SIZED FOR 2 PSI")
        self.assertEqual(lines[2],
                         "MAX PRESSURE LOSS OF 1 PSI PER IFGC TABLE 402.4(12)")
        self.assertEqual(lines[3], "PIPING MATERIAL: SEMIRIGID COPPER TUBING")

    def test_prv_line_only_with_downstream_table(self):
        base = gl.notes_lines("402.4(5)", 500, 150)
        with_prv = gl.notes_lines("402.4(5)", 500, 150,
                                  downstream_table_id="402.4(2)")
        self.assertEqual(len(with_prv), len(base) + 1)
        self.assertEqual(
            with_prv[-1],
            "DOWNSTREAM OF PRV: LESS THAN 2 PSI, MAX PRESSURE LOSS OF "
            "0.5 IN. W.C. PER IFGC TABLE 402.4(2)")

    def test_prv_line_states_each_side_run(self):
        # n59 (Colin 2026-11-20): the PRV line states the developed length
        # Size Gas used on each side, rounded like the other notes lengths;
        # TOTAL DEVELOPED LENGTH stays the whole system
        zones = [{"label": "Meter to regulator(s)", "run_ft": 84.6},
                 {"label": "After PRV 7", "run_ft": 139.5}]
        lines = gl.notes_lines("402.4(5)", 500, 224.1,
                               downstream_table_id="402.4(2)", zones=zones)
        self.assertEqual(
            lines[-1],
            "DOWNSTREAM OF PRV: LESS THAN 2 PSI, MAX PRESSURE LOSS OF "
            "0.5 IN. W.C. PER IFGC TABLE 402.4(2), DEVELOPED LENGTH "
            "METER SIDE 85', PRV SIDE 140'")
        self.assertIn("TOTAL DEVELOPED LENGTH: 224'", lines)

    def test_zone_runs_text(self):
        self.assertEqual(gl.zone_runs_text(None), "")
        self.assertEqual(gl.zone_runs_text([{"run_ft": 120.0}]), "")
        self.assertEqual(
            gl.zone_runs_text([{"run_ft": 0.0}, {"run_ft": 140.2},
                               {"run_ft": 90.4}]),
            ", DEVELOPED LENGTH METER SIDE 0', PRV SIDE 140', 90'")

    def test_tie_in_notes(self):
        lines = gl.notes_lines("402.4(9)", 270, 120, existing_mbh=400)
        self.assertTrue(lines[0].startswith("CONTRACTOR SHALL FIELD VERIFY"))
        self.assertIn("EXISTING CONNECTED LOAD: 400 MBH", lines)
        self.assertIn("NEW CONNECTED LOAD: 270 MBH", lines)
        self.assertIn("TOTAL CONNECTED LOAD: 670 MBH", lines)
        self.assertIn("TOTAL DEVELOPED LENGTH: 120'", lines)
        self.assertNotIn("METER SERVICE", " ".join(lines))

    def test_template_swappable(self):
        lines = gl.notes_lines("402.4(2)", 10, 5,
                               template=["TABLE {table_id}", "{total_mbh}"])
        self.assertEqual(lines, ["TABLE 402.4(2)", "10"])

    def test_every_table_formats(self):
        import gas_tables
        for opt in gas_tables.get_table_options():
            lines = gl.notes_lines(opt["table_id"], 100, 50)
            self.assertTrue(lines[1].strip() != "GAS PIPING SIZED FOR",
                            opt["table_id"])


class AsciiSource(unittest.TestCase):

    def test_ascii_no_fstrings(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "gas_labels.py")
        f = open(path, "rb")
        try:
            src = f.read()
        finally:
            f.close()
        src.decode("ascii")
        import re
        self.assertIsNone(re.search(r'\bf"|\bf\'', src.decode("ascii")))


if __name__ == "__main__":
    unittest.main()
