# -*- coding: ascii -*-
"""Checks on the themed gas and water sizer dialogs that can be made without
Revit.

WPF cannot run in plain CPython, so this does NOT show a window and proves
nothing about how the dialogs look or behave. It checks the two things that
would stop a window opening at all: that every XAML string is well formed
XML, and that every element the code looks up by name is really in the XAML.
It also checks the gas Calculation basis rows.

Run:  python _test_sizer_dialogs.py
"""

import os
import re
import sys
import types
import unittest
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_prv_zones import _install_revit_stubs   # noqa: E402  (Revit stubs)


def _install_wpf_stubs():
    """Just enough for ui_helpers to import outside Revit."""
    clr = types.ModuleType("clr")
    clr.AddReference = lambda name: None
    sys.modules["clr"] = clr

    for name in ("System", "System.Windows", "System.Windows.Markup",
                 "System.Windows.Media"):
        sys.modules[name] = types.ModuleType(name)
    sys.modules["System.Windows.Markup"].XamlReader = object()
    sys.modules["System.Windows.Media"].Brushes = object()


_install_wpf_stubs()

import shared_params    # noqa: E402
import vav_checks       # noqa: E402
import sizing_engine    # noqa: E402
import ui_helpers       # noqa: E402
import sizer_dialogs    # noqa: E402

NS = ('xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
      'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"')
X_NAME = "{http://schemas.microsoft.com/winfx/2006/xaml}Name"

BODIES = {
    "water": sizer_dialogs._WATER_BODY,
    "gas": sizer_dialogs._GAS_BODY,
    "oneline": sizer_dialogs._GAS_BODY,
    "message": sizer_dialogs._MESSAGE_BODY,
    "performance": sizer_dialogs._PERF_BODY,
    "start": sizer_dialogs._START_BODY,
    "duct": sizer_dialogs._DUCT_BODY,
    "fan": sizer_dialogs._FAN_BODY,
}

# The start point picker is a choice, not a calculation, so it carries no
# Calculation basis block.
NO_BASIS = set(["start", "message"])

# Every name each dialog function looks up, read from its source so the test
# cannot drift from the code.
FUNCTIONS = {
    "water": ["show_water_dialog", "_init_project", "_init_recirc"],
    "gas": ["show_gas_picker", "_gas_picker"],
    "oneline": ["show_oneline_picker", "_gas_picker"],
    "message": ["show_message"],
    "performance": ["show_water_performance_dialog", "_init_project",
                    "_init_recirc"],
    "start": ["show_start_point_dialog"],
    # The per system limits boxes are generated names ('tb_v_%d') the scan
    # cannot see; TestSizeDuctLimitsGrid covers those. This catches the VAV
    # names.
    "duct": ["show_size_duct_dialog"],
    "fan": ["show_exhaust_fan_dialog"],
}


def _names_in(xaml_fragment):
    root = ET.fromstring("<Root %s>%s</Root>" % (NS, xaml_fragment))
    return set(el.get(X_NAME) for el in root.iter() if el.get(X_NAME))


def _names_looked_up(function_name):
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "sizer_dialogs.py")
    source = open(path).read()
    start = source.index("def %s(" % function_name)
    following = source.find("\ndef ", start + 1)
    body = source[start:following if following != -1 else len(source)]
    return set(re.findall(r"_named\(win, '([A-Za-z_0-9]+)'\)", body))


class TestXamlIsWellFormed(unittest.TestCase):

    def test_every_body_parses(self):
        for key, body in BODIES.items():
            try:
                _names_in(body)
            except ET.ParseError as exc:
                self.fail("%s body is not well formed XML: %s" % (key, exc))

    def test_the_whole_window_parses(self):
        header = ui_helpers.rja_header_xaml("Size Water", "Purpose & more \"x\"")
        for key, body in BODIES.items():
            xaml = (sizer_dialogs._SHELL_XAML
                    .replace("@@TITLE@@", sizer_dialogs._q('Title "quoted" & co'))
                    .replace("@@WIDTH@@", "560")
                    .replace("@@RESOURCES@@", "")
                    .replace("@@HEADER@@", header)
                    .replace("@@BODY@@", body))
            try:
                ET.fromstring(xaml.strip())
            except ET.ParseError as exc:
                self.fail("%s window is not well formed XML: %s" % (key, exc))

    def test_no_placeholder_is_left_behind(self):
        for body in BODIES.values():
            self.assertNotIn("@@", body)


class TestEveryLookedUpNameExists(unittest.TestCase):

    def test_names_used_by_each_dialog_are_in_its_xaml(self):
        shell = set(["tb_error", "btn_ok", "btn_cancel", "img_logo"])
        for key, function_names in FUNCTIONS.items():
            wanted = set()
            for function_name in function_names:
                wanted |= _names_looked_up(function_name)
            function_name = ", ".join(function_names)
            self.assertTrue(wanted, "no lookups found in " + function_name)
            present = _names_in(BODIES[key]) | shell
            missing = sorted(wanted - present)
            self.assertEqual(missing, [], "%s looks up names that are not in "
                             "its XAML: %s" % (function_name, missing))

    def test_shared_helpers_names_are_in_the_shell_or_basis(self):
        shell_names = _names_in(
            sizer_dialogs._SHELL_XAML.replace("@@TITLE@@", "t")
            .replace("@@WIDTH@@", "560").replace("@@RESOURCES@@", "")
            .replace("@@HEADER@@", ui_helpers.rja_header_xaml("t", "p"))
            .replace("@@BODY@@", sizer_dialogs._BASIS_XAML)
            .strip().split(">", 1)[1].rsplit("</Window>", 1)[0])
        for name in ("tb_error", "btn_ok", "btn_cancel", "img_logo",
                     "grid_ref"):
            self.assertIn(name, shell_names)

    def test_every_dialog_has_a_calculation_basis_block(self):
        for key, body in BODIES.items():
            if key in NO_BASIS:
                continue
            self.assertIn("grid_ref", _names_in(body), key)
            self.assertIn('Text="Calculation basis"', body, key)


class TestGasScope(unittest.TestCase):
    """Scope at the top of the one gas dialog (Colin, 2026-11-20): New
    Construction (default) or Connect into Existing, with the existing gas
    piping in a Scope options box under Connect into Existing, laid out like
    the Duct Performance Scope. The Existing Conditions popup is gone."""

    def _root(self):
        return ET.fromstring("<Root %s>%s</Root>"
                             % (NS, sizer_dialogs._GAS_BODY))

    def _by_name(self, name):
        found = [el for el in self._root().iter() if el.get(X_NAME) == name]
        self.assertEqual(len(found), 1, name)
        return found[0]

    def _section_titles(self):
        return [el.get("Text") for el in self._root().iter()
                if el.get("Style") == "{DynamicResource RjaSectionTitle}"]

    def test_scope_is_the_first_section(self):
        self.assertEqual(self._section_titles(),
                         ["Scope", "Pipe and table", "Outputs"])
        first = list(self._root())[0]
        self.assertEqual(first.get(X_NAME), "pnl_scope")

    def test_two_radio_choices_new_construction_default(self):
        radios = [el for el in self._root().iter()
                  if el.tag.endswith("RadioButton")]
        self.assertEqual([r.get(X_NAME) for r in radios],
                         ["rb_scope_new", "rb_scope_existing"])
        self.assertEqual([r.get("Content") for r in radios],
                         ["New Construction", "Connect into Existing"])
        self.assertEqual(len(set(r.get("GroupName") for r in radios)), 1)
        self.assertEqual(radios[0].get("IsChecked"), "True")
        self.assertIn(radios[1].get("IsChecked"), (None, "False"))

    def test_scope_options_sit_under_connect_into_existing(self):
        scope = list(self._by_name("pnl_scope"))
        names = [el.get(X_NAME) for el in scope]
        i_existing = names.index("rb_scope_existing")
        i_opts = names.index("pnl_scope_opts")
        # radio, its hint line, then the options box (as in Duct Performance)
        self.assertEqual(i_opts, i_existing + 2)
        opts = self._by_name("pnl_scope_opts")
        self.assertEqual(opts.get("Style"), "{DynamicResource RjaGroupBorder}")
        self.assertEqual(opts.get("Visibility"), "Collapsed")
        inside = set(el.get(X_NAME) for el in opts.iter() if el.get(X_NAME))
        self.assertEqual(inside, set(["pnl_scope_opts", "cb_size",
                                      "tb_length", "tb_load"]))

    def test_existing_fields_start_empty_with_no_size(self):
        self.assertTrue(self._by_name("cb_size").tag.endswith("ComboBox"))
        self.assertIsNone(self._by_name("cb_size").get("SelectedIndex"))
        for n in ("tb_length", "tb_load"):
            self.assertIn(self._by_name(n).get("Text"), (None, ""))
        src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "sizer_dialogs.py")).read()
        body = src[src.index("def _gas_picker("):]
        body = body[:body.index("\ndef ")]
        # never preselected, never recalled, never prefilled
        self.assertIn("cb_size.SelectedIndex = keep", body)
        self.assertIn("keep = -1", body)
        self.assertNotIn("recall_existing", body)
        self.assertNotIn("tb_length.Text =", body)
        self.assertNotIn("tb_load.Text =", body)

    def test_existing_is_validated_inline(self):
        src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "sizer_dialogs.py")).read()
        body = src[src.index("def _gas_picker("):]
        body = body[:body.index("\ndef ")]
        self.assertIn("'Pick the existing pipe size at the '", body)
        self.assertEqual(body.count("kind='nonneg'"), 2)
        self.assertIn('"existing":               existing,', body)

    def test_heat_content_is_in_pipe_and_table(self):
        body = sizer_dialogs._GAS_BODY
        self.assertNotIn('Text="Gas"', body)
        i_table = body.index('x:Name="cb_table"')
        i_heat = body.index('x:Name="tb_heat"')
        i_prv = body.index('x:Name="cb_prv"')
        self.assertTrue(i_table < i_heat < i_prv)
        self.assertIn('Text="BTU/CF"', body)

    def test_existing_pipe_sizes_follow_the_table(self):
        self.assertEqual(sizer_dialogs.existing_pipe_sizes("402.4(2)")[:3],
                         ["1/2", "3/4", "1"])
        self.assertEqual(sizer_dialogs.existing_pipe_sizes(None), [])
        self.assertEqual(sizer_dialogs.existing_pipe_sizes("402.4(99)"), [])

    def test_the_popup_is_gone(self):
        for mod in (sizer_dialogs, ui_helpers):
            self.assertFalse(hasattr(mod, "show_existing_gas_dialog"))
        self.assertFalse(hasattr(sizer_dialogs, "_EXISTING_BODY"))
        self.assertFalse(hasattr(ui_helpers, "_EXISTING_GAS_XAML"))
        self.assertFalse(hasattr(ui_helpers,
                                 "_show_existing_gas_dialog_legacy"))

    def test_basis_rows_explain_both_scopes(self):
        labels = [l for l, _v in sizer_dialogs._EXISTING_BASIS_ROWS]
        # Colin, 2026-11-20: only these two, the rest of the old scope rows
        # (New Construction, Existing pipe, second Developed length) dropped.
        self.assertEqual(labels, ["Longest run"])
        self.assertIn("Longest Run Method", dict(
            sizer_dialogs._EXISTING_BASIS_ROWS)["Longest run"])
        ol = [l for l, _v in sizing_engine.oneline_basis_rows()]
        self.assertEqual(ol, ["Developed length", "Heat content", "Loads",
                              "Mid stream PRV"])
        for l, v in sizer_dialogs._EXISTING_BASIS_ROWS:
            (l + v).encode("ascii")
            self.assertNotIn(" - ", v)
            self.assertNotIn("tick", v.lower())
        gas_labels = [l for l, _v in sizing_engine.basis_of_design_rows()]
        self.assertNotIn("Existing building", gas_labels)

    def test_basis_appended_for_both_tools(self):
        src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "sizer_dialogs.py")).read()
        self.assertIn("basis = basis + _EXISTING_BASIS_ROWS", src)

    def test_fallback_matches(self):
        xaml = ui_helpers._PICKER_XAML
        root = ET.fromstring(xaml)
        names = [el.get("Name") for el in root.iter() if el.get("Name")]
        for n in ("pnlScope", "rbScopeNew", "rbScopeExisting", "pnlScopeOpts",
                  "cbSize", "tbLength", "tbLoad", "tbScopeError",
                  "cbMaterial", "cbTable", "tbHeatContent", "cbPrv",
                  "cbDownstream"):
            self.assertIn(n, names)
        self.assertTrue(names.index("pnlScope") < names.index("cbMaterial")
                        < names.index("tbHeatContent") < names.index("cbPrv"))
        self.assertIn('Name="rbScopeNew" GroupName="gas_scope" '
                      'Content="New Construction" IsChecked="True"',
                      xaml.replace("'", "").replace("\n", ""))
        src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "ui_helpers.py")).read()
        legacy = src[src.index("def _show_picker_legacy("):]
        legacy = legacy[:legacy.index("\ndef ")]
        for n in ("pnlScope", "rbScopeNew", "rbScopeExisting", "pnlScopeOpts",
                  "cbSize", "tbLength", "tbLoad", "tbScopeError"):
            self.assertIn("FindName('%s')" % n, legacy)
        self.assertIn('"existing":               existing,', legacy)

    def test_pick_prompts(self):
        self.assertIn("gas meter", ui_helpers.PICK_GAS)
        self.assertIn("first new gas pipe", ui_helpers.PICK_GAS_EXISTING)
        self.assertNotIn("meter", ui_helpers.PICK_GAS_EXISTING)


class TestGasOutputs(unittest.TestCase):
    """Developed Length Breakdown tick box (Colin, 2026-11-20): One-Line
    only, default off, in an Outputs section above the Calculation basis."""

    HERE = os.path.dirname(os.path.abspath(__file__))

    def _root(self):
        return ET.fromstring("<Root %s>%s</Root>"
                             % (NS, sizer_dialogs._GAS_BODY))

    def _src(self, name):
        f = open(os.path.join(self.HERE, name))
        try:
            return f.read()
        finally:
            f.close()

    def _func(self, name, header):
        src = self._src(name)
        part = src[src.index(header):]
        return part[:part.index("\ndef ")]

    def test_box_is_unticked_in_a_collapsed_outputs_panel(self):
        root = self._root()
        pnl = [e for e in root.iter() if e.get(X_NAME) == "pnl_outputs"]
        self.assertEqual(len(pnl), 1)
        self.assertEqual(pnl[0].get("Visibility"), "Collapsed")
        boxes = [e for e in pnl[0].iter() if e.get(X_NAME) == "cb_breakdown"]
        self.assertEqual(len(boxes), 1)
        self.assertEqual(boxes[0].get("Content"), "Developed Length Breakdown")
        self.assertEqual(boxes[0].get("IsChecked"), "False")

    def test_outputs_sit_before_the_calculation_basis(self):
        names = [e.get(X_NAME) for e in self._root().iter() if e.get(X_NAME)]
        self.assertTrue(names.index("pnl_outputs") < names.index("exp_ref"))
        self.assertTrue(names.index("pnl_prv") < names.index("pnl_outputs"))

    def test_shown_and_returned_for_oneline_only(self):
        picker = self._func("sizer_dialogs.py", "def _gas_picker(")
        self.assertIn("if oneline:", picker)
        self.assertIn("pnl_outputs.Visibility = Visibility.Visible", picker)
        self.assertIn('"developed_length_breakdown"', picker)
        self.assertIn("if oneline else False", picker)

    def test_fallback_has_the_box_for_oneline_only(self):
        root = ET.fromstring(ui_helpers._PICKER_XAML)
        by_name = dict((el.get("Name"), el) for el in root.iter()
                       if el.get("Name"))
        self.assertEqual(by_name["pnlOutputs"].get("Visibility"), "Collapsed")
        self.assertEqual(by_name["cbBreakdown"].get("IsChecked"), "False")
        self.assertEqual(by_name["cbBreakdown"].get("Content"),
                         "Developed Length Breakdown")
        legacy = self._func("ui_helpers.py", "def _show_picker_legacy(")
        self.assertIn("oneline=False", legacy)
        self.assertIn("if oneline:", legacy)
        self.assertIn("pnl_outputs.Visibility = Visibility.Visible", legacy)
        self.assertIn('"developed_length_breakdown"', legacy)
        self.assertIn("if oneline else False", legacy)
        one = self._func("ui_helpers.py", "def show_oneline_picker(")
        self.assertIn("_show_picker_legacy(title, True, oneline=True)", one)

    def test_size_gas_script_does_not_read_the_flag(self):
        base = os.path.join(self.HERE, "..", "Developer Tools.tab")
        flag = "developed_length_breakdown"
        seen_oneline = False
        for root, _d, files in os.walk(base):
            for fn in files:
                if fn != "script.py":
                    continue
                f = open(os.path.join(root, fn), "rb")
                try:
                    text = f.read().decode("latin-1")
                finally:
                    f.close()
                is_oneline = "One-Line.pushbutton" in root
                seen_oneline = seen_oneline or is_oneline
                self.assertEqual(flag in text, is_oneline, root)
        self.assertTrue(seen_oneline)


class TestSizeDuctLimitsGrid(unittest.TestCase):
    """The limits grid names are generated per system, so they are checked
    here rather than by the source scan above."""

    def test_every_system_has_both_boxes(self):
        names = _names_in(sizer_dialogs._DUCT_BODY)
        for i in range(len(sizer_dialogs.DUCT_SYSTEMS)):
            self.assertIn("tb_v_%d" % i, names)
            self.assertIn("tb_f_%d" % i, names)

    def test_systems_match_the_engine_defaults(self):
        # Deliberately unchanged by the VAV plan (2026-10-22): MP is a position
        # in the graph, not a System Type, so it is NOT a DUCT_SYSTEMS row. Its
        # two boxes live in the VAV section (TestSizeDuctVav below).
        self.assertEqual(sorted(sizer_dialogs.DUCT_SYSTEMS),
                         sorted(["Supply Air", "Return Air", "Exhaust Air",
                                 "Transfer Air"]))
        self.assertNotIn(vav_checks.MP_KEY, sizer_dialogs.DUCT_SYSTEMS)


class TestSizeDuctVav(unittest.TestCase):
    """The VAV checkbox and the MP velocity and friction rows."""

    def test_vav_controls_are_in_the_xaml(self):
        names = _names_in(sizer_dialogs._DUCT_BODY)
        for name in ("cb_vav", "pnl_vav"):
            self.assertIn(name, names)
        # MP rows are generated from vav_checks.VAV_FIELDS[:2] (review S5).
        for key, _l, _u, _d, _k in vav_checks.VAV_FIELDS[:2]:
            self.assertIn("tb_vav_" + key, names)
        self.assertEqual([f[0] for f in vav_checks.VAV_FIELDS[:2]],
                         ["mp_fpm", "mp_fric"])

    def test_vav_panel_starts_greyed_out(self):
        root = ET.fromstring("<Root %s>%s</Root>" % (NS, sizer_dialogs._DUCT_BODY))
        pnl = [el for el in root.iter() if el.get(X_NAME) == "pnl_vav"][0]
        self.assertEqual(pnl.get("IsEnabled"), "False")
        cb = [el for el in root.iter() if el.get(X_NAME) == "cb_vav"][0]
        self.assertIn(cb.get("IsChecked"), (None, "False"))

    def test_mp_rows_are_filled_from_vav_checks(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "sizer_dialogs.py")
        source = open(path).read()
        self.assertIn("vav_checks.VAV_FIELDS[:2]", source)
        self.assertIn("limits[vav_checks.MP_KEY]", source)
        self.assertIn("'VAV, inlet'", source)


class TestExhaustFanDialog(unittest.TestCase):
    """Ceiling exhaust fan prompt: airflow only (Colin, 2026-11-03)."""

    def test_airflow_only(self):
        names = _names_in(sizer_dialogs._FAN_BODY)
        self.assertIn("tb_cfm", names)
        self.assertNotIn("tb_sp", names)
        self.assertNotIn("rated static", sizer_dialogs._FAN_BODY.lower())


class TestGasHeatContentOnMbhTables(unittest.TestCase):
    """Propane Tables 402.4(25) to (36) publish MBH: the heat content box
    is greyed out with a note (Colin, 2026-11-04)."""

    def test_note_element_starts_collapsed(self):
        self.assertIn("tb_heat_note", _names_in(sizer_dialogs._GAS_BODY))
        root = ET.fromstring("<Root %s>%s</Root>" % (NS, sizer_dialogs._GAS_BODY))
        note = [el for el in root.iter()
                if el.get("{http://schemas.microsoft.com/winfx/2006/xaml}Name")
                == "tb_heat_note"][0]
        self.assertEqual(note.get("Visibility"), "Collapsed")

    def test_note_text_per_table(self):
        for n in range(25, 37):
            tid = "402.4(%d)" % n
            self.assertEqual(sizer_dialogs.heat_content_note(tid),
                             "Not used: Table %s is published in MBH." % tid)
        for tid in ("402.4(1)", "402.4(10)", "402.4(15)", "402.4(24)"):
            self.assertEqual(sizer_dialogs.heat_content_note(tid), "", tid)
        # n=25: 402.4(37) is propane in CFH, it takes the propane box
        self.assertIn("is propane",
                      sizer_dialogs.heat_content_note("402.4(37)"))
        self.assertEqual(sizer_dialogs.heat_content_note(None), "")
        self.assertEqual(sizer_dialogs.heat_content_note("402.4(99)"), "")


class TestGasCalculationBasis(unittest.TestCase):

    def test_basis_starts_closed(self):
        # Colin, 2026-11-20: Calculation basis defaults to closed.
        self.assertIn('x:Name="exp_ref" IsExpanded="False"',
                      sizer_dialogs._BASIS_XAML)

    def test_rows_are_ascii_label_value_pairs(self):
        rows = sizing_engine.basis_of_design_rows()
        self.assertGreater(len(rows), 0)
        for label, value in rows:
            label.encode("ascii")
            value.encode("ascii")
            self.assertTrue(label.strip())
            self.assertTrue(value.strip())

    def test_rows_carry_the_constants_the_engine_uses(self):
        joined = " ".join(l + " " + v for l, v in
                          sizing_engine.basis_of_design_rows())
        self.assertIn("2024 IFGC", joined)
        self.assertIn("402.4", joined)
        self.assertIn("%g" % shared_params.SPECIFIC_GRAVITY, joined)
        self.assertIn("%g ft x elbows" % shared_params.ELBOW_EQUIV_LENGTH_FT,
                      joined)
        self.assertIn("%g ft x tees" % shared_params.TEE_EQUIV_LENGTH_FT,
                      joined)
        self.assertIn("CFH = (MBH x 1000) / Heat Content", joined)
        self.assertIn("MBH(PRV) =", joined)
        self.assertNotIn("Regulators", joined)
        self.assertIn("3/4", joined)


class TestGasPropaneHeatContent(unittest.TestCase):
    """Table 402.4(37) is undiluted propane published in CFH (2024 IFGC p.176,
    n=24/25): its own Propane heat content box, default 2500 BTU/CF,
    shown when 402.4(37) is the table above or after a mid-stream PRV. The
    natural gas box is greyed whenever no natural gas CFH table is in use.
    Heat content is checked 500 to 3000 BTU/CF, nan and inf rejected
    (n=34, 39), and a box not in use is never validated (n=40)."""

    def _root(self):
        return ET.fromstring("<Root %s>%s</Root>"
                             % (NS, sizer_dialogs._GAS_BODY))

    def _by_name(self, name):
        found = [el for el in self._root().iter() if el.get(X_NAME) == name]
        self.assertEqual(len(found), 1, name)
        return found[0]

    def _src(self, module, function_name):
        src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                module)).read()
        body = src[src.index("def %s(" % function_name):]
        return body[:body.index("\ndef ")]

    def test_propane_box_in_the_themed_body_starts_collapsed(self):
        pnl = self._by_name("pnl_propane")
        self.assertEqual(pnl.get("Visibility"), "Collapsed")
        inside = set(el.get(X_NAME) for el in pnl.iter() if el.get(X_NAME))
        self.assertEqual(inside, set(["pnl_propane", "tb_propane_heat"]))
        self.assertIn('Text="Propane heat content"', sizer_dialogs._GAS_BODY)
        body = sizer_dialogs._GAS_BODY
        self.assertTrue(body.index('x:Name="tb_heat"')
                        < body.index('x:Name="tb_propane_heat"')
                        < body.index('x:Name="cb_prv"'))

    def test_propane_box_in_the_fallback(self):
        root = ET.fromstring(ui_helpers._PICKER_XAML)
        names = [el.get("Name") for el in root.iter() if el.get("Name")]
        for n in ("tbHeatNote", "pnlPropane", "tbPropaneHeat", "tbHeatError"):
            self.assertIn(n, names)
        self.assertTrue(names.index("tbHeatContent")
                        < names.index("tbPropaneHeat") < names.index("cbPrv"))
        legacy = self._src("ui_helpers.py", "_show_picker_legacy")
        for n in ("tbHeatNote", "pnlPropane", "tbPropaneHeat", "tbHeatError"):
            self.assertIn("FindName('%s')" % n, legacy)
        self.assertIn("sizer_dialogs.gas_heat_contents(", legacy)
        self.assertIn('"propane_heat_content_btu_per_cf": propane_heat,',
                      legacy)
        self.assertNotIn("float(tb_heat_content.Text", legacy)

    def test_themed_collect_reads_both_boxes(self):
        body = self._src("sizer_dialogs.py", "_gas_picker")
        self.assertIn("gas_heat_contents(", body)
        self.assertIn('"propane_heat_content_btu_per_cf": propane_heat,', body)
        self.assertNotIn("_number(tb_heat,", body)
        # the PRV tick and the downstream table both refresh the boxes
        self.assertIn("cb_downstream.SelectionChanged += on_downstream_changed",
                      body)

    def test_default_is_the_shared_constant(self):
        # item C: no literal 2500 fallback, shared_params is the one value
        self.assertEqual(sizer_dialogs.default_propane_heat_content(),
                         float(shared_params
                               .DEFAULT_PROPANE_HEAT_CONTENT_BTU_PER_CF))
        src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "sizer_dialogs.py")).read()
        self.assertNotIn("2500.0", src)

    def test_heat_rule_lives_in_gas_tables(self):
        # item C: the range, read_heat_content and the table kind are
        # imported from gas_tables (no WPF), never copied
        import gas_tables
        self.assertIs(sizer_dialogs.read_heat_content,
                      gas_tables.read_heat_content)
        self.assertIs(sizer_dialogs._gas_table_kind, gas_tables.table_kind)
        self.assertEqual((sizer_dialogs.HEAT_CONTENT_MIN_BTU_PER_CF,
                          sizer_dialogs.HEAT_CONTENT_MAX_BTU_PER_CF),
                         (gas_tables.HEAT_CONTENT_MIN_BTU_PER_CF,
                          gas_tables.HEAT_CONTENT_MAX_BTU_PER_CF))
        self.assertEqual(gas_tables.table_kind("402.4(2)"), "natural")
        self.assertEqual(gas_tables.table_kind("402.4(30)"), "mbh")
        self.assertEqual(gas_tables.table_kind("402.4(37)"), "propane")
        self.assertIsNone(gas_tables.table_kind("402.4(99)"))
        self.assertIsNone(gas_tables.table_kind(None))
        src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "sizer_dialogs.py")).read()
        self.assertNotIn("def read_heat_content", src)
        self.assertNotIn("def _gas_table_kind", src)

    def test_which_boxes_per_table(self):
        boxes = sizer_dialogs.gas_heat_boxes
        self.assertEqual(boxes("402.4(2)"), ("", False))
        self.assertEqual(boxes("402.4(25)"),
                         ("Not used: Table 402.4(25) is published in MBH.",
                          False))
        note, propane = boxes("402.4(37)")
        self.assertTrue(propane)
        self.assertIn("402.4(37) is propane", note)
        # n=24: MBH table above, 402.4(37) after the PRV
        note, propane = boxes("402.4(36)", "402.4(37)")
        self.assertTrue(propane)
        self.assertIn("402.4(36) is published in MBH", note)
        self.assertIn("402.4(37) is propane", note)
        self.assertEqual(boxes("402.4(36)", None)[1], False)
        self.assertEqual(boxes(None), ("", False))

    def test_propane_value_is_read_on_402_4_37(self):
        nat, prop, problems = sizer_dialogs.gas_heat_contents(
            "840", "2490", "402.4(36)", "402.4(37)")
        self.assertEqual(problems, [])
        self.assertEqual(prop, 2490.0)
        nat, prop, problems = sizer_dialogs.gas_heat_contents(
            "840", "84", "402.4(37)")
        self.assertEqual([k for k, _m in problems], ["propane"])
        self.assertIn("Propane heat content", problems[0][1])

    def test_range_and_non_finite_rejected(self):
        read = sizer_dialogs.read_heat_content
        self.assertEqual(read("840", "x"), (840.0, None))
        self.assertEqual(read(" 500 ", "x"), (500.0, None))
        self.assertEqual(read("3000", "x"), (3000.0, None))
        for bad in ("8400", "84", "0.84", "499.9", "3000.1", "0", "-840"):
            value, message = read(bad, "Heat content of gas")
            self.assertIsNone(value, bad)
            self.assertIn("between 500 and 3000 BTU/CF", message, bad)
        for bad in ("nan", "inf", "-inf", "1e999", "", "abc"):
            value, message = read(bad, "Heat content of gas")
            self.assertIsNone(value, bad)
            self.assertEqual(message, "Heat content of gas shall be a number.")
        nat, prop, problems = sizer_dialogs.gas_heat_contents(
            "8400", "2500", "402.4(2)")
        self.assertIsNone(nat)
        self.assertEqual([k for k, _m in problems], ["natural"])

    def test_box_not_in_use_is_not_validated(self):
        # n=40: a cleared natural gas box on an MBH table does not block OK
        nat, prop, problems = sizer_dialogs.gas_heat_contents(
            "", "", "402.4(25)")
        self.assertEqual(problems, [])
        self.assertEqual(nat, float(
            shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF))
        self.assertEqual(prop, sizer_dialogs.default_propane_heat_content())
        # the hidden propane box on a natural gas table is not validated
        nat, prop, problems = sizer_dialogs.gas_heat_contents(
            "1000", "garbage", "402.4(2)")
        self.assertEqual((nat, problems), (1000.0, []))


class TestPickElementFailure(unittest.TestCase):
    """n=38: pick_element logs any exception other than the user's cancel and
    tells the user the pick failed."""

    class _OCE(Exception):
        pass

    def setUp(self):
        ui = types.ModuleType("Autodesk.Revit.UI")
        sel = types.ModuleType("Autodesk.Revit.UI.Selection")

        class _ObjectType(object):
            Element = "Element"
        sel.ObjectType = _ObjectType
        exc = types.ModuleType("Autodesk.Revit.Exceptions")
        exc.OperationCanceledException = self._OCE
        self._mods = {}
        for name, mod in (("Autodesk.Revit.UI", ui),
                          ("Autodesk.Revit.UI.Selection", sel),
                          ("Autodesk.Revit.Exceptions", exc)):
            self._mods[name] = sys.modules.get(name)
            sys.modules[name] = mod
        self.logged = []
        logged = self.logged

        class _Logger(object):
            def error(self, text):
                logged.append(text)
        script = types.ModuleType("pyrevit.script")
        script.get_logger = lambda: _Logger()
        self._pyrevit_script = getattr(sys.modules["pyrevit"], "script", None)
        sys.modules["pyrevit"].script = script
        self.alerts = []
        self._alert = ui_helpers.rja_alert
        alerts = self.alerts
        ui_helpers.rja_alert = lambda *a, **k: alerts.append(a)

    def tearDown(self):
        for name, mod in self._mods.items():
            if mod is None:
                sys.modules.pop(name, None)
            else:
                sys.modules[name] = mod
        if self._pyrevit_script is None:
            del sys.modules["pyrevit"].script
        else:
            sys.modules["pyrevit"].script = self._pyrevit_script
        ui_helpers.rja_alert = self._alert

    def _uidoc(self, ex):
        class _Sel(object):
            def PickObject(self, *args):
                raise ex

        class _UiDoc(object):
            Selection = _Sel()
        return _UiDoc()

    def test_real_failure_is_logged_and_reported(self):
        res = ui_helpers.pick_element(
            self._uidoc(RuntimeError("view does not allow picking")), None,
            "Size Gas", "Select", popup=False)
        self.assertIsNone(res)
        self.assertEqual(len(self.logged), 1)
        self.assertIn("view does not allow picking", self.logged[0])
        self.assertEqual(len(self.alerts), 1)
        self.assertEqual(self.alerts[0][2], "error")
        self.assertIn("RuntimeError", ui_helpers.last_pick_error())

    def test_cancel_stays_silent(self):
        res = ui_helpers.pick_element(self._uidoc(self._OCE("Esc")), None,
                                      "Size Gas", "Select", popup=False)
        self.assertIsNone(res)
        self.assertEqual(self.logged, [])
        self.assertEqual(self.alerts, [])
        self.assertIsNone(ui_helpers.last_pick_error())


def _run():
    suite = unittest.TestLoader().loadTestsFromModule(sys.modules[__name__])
    result = unittest.TextTestRunner(verbosity=1).run(suite)
    print("=" * 70)
    print("SUMMARY: {} run, {} passed, {} failed, {} errors".format(
        result.testsRun,
        result.testsRun - len(result.failures) - len(result.errors),
        len(result.failures), len(result.errors)))
    print("=" * 70)
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    sys.exit(_run())
