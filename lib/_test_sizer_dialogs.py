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
import sizing_engine    # noqa: E402
import ui_helpers       # noqa: E402
import sizer_dialogs    # noqa: E402

NS = ('xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
      'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"')
X_NAME = "{http://schemas.microsoft.com/winfx/2006/xaml}Name"

BODIES = {
    "water": sizer_dialogs._WATER_BODY,
    "gas": sizer_dialogs._GAS_BODY,
    "existing": sizer_dialogs._EXISTING_BODY,
    "performance": sizer_dialogs._PERF_BODY,
    "start": sizer_dialogs._START_BODY,
    "duct": sizer_dialogs._DUCT_BODY,
}

# The start point picker is a choice, not a calculation, so it carries no
# Calculation basis block.
NO_BASIS = set(["start"])

# Every name each dialog function looks up, read from its source so the test
# cannot drift from the code.
FUNCTIONS = {
    "water": ["show_water_dialog", "_init_project", "_init_recirc"],
    "gas": ["show_gas_picker"],
    "existing": ["show_existing_gas_dialog"],
    "performance": ["show_water_performance_dialog", "_init_project",
                    "_init_recirc"],
    "start": ["show_start_point_dialog"],
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


class TestSizeDuctLimitsGrid(unittest.TestCase):
    """The limits grid names are generated per system, so they are checked
    here rather than by the source scan above."""

    def test_every_system_has_both_boxes(self):
        names = _names_in(sizer_dialogs._DUCT_BODY)
        for i in range(len(sizer_dialogs.DUCT_SYSTEMS)):
            self.assertIn("tb_v_%d" % i, names)
            self.assertIn("tb_f_%d" % i, names)

    def test_systems_match_the_engine_defaults(self):
        self.assertEqual(sorted(sizer_dialogs.DUCT_SYSTEMS),
                         sorted(["Supply Air", "Return Air", "Exhaust Air",
                                 "Transfer Air"]))


class TestGasCalculationBasis(unittest.TestCase):

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
        self.assertIn("Longest Run Method", joined)
        self.assertIn("402.4", joined)
        self.assertIn("%g" % shared_params.SPECIFIC_GRAVITY, joined)
        self.assertIn("%g ft" % shared_params.PRV_MIDSTREAM_MIN_DOWNSTREAM_FT,
                      joined)
        self.assertIn("5 ft per elbow", joined)
        self.assertIn("3/4", joined)
        self.assertIn("Heat Content of Gas", joined)


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
