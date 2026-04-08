# encoding: utf-8
"""Pipe Takeoffs - Toggle tool for automated domestic water stub-outs.

Automates branch pipe takeoff placement from domestic water mains.
Click 1: Select existing CW/HW/HWC/NG main pipe
Click 2: Pick destination point (fixture/wall location)
Script builds: tee, 6in rise, elbow, horizontal run, elbow, drop to AFF,
               elbow turning toward or away from main, 6in stub.

Assembly: 4 pipe segments + 1 tee + 3 elbows
All branch pipe properties copied from the clicked main pipe.

On activation a fixture picker dialog appears. ESC or re-click to deactivate.
"""

# ============================================================================
# IMPORTS
# ============================================================================
import clr
import math
import re
from collections import OrderedDict

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from Autodesk.Revit.DB import (
    XYZ,
    ElementId,
    BuiltInParameter,
    BuiltInCategory,
    FilteredElementCollector,
    Level
)
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import (
    OperationCanceledException,
    InvalidOperationException
)

import System
from System.Windows import Window, GridLength, Thickness, HorizontalAlignment, VerticalAlignment
from System.Windows.Controls import (
    Grid, StackPanel, Border, RadioButton,
    Label, TextBox, ComboBox, ComboBoxItem, Button,
    ColumnDefinition, Separator
)
from System.Windows.Controls import Orientation as WpfOrientation
from System.Windows.Media import SolidColorBrush, Color
from System.Windows import FontWeights

from pyrevit import revit, DB, UI, script, forms

# ============================================================================
# CONSTANTS
# ============================================================================
ENVVAR_ACTIVE      = "PIPE_TAKEOFFS_ACTIVE"
ENVVAR_FIXTURE     = "PIPE_TAKEOFFS_FIXTURE"
ENVVAR_CUSTOM_SIZE = "PIPE_TAKEOFFS_CUSTOM_SIZE_RAW"
ENVVAR_CUSTOM_AFF  = "PIPE_TAKEOFFS_CUSTOM_AFF_RAW"
ENVVAR_LEVEL       = "PIPE_TAKEOFFS_LEVEL_ID"
ENVVAR_STUB_DIR    = "PIPE_TAKEOFFS_STUB_DIR"   # "IN" or "OUT"

RISE_HEIGHT       = 0.5    # 6 inches in feet
STUB_LENGTH       = 0.5    # 6 inches in feet
DIAGONAL_WARN_DEG = 5.0

DEFAULT_FIXTURE    = "Lavatory"
DEFAULT_CUSTOM_SIZE = "1/2"
DEFAULT_CUSTOM_AFF  = "36"
DEFAULT_STUB_DIR    = "IN"

# Fixture presets: label -> (nominal diameter inches, AFF inches)
FIXTURES = OrderedDict([
    ('WC - Tank',         (0.5,   19.0)),
    ('WC - Valve',        (1.5,   17.0)),
    ('Lavatory',          (0.5,   34.0)),
    ('Hand Sink',         (0.5,   34.0)),
    ('Shower',            (0.5,   48.0)),
    ('Mop Sink',          (0.75,  24.0)),
    ('Drinking Fountain', (0.5,   36.0)),
    ('Urinal',            (0.75,  24.0)),
])

# Nominal size lookup for display
NOMINAL_SIZE_LABELS = {
    0.25:  '1/4"',  0.375: '3/8"',  0.5:  '1/2"',
    0.75:  '3/4"',  1.0:   '1"',    1.25: '1-1/4"',
    1.5:   '1-1/2"', 2.0:  '2"',    2.5:  '2-1/2"',
    3.0:   '3"',
}

VALID_SYSTEMS = ["CW", "HW", "HWC", "NG", "NATURAL GAS", "G"]

doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()
output = script.get_output()


# ============================================================================
# LEVEL COLLECTOR
# ============================================================================
def get_project_levels():
    """Return list of (name, elevation_ft, id_int) sorted by elevation."""
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    result = []
    for lvl in levels:
        try:
            result.append((lvl.Name, lvl.Elevation, lvl.Id.IntegerValue))
        except Exception:
            pass
    result.sort(key=lambda x: x[1])
    return result or [("Project Base Point", 0.0, -1)]


# ============================================================================
# ROBUST INPUT PARSERS
# ============================================================================
def _clean(text):
    """Strip symbols, normalize whitespace."""
    return text.strip().replace('\u2019', "'").replace('\u201c', '"').replace('\u201d', '"')


def parse_pipe_size(raw):
    """Parse pipe size input. Returns diameter in FEET or None.

    Accepts virtually any reasonable input:
      Inches (no unit = assumed inches):
        1/2   0.5   3/4   1-1/2   1.5   2
      Inches with symbol:
        1/2"  0.5"  3/4"  1-1/2"
      Feet explicit (only sensible if someone types ft or '):
        0.0417'  (= 1/2 inch)
      Mixed feet-inches (unlikely for pipe size but handled):
        0'1/2"
    Returns None with a reason string on failure.
    """
    t = _clean(raw)
    if not t:
        return None, "Empty input."

    # Strip trailing inch symbol to get bare number/fraction
    has_inch_sym = t.endswith('"')
    has_ft_sym   = t.endswith("'") or t.lower().endswith('ft') or t.lower().endswith('feet')

    t_bare = t.rstrip('"').rstrip("'")
    t_bare = re.sub(r'(?i)(feet|ft)\s*$', '', t_bare).strip()

    # Feet-inches format: 0'6" or 0 ft 6 in
    ft_in = re.match(
        r"^(\d+(?:\.\d+)?)\s*['\u2019ft]+\s*(\d+(?:[\/\-]\d+)?(?:\.\d+)?)\s*\"?$",
        t, re.IGNORECASE
    )
    if ft_in:
        try:
            feet_part   = float(ft_in.group(1))
            inches_part = _eval_fraction(ft_in.group(2))
            if inches_part is None:
                return None, "Could not parse inches portion: {}".format(ft_in.group(2))
            total_inches = feet_part * 12.0 + inches_part
            if total_inches <= 0 or total_inches > 12:
                return None, "Pipe size out of range (0-12 inches)."
            return total_inches / 12.0, None
        except Exception:
            pass

    # Explicit feet only (e.g. 0.0417ft or 0.0417')
    if has_ft_sym and not has_inch_sym:
        try:
            feet = float(t_bare)
            inches = feet * 12.0
            if inches <= 0 or inches > 12:
                return None, "Pipe size out of range."
            return feet, None
        except Exception:
            pass

    # Standard: bare number or fraction, assumed inches
    inches = _eval_mixed(t_bare)
    if inches is None:
        return None, (
            'Could not understand "{}". '
            'Try: 1/2  or  3/4  or  1-1/2  or  0.75'.format(raw)
        )
    if inches <= 0 or inches > 12:
        return None, "Pipe size out of range (0 to 12 inches)."
    return inches / 12.0, None


def parse_aff(raw):
    """Parse AFF height input. Returns height in FEET or None.

    Accepts:
      Inches (no unit = assumed inches):
        36   48   34.5   18
      Inches with symbol:
        36"  48"
      Feet explicit:
        3'   3ft   3feet   3.0'
      Feet-inches:
        3'0"   2'10"   3 ft 4 in   3'-0"
      Decimal feet:
        3.0   (ambiguous - if > 12 treated as inches, else feet)
    """
    t = _clean(raw)
    if not t:
        return None, "Empty input."

    t_upper = t.upper()

    # Feet-inches: 3'4"  or  3'-4"  or  3 ft 4 in  or  3'4
    ft_in = re.match(
        r"^(\d+(?:\.\d+)?)\s*['\u2019][\-\s]*(\d+(?:[\/\-]\d+)?(?:\.\d+)?)\s*\"?(?:\s*in)?$",
        t, re.IGNORECASE
    )
    if ft_in:
        try:
            feet_part   = float(ft_in.group(1))
            inches_part = _eval_fraction(ft_in.group(2))
            if inches_part is None:
                return None, "Could not parse inches in: {}".format(raw)
            total_ft = feet_part + inches_part / 12.0
            if total_ft <= 0 or total_ft > 12:
                return None, "AFF out of range (0-144 inches)."
            return total_ft, None
        except Exception:
            pass

    # Explicit feet only: 3'  3ft  3feet  3.0ft
    ft_only = re.match(
        r"^(\d+(?:\.\d+)?)\s*(?:'|ft|feet)$", t, re.IGNORECASE
    )
    if ft_only:
        try:
            feet = float(ft_only.group(1))
            if feet <= 0 or feet > 12:
                return None, "AFF out of range (0-12 ft)."
            return feet, None
        except Exception:
            pass

    # Explicit inches: 36"  48"
    inch_only = re.match(r'^(\d+(?:\.\d+)?)"$', t)
    if inch_only:
        try:
            inches = float(inch_only.group(1))
            if inches <= 0 or inches > 144:
                return None, "AFF out of range (0-144 inches)."
            return inches / 12.0, None
        except Exception:
            pass

    # Bare number - ambiguous
    try:
        val = float(t)
        # Heuristic: if > 12, almost certainly inches (36, 48...)
        # if <= 12 and looks like a round number of feet, ask
        # We just treat > 12 as inches, <= 12 as feet
        if val <= 0:
            return None, "AFF must be positive."
        if val > 144:
            return None, "AFF out of range (max 144 inches / 12 ft)."
        if val > 12:
            # clearly inches
            return val / 12.0, None
        else:
            # <= 12: could be feet (3.0) or inches (6, 8, 9...)
            # Treat as inches since AFF in feet would be unusual input
            # unless it's a whole number <= 12 - still inches is safer
            return val / 12.0, None
    except Exception:
        pass

    return None, (
        'Could not understand "{}". '
        'Try: 36  or  36"  or  3\'  or  3\'0"'.format(raw)
    )


def _eval_mixed(text):
    """Parse whole+fraction like 1-1/2 or bare fraction 3/4 or decimal. Returns float inches or None."""
    t = text.strip()
    # Mixed: 1-1/2
    m = re.match(r'^(\d+)\s*[\-]\s*(\d+\s*/\s*\d+)$', t)
    if m:
        whole = _eval_fraction(m.group(1))
        frac  = _eval_fraction(m.group(2))
        if whole is not None and frac is not None:
            return whole + frac
    # Fraction: 1/2
    if '/' in t:
        return _eval_fraction(t)
    # Decimal or whole
    try:
        return float(t)
    except Exception:
        return None


def _eval_fraction(text):
    """Evaluate a/b fraction or plain number. Returns float or None."""
    t = text.strip()
    try:
        if '/' in t:
            parts = t.split('/', 1)
            return float(parts[0].strip()) / float(parts[1].strip())
        return float(t)
    except Exception:
        return None


# ============================================================================
# FIXTURE PICKER WPF DIALOG
# ============================================================================
class FixturePickerDialog(Window):

    CLR_BG        = Color.FromRgb(45,  45,  45)
    CLR_PANEL     = Color.FromRgb(55,  55,  55)
    CLR_ROW_ALT   = Color.FromRgb(50,  50,  50)
    CLR_BORDER    = Color.FromRgb(80,  80,  80)
    CLR_TEXT      = Color.FromRgb(220, 220, 220)
    CLR_TEXT_DIM  = Color.FromRgb(160, 160, 160)
    CLR_BTN       = Color.FromRgb(0,   100, 180)
    CLR_BTN_TEXT  = Color.FromRgb(255, 255, 255)
    CLR_ACTIVE    = Color.FromRgb(0,   180, 100)
    CLR_INACTIVE  = Color.FromRgb(160, 60,  60)

    def __init__(self, saved_fixture, saved_custom_size, saved_custom_aff,
                 levels, saved_level_idx, saved_stub_dir):
        self.result_fixture    = None
        self.result_dia_ft     = None
        self.result_aff_ft     = None
        self.result_stub_dir   = saved_stub_dir or DEFAULT_STUB_DIR

        self._saved_fixture     = saved_fixture
        self._saved_custom_size = saved_custom_size
        self._saved_custom_aff  = saved_custom_aff
        self._levels            = levels
        self._saved_level_idx   = saved_level_idx
        self._saved_stub_dir    = saved_stub_dir or DEFAULT_STUB_DIR

        self._radio_buttons     = {}
        self._custom_size_input = None
        self._custom_aff_input  = None
        self._level_combo       = None
        self._stub_btn          = None
        self._stub_dir          = self._saved_stub_dir  # current toggle state

        self._build_ui()

    def _brush(self, color):
        return SolidColorBrush(color)

    def _build_ui(self):
        self.Title             = "Pipe Takeoffs - Select Fixture"
        self.Width             = 500
        self.SizeToContent     = System.Windows.SizeToContent.Height
        self.ResizeMode        = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background        = self._brush(self.CLR_BG)

        outer = StackPanel()
        outer.Margin = Thickness(16, 16, 16, 16)

        # Title
        title = Label()
        title.Content    = "Select Fixture Type"
        title.Foreground = self._brush(self.CLR_TEXT)
        title.FontSize   = 14
        title.FontWeight = FontWeights.Bold
        title.Margin     = Thickness(0, 0, 0, 10)
        outer.Children.Add(title)

        # ---- Level row ----
        level_row = StackPanel()
        level_row.Orientation = WpfOrientation.Horizontal
        level_row.Margin      = Thickness(0, 0, 0, 8)

        lbl = Label()
        lbl.Content    = "Finished Floor Level:"
        lbl.Foreground = self._brush(self.CLR_TEXT)
        lbl.FontSize   = 11
        lbl.Padding    = Thickness(0, 4, 8, 0)
        lbl.VerticalContentAlignment = VerticalAlignment.Center
        level_row.Children.Add(lbl)

        level_cb = ComboBox()
        level_cb.FontSize = 11
        level_cb.MinWidth = 240
        for (name, elev_ft, _) in self._levels:
            elev_in = elev_ft * 12.0
            item = ComboBoxItem()
            sign = "+" if elev_in >= 0 else ""
            item.Content = "{}  ({}{:.0f}\")".format(name, sign, elev_in)
            level_cb.Items.Add(item)
        level_cb.SelectedIndex = max(0, min(self._saved_level_idx, len(self._levels) - 1))
        level_row.Children.Add(level_cb)
        self._level_combo = level_cb
        outer.Children.Add(level_row)

        # ---- Stub direction row ----
        stub_row = StackPanel()
        stub_row.Orientation = WpfOrientation.Horizontal
        stub_row.Margin      = Thickness(0, 0, 0, 12)

        stub_lbl = Label()
        stub_lbl.Content    = "Stub Direction:"
        stub_lbl.Foreground = self._brush(self.CLR_TEXT)
        stub_lbl.FontSize   = 11
        stub_lbl.Padding    = Thickness(0, 4, 8, 0)
        stub_lbl.VerticalContentAlignment = VerticalAlignment.Center
        stub_row.Children.Add(stub_lbl)

        stub_btn = Button()
        stub_btn.Width       = 160
        stub_btn.Height      = 26
        stub_btn.FontSize    = 11
        stub_btn.FontWeight  = FontWeights.Bold
        stub_btn.BorderThickness = Thickness(0)
        stub_btn.Click      += self._on_stub_toggle
        self._stub_btn = stub_btn
        self._update_stub_btn()
        stub_row.Children.Add(stub_btn)

        stub_hint = Label()
        stub_hint.Content    = "  (click to toggle)"
        stub_hint.Foreground = self._brush(self.CLR_TEXT_DIM)
        stub_hint.FontSize   = 10
        stub_hint.VerticalContentAlignment = VerticalAlignment.Center
        stub_row.Children.Add(stub_hint)

        outer.Children.Add(stub_row)

        # ---- Table header ----
        outer.Children.Add(self._make_header())

        # ---- Fixture rows ----
        for i, (name, (dia_in, aff_in)) in enumerate(FIXTURES.items()):
            size_lbl = NOMINAL_SIZE_LABELS.get(dia_in, '{}"'.format(dia_in))
            aff_lbl  = '{}"'.format(int(aff_in))
            outer.Children.Add(self._make_fixture_row(name, size_lbl, aff_lbl, i))

        # ---- Separator ----
        sep = Separator()
        sep.Margin     = Thickness(0, 8, 0, 4)
        sep.Background = self._brush(self.CLR_BORDER)
        outer.Children.Add(sep)

        # ---- Custom row ----
        outer.Children.Add(self._make_custom_row(len(FIXTURES)))

        # ---- Separator ----
        sep2 = Separator()
        sep2.Margin     = Thickness(0, 8, 0, 10)
        sep2.Background = self._brush(self.CLR_BORDER)
        outer.Children.Add(sep2)

        # ---- Start button ----
        btn = Button()
        btn.Content             = "Start"
        btn.Width               = 120
        btn.Height              = 32
        btn.HorizontalAlignment = HorizontalAlignment.Right
        btn.Background          = self._brush(self.CLR_BTN)
        btn.Foreground          = self._brush(self.CLR_BTN_TEXT)
        btn.FontWeight          = FontWeights.Bold
        btn.BorderThickness     = Thickness(0)
        btn.Click              += self._on_start
        outer.Children.Add(btn)

        self.Content = outer
        self._select_initial(self._saved_fixture)

    def _update_stub_btn(self):
        if self._stub_dir == "IN":
            self._stub_btn.Content    = "Stub IN  (toward main)"
            self._stub_btn.Background = self._brush(self.CLR_ACTIVE)
            self._stub_btn.Foreground = self._brush(self.CLR_BTN_TEXT)
        else:
            self._stub_btn.Content    = "Stub OUT  (away from main)"
            self._stub_btn.Background = self._brush(self.CLR_INACTIVE)
            self._stub_btn.Foreground = self._brush(self.CLR_BTN_TEXT)

    def _on_stub_toggle(self, sender, args):
        self._stub_dir = "OUT" if self._stub_dir == "IN" else "IN"
        self._update_stub_btn()

    def _make_header(self):
        grid, inner = self._row_grid(self.CLR_PANEL)
        for col_idx, text in enumerate(["Fixture", "Pipe Size", "AFF"]):
            lbl = Label()
            lbl.Content    = text
            lbl.Foreground = self._brush(self.CLR_TEXT)
            lbl.FontWeight = FontWeights.Bold
            lbl.FontSize   = 11
            lbl.Padding    = Thickness(2, 0, 2, 0)
            Grid.SetColumn(lbl, col_idx)
            inner.Children.Add(lbl)
        return grid

    def _row_grid(self, bg_color):
        border = Border()
        border.Background = self._brush(bg_color)
        border.Padding    = Thickness(6, 3, 6, 3)
        border.Margin     = Thickness(0, 1, 0, 1)

        outer_grid = Grid()
        for w in [230, 110, 110]:
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            outer_grid.ColumnDefinitions.Add(cd)

        inner = Grid()
        for w in [230, 110, 110]:
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            inner.ColumnDefinitions.Add(cd)

        border.Child   = inner
        Grid.SetColumnSpan(border, 3)
        outer_grid.Children.Add(border)
        return outer_grid, inner

    def _make_fixture_row(self, name, size_label, aff_label, index):
        bg = self.CLR_ROW_ALT if index % 2 == 0 else self.CLR_BG
        _, inner = self._row_grid(bg)
        # get the border (first child of outer)
        border = None
        for ch in _.Children:
            border = ch
            break
        # rebuild properly - just use a flat border
        b = Border()
        b.Background = self._brush(bg)
        b.Padding    = Thickness(6, 3, 6, 3)
        b.Margin     = Thickness(0, 1, 0, 1)

        g = Grid()
        for w in [230, 110, 110]:
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            g.ColumnDefinitions.Add(cd)

        rb = RadioButton()
        rb.Content    = name
        rb.GroupName  = "FixtureGroup"
        rb.Foreground = self._brush(self.CLR_TEXT)
        rb.FontSize   = 11
        rb.VerticalContentAlignment = VerticalAlignment.Center
        Grid.SetColumn(rb, 0)
        g.Children.Add(rb)

        for col, text in [(1, size_label), (2, aff_label)]:
            lbl = Label()
            lbl.Content    = text
            lbl.Foreground = self._brush(self.CLR_TEXT)
            lbl.FontSize   = 11
            lbl.Padding    = Thickness(2, 0, 2, 0)
            Grid.SetColumn(lbl, col)
            g.Children.Add(lbl)

        b.Child = g
        self._radio_buttons[name] = rb
        return b

    def _make_custom_row(self, index):
        bg = self.CLR_ROW_ALT if index % 2 == 0 else self.CLR_BG

        b = Border()
        b.Background = self._brush(bg)
        b.Padding    = Thickness(6, 4, 6, 4)
        b.Margin     = Thickness(0, 1, 0, 1)

        g = Grid()
        for w in [230, 110, 110]:
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            g.ColumnDefinitions.Add(cd)

        rb = RadioButton()
        rb.Content    = "Custom"
        rb.GroupName  = "FixtureGroup"
        rb.Foreground = self._brush(self.CLR_TEXT)
        rb.FontSize   = 11
        rb.VerticalContentAlignment = VerticalAlignment.Center
        rb.Checked   += self._on_custom_checked
        rb.Unchecked += self._on_custom_unchecked
        Grid.SetColumn(rb, 0)
        g.Children.Add(rb)

        size_tb = TextBox()
        size_tb.FontSize  = 11
        size_tb.IsEnabled = False
        size_tb.Margin    = Thickness(2, 0, 4, 0)
        size_tb.Padding   = Thickness(3, 1, 3, 1)
        size_tb.Text      = self._saved_custom_size or DEFAULT_CUSTOM_SIZE
        size_tb.ToolTip   = (
            "Pipe size in inches.\n"
            "Examples: 1/2  or  3/4  or  1-1/2  or  0.75  or  1/2\""
        )
        Grid.SetColumn(size_tb, 1)
        g.Children.Add(size_tb)

        aff_tb = TextBox()
        aff_tb.FontSize  = 11
        aff_tb.IsEnabled = False
        aff_tb.Margin    = Thickness(2, 0, 0, 0)
        aff_tb.Padding   = Thickness(3, 1, 3, 1)
        aff_tb.Text      = self._saved_custom_aff or DEFAULT_CUSTOM_AFF
        aff_tb.ToolTip   = (
            "AFF height. Accepts inches or feet.\n"
            "Examples: 36  or  36\"  or  3'  or  3'0\"  or  3ft"
        )
        Grid.SetColumn(aff_tb, 2)
        g.Children.Add(aff_tb)

        b.Child = g
        self._radio_buttons['Custom'] = rb
        self._custom_size_input = size_tb
        self._custom_aff_input  = aff_tb
        return b

    def _on_custom_checked(self, sender, args):
        if self._custom_size_input:
            self._custom_size_input.IsEnabled = True
            self._custom_size_input.Focus()
        if self._custom_aff_input:
            self._custom_aff_input.IsEnabled = True

    def _on_custom_unchecked(self, sender, args):
        if self._custom_size_input:
            self._custom_size_input.IsEnabled = False
        if self._custom_aff_input:
            self._custom_aff_input.IsEnabled = False

    def _select_initial(self, saved_fixture):
        if saved_fixture in self._radio_buttons:
            self._radio_buttons[saved_fixture].IsChecked = True
        else:
            first = list(self._radio_buttons.keys())[0]
            self._radio_buttons[first].IsChecked = True

    def _on_start(self, sender, args):
        for name, rb in self._radio_buttons.items():
            if not rb.IsChecked:
                continue

            if name == 'Custom':
                size_raw = (self._custom_size_input.Text or '').strip()
                aff_raw  = (self._custom_aff_input.Text  or '').strip()

                dia_ft, size_err = parse_pipe_size(size_raw)
                if dia_ft is None:
                    forms.alert(
                        "Invalid pipe size: {}\n\n{}\n\n"
                        "Examples: 1/2  or  3/4  or  1-1/2  or  0.75  or  1/2\""
                        .format(size_raw, size_err),
                        title="Invalid Pipe Size"
                    )
                    return

                aff_offset_ft, aff_err = parse_aff(aff_raw)
                if aff_offset_ft is None:
                    forms.alert(
                        "Invalid AFF height: {}\n\n{}\n\n"
                        "Examples: 36  or  36\"  or  3'  or  3'0\"  or  3ft"
                        .format(aff_raw, aff_err),
                        title="Invalid AFF Height"
                    )
                    return

                dia_in     = dia_ft * 12.0
                aff_in     = aff_offset_ft * 12.0
                size_label = NOMINAL_SIZE_LABELS.get(round(dia_in, 4),
                                                     '{:.3g}"'.format(dia_in))
                self.result_fixture  = 'Custom ({}, {:.0f}" AFF)'.format(size_label, aff_in)
                self.result_dia_ft   = dia_ft
                self.result_aff_ft   = aff_offset_ft

                script.set_envvar(ENVVAR_CUSTOM_SIZE, size_raw)
                script.set_envvar(ENVVAR_CUSTOM_AFF,  aff_raw)

            else:
                dia_in, aff_in      = FIXTURES[name]
                self.result_fixture = name
                self.result_dia_ft  = dia_in / 12.0
                self.result_aff_ft  = aff_in / 12.0

            # Add level elevation to AFF offset
            lvl_idx = self._level_combo.SelectedIndex if self._level_combo else 0
            if 0 <= lvl_idx < len(self._levels):
                lvl_elev_ft = self._levels[lvl_idx][1]
                script.set_envvar(ENVVAR_LEVEL, str(lvl_idx))
            else:
                lvl_elev_ft = 0.0
            self.result_aff_ft += lvl_elev_ft

            # Stub direction
            self.result_stub_dir = self._stub_dir
            script.set_envvar(ENVVAR_STUB_DIR, self._stub_dir)

            self.DialogResult = True
            self.Close()
            return

        forms.alert("Please select a fixture.", title="No Selection")

    def show(self):
        result = self.ShowDialog()
        if result:
            return (
                self.result_fixture,
                self.result_dia_ft,
                self.result_aff_ft,
                self.result_stub_dir
            )
        return None


# ============================================================================
# PICK FIXTURE ENTRY POINT
# ============================================================================
def pick_fixture():
    saved_fixture     = script.get_envvar(ENVVAR_FIXTURE)     or DEFAULT_FIXTURE
    saved_custom_size = script.get_envvar(ENVVAR_CUSTOM_SIZE) or DEFAULT_CUSTOM_SIZE
    saved_custom_aff  = script.get_envvar(ENVVAR_CUSTOM_AFF)  or DEFAULT_CUSTOM_AFF
    saved_stub_dir    = script.get_envvar(ENVVAR_STUB_DIR)    or DEFAULT_STUB_DIR

    if saved_fixture not in FIXTURES and not saved_fixture.startswith('Custom'):
        saved_fixture = DEFAULT_FIXTURE

    levels = get_project_levels()

    saved_level_idx = 0
    try:
        saved_level_idx = int(script.get_envvar(ENVVAR_LEVEL) or 0)
    except Exception:
        saved_level_idx = 0

    dlg = FixturePickerDialog(
        saved_fixture, saved_custom_size, saved_custom_aff,
        levels, saved_level_idx, saved_stub_dir
    )
    return dlg.show()


# ============================================================================
# SELECTION FILTER - domestic water + gas pipes only
# ============================================================================
class WaterPipeFilter(ISelectionFilter):

    def AllowElement(self, element):
        cat = element.Category
        if cat is None:
            return False
        if cat.Id.IntegerValue != int(BuiltInCategory.OST_PipeCurves):
            return False

        sys_param = element.get_Parameter(
            BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
        )
        if sys_param is None:
            return False

        sys_type_id = sys_param.AsElementId()
        if sys_type_id == ElementId.InvalidElementId:
            return False

        sys_type = doc.GetElement(sys_type_id)
        if sys_type is None:
            return False

        sys_name = ""
        abbr_param = sys_type.get_Parameter(
            BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM
        )
        if abbr_param and abbr_param.AsString():
            sys_name = abbr_param.AsString().strip().upper()
        else:
            name_param = sys_type.get_Parameter(
                BuiltInParameter.ALL_MODEL_TYPE_NAME
            )
            if name_param and name_param.AsString():
                sys_name = name_param.AsString().strip().upper()

        # Exact match for short abbreviations (avoids false positives on "G")
        if sys_name in VALID_SYSTEMS:
            return True
        # Substring match for longer names like "NATURAL GAS"
        return any(v in sys_name for v in VALID_SYSTEMS if len(v) > 1)

    def AllowReference(self, reference, position):
        return False


# ============================================================================
# GEOMETRY HELPERS
# ============================================================================
def project_point_onto_line(point, line_start, line_end):
    line_vec       = line_end - line_start
    point_vec      = point - line_start
    line_length_sq = line_vec.DotProduct(line_vec)
    if line_length_sq < 1e-10:
        return line_start
    t = max(0.0, min(1.0, point_vec.DotProduct(line_vec) / line_length_sq))
    return XYZ(
        line_start.X + t * line_vec.X,
        line_start.Y + t * line_vec.Y,
        line_start.Z + t * line_vec.Z
    )


def get_perpendicular_toward_target(main_dir, tee_point, target_point):
    perp_a    = XYZ(-main_dir.Y,  main_dir.X, 0)
    perp_b    = XYZ( main_dir.Y, -main_dir.X, 0)
    to_target = XYZ(
        target_point.X - tee_point.X,
        target_point.Y - tee_point.Y,
        0
    )
    return perp_a.Normalize() if perp_a.DotProduct(to_target) >= 0 else perp_b.Normalize()


def check_diagonal_main(main_dir):
    abs_x = abs(main_dir.X)
    abs_y = abs(main_dir.Y)
    off_axis_deg = math.degrees(
        math.atan2(abs_y, abs_x) if abs_x >= abs_y else math.atan2(abs_x, abs_y)
    )
    if off_axis_deg > DIAGONAL_WARN_DEG:
        return (
            "Main pipe is {:.1f} deg off-axis.\n"
            "Branch will run perpendicular to the main, "
            "which may not be parallel to walls.\n\nContinue anyway?"
        ).format(off_axis_deg)
    return None


def get_open_connector_closest_to(element, target_point):
    best      = None
    best_dist = float('inf')
    for conn in element.ConnectorManager.Connectors:
        if conn.IsConnected:
            continue
        d = conn.Origin.DistanceTo(target_point)
        if d < best_dist:
            best_dist = d
            best      = conn
    return best


# ============================================================================
# PROPERTY COPY FROM MAIN PIPE
# ============================================================================
def copy_main_properties(pipe):
    location = pipe.Location
    if location is None:
        raise ValueError("Selected pipe has no location curve.")

    curve  = location.Curve
    start  = curve.GetEndPoint(0)
    end_pt = curve.GetEndPoint(1)

    raw_dir  = end_pt - start
    main_dir = XYZ(raw_dir.X, raw_dir.Y, 0).Normalize()

    if abs(end_pt.Z - start.Z) > 0.1:
        raise ValueError(
            "Selected pipe is not horizontal. "
            "Z difference: {:.2f} ft. Select a horizontal main."
            .format(abs(end_pt.Z - start.Z))
        )

    pipe_type_id   = pipe.GetTypeId()
    sys_param      = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
    system_type_id = sys_param.AsElementId() if sys_param else None

    lvl_param = pipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM)
    level_id  = lvl_param.AsElementId() if lvl_param else None
    if level_id is None or level_id == ElementId.InvalidElementId:
        level_id = pipe.ReferenceLevel.Id

    dia_param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
    if dia_param is None:
        raise ValueError("Cannot read pipe diameter.")
    diameter = dia_param.AsDouble()

    pipe_type_elem = doc.GetElement(pipe_type_id)
    pipe_type_name = "Unknown"
    if pipe_type_elem:
        name_param = pipe_type_elem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if name_param and name_param.AsString():
            pipe_type_name = name_param.AsString()

    return {
        'pipe_type_id':   pipe_type_id,
        'system_type_id': system_type_id,
        'level_id':       level_id,
        'start':          start,
        'end':            end_pt,
        'direction':      main_dir,
        'centerline_z':   start.Z,
        'diameter':       diameter,
        'radius':         diameter / 2.0,
        'pipe_type_name': pipe_type_name,
        'element':        pipe,
    }


# ============================================================================
# ROUTING PREFERENCE PRE-CHECK
# ============================================================================
def check_routing_preferences(pipe_type_id, branch_dia_ft, pipe_type_name):
    warnings = []
    pipe_type_elem = doc.GetElement(pipe_type_id)
    if pipe_type_elem is None:
        warnings.append("Cannot read pipe type from main pipe.")
        return warnings

    rpm = None
    try:
        rpm = pipe_type_elem.RoutingPreferenceManager
    except Exception:
        pass

    if rpm is None:
        warnings.append(
            "Pipe type '{}' has no routing preferences. "
            "Fittings may not place correctly.".format(pipe_type_name)
        )
        return warnings

    try:
        from Autodesk.Revit.DB.Plumbing import RoutingPreferenceRuleGroupType
    except ImportError:
        return warnings

    for group, label in [
        (RoutingPreferenceRuleGroupType.Junctions, "tee"),
        (RoutingPreferenceRuleGroupType.Elbows,    "elbow"),
        (RoutingPreferenceRuleGroupType.Segments,  "segment"),
    ]:
        try:
            if rpm.GetNumberOfRules(group) == 0:
                warnings.append(
                    "No {} rule in routing preferences for '{}'.".format(label, pipe_type_name)
                )
        except Exception:
            if label != "segment":
                warnings.append(
                    "Cannot read {} rules for '{}'.".format(label, pipe_type_name)
                )
    return warnings


# ============================================================================
# GEOMETRY CALCULATION
# ============================================================================
def calculate_takeoff_geometry(props, click1, click2, aff_height, stub_dir):
    cl_z        = props['centerline_z']
    main_radius = props['radius']
    main_dir    = props['direction']

    tee_center = project_point_onto_line(
        XYZ(click1.X, click1.Y, cl_z), props['start'], props['end']
    )

    rise_start = XYZ(tee_center.X, tee_center.Y, cl_z)
    rise_end   = XYZ(tee_center.X, tee_center.Y, cl_z + main_radius + RISE_HEIGHT)

    perp_dir       = get_perpendicular_toward_target(main_dir, tee_center, click2)
    to_target_xy   = XYZ(click2.X - tee_center.X, click2.Y - tee_center.Y, 0)
    horiz_distance = abs(to_target_xy.DotProduct(perp_dir))

    horiz_end  = XYZ(
        tee_center.X + perp_dir.X * horiz_distance,
        tee_center.Y + perp_dir.Y * horiz_distance,
        rise_end.Z
    )
    drop_start = XYZ(horiz_end.X, horiz_end.Y, horiz_end.Z)
    drop_end   = XYZ(horiz_end.X, horiz_end.Y, aff_height)

    # Stub direction: IN = back toward main, OUT = away from main
    if stub_dir == "OUT":
        stub_dir_vec = XYZ(perp_dir.X, perp_dir.Y, 0)
    else:
        stub_dir_vec = XYZ(-perp_dir.X, -perp_dir.Y, 0)

    stub_start = XYZ(drop_end.X, drop_end.Y, drop_end.Z)
    stub_end   = XYZ(
        drop_end.X + stub_dir_vec.X * STUB_LENGTH,
        drop_end.Y + stub_dir_vec.Y * STUB_LENGTH,
        drop_end.Z
    )

    if aff_height >= rise_end.Z:
        raise ValueError(
            "Drop terminus ({:.0f}\" AFF) is at or above the horizontal run. "
            "Main pipe is too low or AFF is too high.".format(aff_height * 12.0)
        )
    if horiz_distance < 0.083:
        raise ValueError("Destination point is too close to the main. Pick further away.")
    if (rise_end.Z - rise_start.Z) < 0.01:
        raise ValueError("Rise pipe has zero length. Check main pipe geometry.")
    if (drop_start.Z - drop_end.Z) < 0.01:
        raise ValueError("Drop pipe has zero length. Check AFF height.")

    return {
        'tee_center':     tee_center,
        'rise_start':     rise_start,
        'rise_end':       rise_end,
        'horiz_start':    rise_end,
        'horiz_end':      horiz_end,
        'drop_start':     drop_start,
        'drop_end':       drop_end,
        'stub_start':     stub_start,
        'stub_end':       stub_end,
        'perp_direction': perp_dir,
        'stub_direction': stub_dir_vec,
        'horiz_distance': horiz_distance,
    }


# ============================================================================
# PIPE AND FITTING CREATION
# ============================================================================
def create_pipe_segment(sys_id, type_id, lvl_id, start_pt, end_pt, diameter_ft):
    new_pipe  = Pipe.Create(doc, sys_id, type_id, lvl_id, start_pt, end_pt)
    dia_param = new_pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
    if dia_param:
        dia_param.Set(diameter_ft)
    return new_pipe


def split_main_and_place_tee(main_pipe, tee_center, rise_pipe):
    try:
        from Autodesk.Revit.DB.Plumbing import PlumbingUtils
        new_segment_id = PlumbingUtils.BreakCurve(doc, main_pipe.Id, tee_center)
    except ImportError:
        raise ValueError("PlumbingUtils not available in this Revit version.")

    new_segment = doc.GetElement(new_segment_id)
    if new_segment is None:
        raise ValueError(
            "BreakCurve failed. Click may be too close to a fitting or pipe end."
        )

    doc.Regenerate()

    conn_main_a = get_open_connector_closest_to(main_pipe,   tee_center)
    conn_main_b = get_open_connector_closest_to(new_segment, tee_center)
    conn_branch = get_open_connector_closest_to(rise_pipe,   tee_center)

    if conn_main_a is None:
        raise ValueError("No open connector on main segment A at split point.")
    if conn_main_b is None:
        raise ValueError("No open connector on main segment B at split point.")
    if conn_branch is None:
        raise ValueError("No open connector on rise pipe at tee point.")

    return doc.Create.NewTeeFitting(conn_main_a, conn_main_b, conn_branch)


def place_elbow(pipe_a, point_a, pipe_b, point_b):
    conn_a = get_open_connector_closest_to(pipe_a, point_a)
    conn_b = get_open_connector_closest_to(pipe_b, point_b)
    if conn_a is None:
        raise ValueError("No open connector on pipe near {}".format(point_a))
    if conn_b is None:
        raise ValueError("No open connector on pipe near {}".format(point_b))
    return doc.Create.NewElbowFitting(conn_a, conn_b)


# ============================================================================
# MAIN BUILD FUNCTION
# ============================================================================
def build_takeoff(main_pipe, click1, click2, branch_dia_ft, aff_height, stub_dir):
    props    = copy_main_properties(main_pipe)
    warnings = check_routing_preferences(
        props['pipe_type_id'], branch_dia_ft, props['pipe_type_name']
    )
    if warnings:
        msg  = "Routing preference warnings for '{}':\n\n".format(props['pipe_type_name'])
        msg += "".join("- {}\n".format(w) for w in warnings)
        msg += "\nContinue anyway? Fittings may fail to place."
        if not forms.alert(msg, yes=True, no=True):
            return

    diag_warning = check_diagonal_main(props['direction'])
    if diag_warning:
        if not forms.alert(diag_warning, yes=True, no=True):
            return

    geo     = calculate_takeoff_geometry(props, click1, click2, aff_height, stub_dir)
    sys_id  = props['system_type_id']
    type_id = props['pipe_type_id']
    lvl_id  = props['level_id']

    rise_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id, geo['rise_start'], geo['rise_end'], branch_dia_ft
    )
    doc.Regenerate()
    split_main_and_place_tee(props['element'], geo['tee_center'], rise_pipe)

    horiz_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id, geo['horiz_start'], geo['horiz_end'], branch_dia_ft
    )
    doc.Regenerate()
    place_elbow(rise_pipe, geo['rise_end'], horiz_pipe, geo['horiz_start'])

    drop_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id, geo['drop_start'], geo['drop_end'], branch_dia_ft
    )
    doc.Regenerate()
    place_elbow(horiz_pipe, geo['horiz_end'], drop_pipe, geo['drop_start'])

    stub_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id, geo['stub_start'], geo['stub_end'], branch_dia_ft
    )
    doc.Regenerate()
    place_elbow(drop_pipe, geo['drop_end'], stub_pipe, geo['stub_start'])

    logger.debug(
        "Takeoff complete: {} - {:.3f} ft dia, {:.1f}\" AFF, stub {}"
        .format(props['pipe_type_name'], branch_dia_ft, aff_height * 12.0, stub_dir)
    )


# ============================================================================
# VIEW VALIDATION
# ============================================================================
def validate_view():
    view = doc.ActiveView
    if view.ViewType not in [DB.ViewType.FloorPlan, DB.ViewType.EngineeringPlan]:
        forms.alert(
            "Pipe Takeoffs requires a Floor Plan or Engineering Plan view.\n\n"
            "Current view: {}\n\nSwitch to a plan view and try again."
            .format(view.ViewType),
            title="Wrong View Type"
        )
        return False
    return True


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================
def main():
    is_active = script.get_envvar(ENVVAR_ACTIVE)

    if is_active:
        script.set_envvar(ENVVAR_ACTIVE, False)
        script.toggle_icon(False)
        return

    if not validate_view():
        return

    pick_result = pick_fixture()
    if pick_result is None:
        return

    fixture_name, branch_dia_ft, aff_height_ft, stub_dir = pick_result
    script.set_envvar(ENVVAR_FIXTURE, fixture_name)

    script.set_envvar(ENVVAR_ACTIVE, True)
    script.toggle_icon(True)

    pipe_filter = WaterPipeFilter()

    try:
        while True:
            try:
                ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    pipe_filter,
                    "Pick CW/HW/HWC/NG main pipe  |  {}  Stub {}  (ESC to exit)"
                    .format(fixture_name, stub_dir)
                )
                main_pipe = doc.GetElement(ref.ElementId)
                click1    = ref.GlobalPoint

                if main_pipe is None:
                    forms.alert("Could not read selected pipe. Try again.")
                    continue

                click2 = uidoc.Selection.PickPoint(
                    "Pick destination point  |  {}  Stub {}  (ESC to cancel)"
                    .format(fixture_name, stub_dir)
                )

                with revit.Transaction("Pipe Takeoff"):
                    build_takeoff(
                        main_pipe, click1, click2,
                        branch_dia_ft, aff_height_ft, stub_dir
                    )

            except OperationCanceledException:
                break

            except InvalidOperationException as ex:
                forms.alert(
                    "Operation interrupted:\n{}\n\nTool will deactivate.".format(str(ex)),
                    title="Pipe Takeoff Error"
                )
                break

            except ValueError as ex:
                forms.alert(str(ex), title="Pipe Takeoff - Invalid Geometry")
                continue

            except Exception as ex:
                forms.alert(
                    "Unexpected error:\n{}\n\nTakeoff rolled back. You can try again."
                    .format(str(ex)),
                    title="Pipe Takeoff Error"
                )
                logger.error("Pipe Takeoff error: {}".format(str(ex)))
                continue

    finally:
        script.set_envvar(ENVVAR_ACTIVE, False)
        script.toggle_icon(False)


# ============================================================================
# RUN
# ============================================================================
if __name__ == "__main__" or True:
    main()