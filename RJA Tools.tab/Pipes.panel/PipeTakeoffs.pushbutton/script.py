# encoding: utf-8
"""Pipe Takeoffs - Toggle tool for automated domestic water stub-outs.

Automates branch pipe takeoff placement from domestic water mains.
Click 1: Select existing CW/HW/HWC main pipe
Click 2: Pick destination point (fixture/wall location)
Script builds: tee, 6in rise, elbow, horizontal run, elbow, drop to AFF,
               elbow turning toward main, 6in stub-in.

Assembly: 4 pipe segments + 1 tee + 3 elbows
All branch pipe properties copied from the clicked main pipe.

On activation a fixture picker dialog appears showing all fixtures with
pipe size and AFF. A Custom option allows manual entry. ESC or re-click
the button to deactivate.
"""

# ============================================================================
# IMPORTS
# ============================================================================
import clr
import math
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
    Transaction
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
    Grid, StackPanel, ScrollViewer, Border, RadioButton,
    Label, TextBox, ComboBox, ComboBoxItem, Button,
    ColumnDefinition, RowDefinition, Separator, GroupBox
)
from System.Windows.Media import SolidColorBrush, Color
from System.Windows import FontWeights, FontStyles

from pyrevit import revit, DB, UI, script, forms

# ============================================================================
# CONSTANTS
# ============================================================================
ENVVAR_ACTIVE       = "PIPE_TAKEOFFS_ACTIVE"
ENVVAR_FIXTURE      = "PIPE_TAKEOFFS_FIXTURE"
ENVVAR_CUSTOM_SIZE  = "PIPE_TAKEOFFS_CUSTOM_SIZE_RAW"
ENVVAR_LEVEL        = "PIPE_TAKEOFFS_LEVEL_ID"
ENVVAR_CUSTOM_AFF   = "PIPE_TAKEOFFS_CUSTOM_AFF_RAW"

RISE_HEIGHT       = 0.5
STUB_LENGTH       = 0.5
DIAGONAL_WARN_DEG = 5.0

DEFAULT_FIXTURE    = "Lavatory"
DEFAULT_CUSTOM_SIZE = '1/2"'
DEFAULT_CUSTOM_AFF  = '36"'

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

# Valid nominal pipe sizes in inches for custom input parsing
VALID_PIPE_SIZES = {
    0.25:  '1/4"',
    0.375: '3/8"',
    0.5:   '1/2"',
    0.75:  '3/4"',
    1.0:   '1"',
    1.25:  '1-1/4"',
    1.5:   '1-1/2"',
    2.0:   '2"',
    2.5:   '2-1/2"',
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
    """Return list of (name, elevation_ft) tuples sorted by elevation."""
    from Autodesk.Revit.DB import FilteredElementCollector, Level
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    result = []
    for lvl in levels:
        try:
            result.append((lvl.Name, lvl.Elevation, lvl.Id.IntegerValue))
        except Exception:
            pass
    result.sort(key=lambda x: x[1])
    return result


# ============================================================================
# FIXTURE PICKER WPF DIALOG
# ============================================================================
class FixturePickerDialog(Window):
    """WPF dialog showing fixture table with pipe size and AFF columns.
    Custom row at bottom allows manual size and AFF selection.
    """

    # Dark theme colors matching Revit
    CLR_BG        = Color.FromRgb(45,  45,  45)
    CLR_PANEL     = Color.FromRgb(55,  55,  55)
    CLR_ROW_ALT   = Color.FromRgb(50,  50,  50)
    CLR_HIGHLIGHT = Color.FromRgb(0,   120, 215)
    CLR_BORDER    = Color.FromRgb(80,  80,  80)
    CLR_TEXT      = Color.FromRgb(220, 220, 220)
    CLR_TEXT_DIM  = Color.FromRgb(160, 160, 160)
    CLR_BTN       = Color.FromRgb(0,   100, 180)
    CLR_BTN_TEXT  = Color.FromRgb(255, 255, 255)

    def __init__(self, saved_fixture, saved_custom_size, saved_custom_aff, levels, saved_level_idx):
        self.result_fixture     = None
        self.result_dia_ft      = None
        self.result_aff_ft      = None
        self.result_level_elev  = 0.0
        self._saved_fixture     = saved_fixture
        self._saved_custom_size = saved_custom_size
        self._saved_custom_aff  = saved_custom_aff
        self._levels            = levels  # list of (name, elev_ft, id_int)
        self._saved_level_idx   = saved_level_idx
        self._radio_buttons      = {}
        self._custom_size_input  = None
        self._custom_aff_input   = None
        self._level_combo        = None
        self._build_ui()

    def _brush(self, color):
        return SolidColorBrush(color)

    def _build_ui(self):
        self.Title          = "Pipe Takeoffs - Select Fixture"
        self.Width          = 480
        self.SizeToContent  = System.Windows.SizeToContent.Height
        self.ResizeMode     = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background     = self._brush(self.CLR_BG)

        outer = StackPanel()
        outer.Margin = Thickness(16, 16, 16, 16)

        # Title label
        title_lbl = Label()
        title_lbl.Content    = "Select Fixture Type"
        title_lbl.Foreground = self._brush(self.CLR_TEXT)
        title_lbl.FontSize   = 14
        title_lbl.FontWeight = FontWeights.Bold
        title_lbl.Margin     = Thickness(0, 0, 0, 10)
        outer.Children.Add(title_lbl)

        # Level selector
        level_panel = StackPanel()
        level_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        level_panel.Margin = Thickness(0, 0, 0, 10)

        level_lbl = Label()
        level_lbl.Content    = "Finished Floor Level:"
        level_lbl.Foreground = self._brush(self.CLR_TEXT)
        level_lbl.FontSize   = 11
        level_lbl.Padding    = Thickness(0, 3, 8, 0)
        level_lbl.VerticalContentAlignment = VerticalAlignment.Center
        level_panel.Children.Add(level_lbl)

        level_cb = ComboBox()
        level_cb.FontSize = 11
        level_cb.MinWidth = 220
        for (name, elev_ft, id_int) in self._levels:
            item = ComboBoxItem()
            elev_in = elev_ft * 12.0
            if elev_in >= 0:
                item.Content = "{}  (+{:.0f}\")".format(name, elev_in)
            else:
                item.Content = "{}  ({:.0f}\")".format(name, elev_in)
            level_cb.Items.Add(item)
        level_cb.SelectedIndex = max(0, min(self._saved_level_idx, len(self._levels) - 1))
        level_panel.Children.Add(level_cb)
        outer.Children.Add(level_panel)
        self._level_combo = level_cb

        # Table header
        header = self._make_row(
            "Fixture", "Pipe Size", "AFF",
            is_header=True
        )
        outer.Children.Add(header)

        # Fixture rows
        for i, (name, (dia_in, aff_in)) in enumerate(FIXTURES.items()):
            size_label = self._dia_label(dia_in)
            aff_label  = '{}"'.format(int(aff_in))
            row        = self._make_fixture_row(name, size_label, aff_label, i)
            outer.Children.Add(row)

        # Divider
        sep = Separator()
        sep.Margin     = Thickness(0, 10, 0, 6)
        sep.Background = self._brush(self.CLR_BORDER)
        outer.Children.Add(sep)

        # Custom row
        custom_section = self._make_custom_row(len(FIXTURES))
        outer.Children.Add(custom_section)

        # Divider
        sep2 = Separator()
        sep2.Margin     = Thickness(0, 10, 0, 10)
        sep2.Background = self._brush(self.CLR_BORDER)
        outer.Children.Add(sep2)

        # Start button
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

        # Set initial radio selection
        self._select_initial(self._saved_fixture)

    def _dia_label(self, dia_in):
        """Convert decimal inches to fraction string."""
        mapping = {
            0.25:  '1/4"',
            0.375: '3/8"',
            0.5:   '1/2"',
            0.75:  '3/4"',
            1.0:   '1"',
            1.25:  '1-1/4"',
            1.5:   '1-1/2"',
            2.0:   '2"',
            2.5:   '2-1/2"',
            3.0:   '3"',
        }
        return mapping.get(dia_in, '{}"'.format(dia_in))

    def _make_row(self, col1, col2, col3, is_header=False):
        """Build a 3-column grid row (header version)."""
        grid = Grid()
        grid.Margin = Thickness(0, 1, 0, 1)

        for w in [220, 110, 110]:
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            grid.ColumnDefinitions.Add(cd)

        bg = self.CLR_PANEL if is_header else self.CLR_BG

        border = Border()
        border.Background   = self._brush(bg)
        border.Padding      = Thickness(6, 4, 6, 4)
        Grid.SetColumnSpan(border, 3)
        grid.Children.Add(border)

        inner = Grid()
        for w in [220, 110, 110]:
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            inner.ColumnDefinitions.Add(cd)
        border.Child = inner

        for col_idx, text in enumerate([col1, col2, col3]):
            lbl = Label()
            lbl.Content    = text
            lbl.Foreground = self._brush(
                self.CLR_TEXT if is_header else self.CLR_TEXT_DIM
            )
            lbl.FontWeight = FontWeights.Bold if is_header else FontWeights.Normal
            lbl.FontSize   = 11
            lbl.Padding    = Thickness(2, 0, 2, 0)
            Grid.SetColumn(lbl, col_idx)
            inner.Children.Add(lbl)

        return grid

    def _make_fixture_row(self, name, size_label, aff_label, index):
        """Build a selectable fixture row with radio button."""
        bg_color = self.CLR_ROW_ALT if index % 2 == 0 else self.CLR_BG

        border = Border()
        border.Background   = self._brush(bg_color)
        border.Padding      = Thickness(6, 3, 6, 3)
        border.Margin       = Thickness(0, 1, 0, 1)

        grid = Grid()
        for w in [220, 110, 110]:
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            grid.ColumnDefinitions.Add(cd)

        # Radio button with fixture name
        rb = RadioButton()
        rb.Content    = name
        rb.GroupName  = "FixtureGroup"
        rb.Foreground = self._brush(self.CLR_TEXT)
        rb.FontSize   = 11
        rb.VerticalContentAlignment = VerticalAlignment.Center
        Grid.SetColumn(rb, 0)
        grid.Children.Add(rb)

        # Size column
        size_lbl = Label()
        size_lbl.Content    = size_label
        size_lbl.Foreground = self._brush(self.CLR_TEXT)
        size_lbl.FontSize   = 11
        size_lbl.Padding    = Thickness(2, 0, 2, 0)
        Grid.SetColumn(size_lbl, 1)
        grid.Children.Add(size_lbl)

        # AFF column
        aff_lbl = Label()
        aff_lbl.Content    = aff_label
        aff_lbl.Foreground = self._brush(self.CLR_TEXT)
        aff_lbl.FontSize   = 11
        aff_lbl.Padding    = Thickness(2, 0, 2, 0)
        Grid.SetColumn(aff_lbl, 2)
        grid.Children.Add(aff_lbl)

        border.Child = grid
        self._radio_buttons[name] = rb
        return border

    def _make_custom_row(self, index):
        """Build the Custom row with free-text inputs for size and AFF."""
        bg_color = self.CLR_ROW_ALT if index % 2 == 0 else self.CLR_BG

        border = Border()
        border.Background = self._brush(bg_color)
        border.Padding    = Thickness(6, 4, 6, 4)
        border.Margin     = Thickness(0, 1, 0, 1)

        grid = Grid()
        for w in [220, 110, 110]:
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            grid.ColumnDefinitions.Add(cd)

        # Radio button
        rb = RadioButton()
        rb.Content    = "Custom"
        rb.GroupName  = "FixtureGroup"
        rb.Foreground = self._brush(self.CLR_TEXT)
        rb.FontSize   = 11
        rb.VerticalContentAlignment = VerticalAlignment.Center
        rb.Checked   += self._on_custom_checked
        rb.Unchecked += self._on_custom_unchecked
        Grid.SetColumn(rb, 0)
        grid.Children.Add(rb)

        # Size text input
        size_tb = TextBox()
        size_tb.FontSize        = 11
        size_tb.IsEnabled       = False
        size_tb.Margin          = Thickness(2, 0, 4, 0)
        size_tb.Padding         = Thickness(3, 1, 3, 1)
        size_tb.Text            = self._saved_custom_size or '1/2'
        size_tb.ToolTip         = 'Enter pipe size in inches e.g. 1/2 or 0.5 or 3/4'
        Grid.SetColumn(size_tb, 1)
        grid.Children.Add(size_tb)

        # AFF text input
        aff_tb = TextBox()
        aff_tb.FontSize   = 11
        aff_tb.IsEnabled  = False
        aff_tb.Margin     = Thickness(2, 0, 0, 0)
        aff_tb.Padding    = Thickness(3, 1, 3, 1)
        aff_tb.Text       = self._saved_custom_aff or '36'
        aff_tb.ToolTip    = 'Enter AFF height in inches e.g. 36 or 48'
        Grid.SetColumn(aff_tb, 2)
        grid.Children.Add(aff_tb)

        border.Child = grid

        self._radio_buttons['Custom'] = rb
        self._custom_size_input = size_tb
        self._custom_aff_input  = aff_tb
        return border

    def _select_initial(self, saved_fixture):
        """Pre-select the radio button matching saved_fixture."""
        if saved_fixture in self._radio_buttons:
            self._radio_buttons[saved_fixture].IsChecked = True
        else:
            # Default to first fixture
            first = list(self._radio_buttons.keys())[0]
            self._radio_buttons[first].IsChecked = True

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

    def _on_start(self, sender, args):
        # Find which radio is checked
        for name, rb in self._radio_buttons.items():
            if rb.IsChecked:
                if name == 'Custom':
                    # Parse free-text inputs tolerantly
                    size_raw = (self._custom_size_input.Text or '').strip()
                    aff_raw  = (self._custom_aff_input.Text  or '').strip()

                    dia_ft = self._parse_size(size_raw)
                    if dia_ft is None:
                        forms.alert(
                            'Could not read pipe size "{}".\n\n'
                            'Enter a decimal or fraction, e.g. 1/2 or 0.75 or 1-1/2'
                            .format(size_raw),
                            title='Invalid Pipe Size'
                        )
                        return

                    aff_ft = self._parse_aff(aff_raw)
                    if aff_ft is None:
                        forms.alert(
                            'Could not read AFF height "{}".\n\n'
                            'Enter a number in inches, e.g. 36 or 48'
                            .format(aff_raw),
                            title='Invalid AFF Height'
                        )
                        return

                    dia_in  = dia_ft * 12.0
                    aff_in  = aff_ft * 12.0
                    size_label = VALID_PIPE_SIZES.get(round(dia_in, 4), '{}"'.format(dia_in))

                    self.result_fixture = 'Custom ({}, {}" AFF)'.format(size_label, int(round(aff_in)))
                    self.result_dia_ft  = dia_ft
                    self.result_aff_ft  = aff_ft

                    # Persist raw text for next session
                    script.set_envvar(ENVVAR_CUSTOM_SIZE, size_raw)
                    script.set_envvar(ENVVAR_CUSTOM_AFF,  aff_raw)
                else:
                    dia_in, aff_in      = FIXTURES[name]
                    self.result_fixture = name
                    self.result_dia_ft  = dia_in / 12.0
                    self.result_aff_ft  = aff_in / 12.0

                # Read level elevation and add to AFF
                lvl_idx = self._level_combo.SelectedIndex if self._level_combo else 0
                if 0 <= lvl_idx < len(self._levels):
                    lvl_elev_ft = self._levels[lvl_idx][1]
                    script.set_envvar(ENVVAR_LEVEL, str(lvl_idx))
                else:
                    lvl_elev_ft = 0.0
                # result_aff_ft is now absolute elevation = level elev + aff offset
                self.result_aff_ft  = lvl_elev_ft + self.result_aff_ft
                self.result_level_elev = lvl_elev_ft

                self.DialogResult = True
                self.Close()
                return

        forms.alert("Please select a fixture.", title="No Selection")

    def _parse_size(self, text):
        """Parse a pipe size string entered by the user. Returns decimal inches or None.

        Accepts: 1/2  0.5  1-1/2  1.5  3/4  2  etc.
        Strips inch symbols and whitespace before parsing.
        """
        t = text.replace('"', '').replace("'", '').strip()
        if not t:
            return None
        # Handle mixed number like 1-1/2
        try:
            if '-' in t:
                parts = t.split('-', 1)
                whole = float(parts[0].strip())
                frac  = self._eval_fraction(parts[1].strip())
                if frac is None:
                    return None
                inches = whole + frac
            elif '/' in t:
                inches = self._eval_fraction(t)
            else:
                inches = float(t)
            if inches is None or inches <= 0 or inches > 12:
                return None
            return inches / 12.0
        except Exception:
            return None

    def _eval_fraction(self, text):
        """Evaluate a simple fraction string like 1/2. Returns float or None."""
        try:
            parts = text.split('/')
            if len(parts) == 2:
                return float(parts[0]) / float(parts[1])
            return float(text)
        except Exception:
            return None

    def _parse_aff(self, text):
        """Parse an AFF height string entered by the user. Returns feet or None.

        Accepts: 36  36"  48  18.5  etc.
        Strips inch symbols and whitespace.
        """
        t = text.replace('"', '').replace("'", '').strip()
        if not t:
            return None
        try:
            inches = float(t)
            if inches <= 0 or inches > 144:
                return None
            return inches / 12.0
        except Exception:
            return None

    def show(self):
        """Show dialog. Returns (fixture_label, dia_ft, aff_ft) or None."""
        result = self.ShowDialog()
        if result:
            return self.result_fixture, self.result_dia_ft, self.result_aff_ft
        return None



def pick_fixture():
    """Show the fixture picker dialog. Returns (name, dia_ft, aff_ft) or None."""
    saved_fixture     = script.get_envvar(ENVVAR_FIXTURE)     or DEFAULT_FIXTURE
    saved_custom_size = script.get_envvar(ENVVAR_CUSTOM_SIZE) or DEFAULT_CUSTOM_SIZE
    saved_custom_aff  = script.get_envvar(ENVVAR_CUSTOM_AFF)  or DEFAULT_CUSTOM_AFF

    if saved_fixture not in FIXTURES and not saved_fixture.startswith('Custom'):
        saved_fixture = DEFAULT_FIXTURE

    levels = get_project_levels()
    if not levels:
        levels = [("Project Base Point", 0.0, -1)]

    saved_level_idx = 0
    try:
        saved_level_idx = int(script.get_envvar(ENVVAR_LEVEL) or 0)
    except Exception:
        saved_level_idx = 0

    dlg = FixturePickerDialog(
        saved_fixture, saved_custom_size, saved_custom_aff,
        levels, saved_level_idx
    )
    return dlg.show()


# ============================================================================
# SELECTION FILTER - CW / HW / HWC pipes only
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

        sys_name   = ""
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

        # Exact match first (catches short abbreviations like "G" safely),
        # then substring match for longer names like "NATURAL GAS"
        if sys_name in VALID_SYSTEMS:
            return True
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

    checks = [
        (RoutingPreferenceRuleGroupType.Junctions, "tee"),
        (RoutingPreferenceRuleGroupType.Elbows,    "elbow"),
        (RoutingPreferenceRuleGroupType.Segments,  "segment"),
    ]
    for group, label in checks:
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
def calculate_takeoff_geometry(props, click1, click2, aff_height):
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
    stub_dir   = XYZ(-perp_dir.X, -perp_dir.Y, 0)
    stub_start = XYZ(drop_end.X, drop_end.Y, drop_end.Z)
    stub_end   = XYZ(
        drop_end.X + stub_dir.X * STUB_LENGTH,
        drop_end.Y + stub_dir.Y * STUB_LENGTH,
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
        'stub_direction': stub_dir,
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
def build_takeoff(main_pipe, click1, click2, branch_dia_ft, aff_height):
    props    = copy_main_properties(main_pipe)
    warnings = check_routing_preferences(
        props['pipe_type_id'], branch_dia_ft, props['pipe_type_name']
    )
    if warnings:
        msg = "Routing preference warnings for '{}':\n\n".format(props['pipe_type_name'])
        msg += "".join("- {}\n".format(w) for w in warnings)
        msg += "\nContinue anyway? Fittings may fail to place."
        if not forms.alert(msg, yes=True, no=True):
            return

    diag_warning = check_diagonal_main(props['direction'])
    if diag_warning:
        if not forms.alert(diag_warning, yes=True, no=True):
            return

    geo     = calculate_takeoff_geometry(props, click1, click2, aff_height)
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
        "Takeoff complete: {} - {:.3f} ft dia, {:.1f}\" AFF"
        .format(props['pipe_type_name'], branch_dia_ft, aff_height * 12.0)
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

    # Show fixture picker dialog
    pick_result = pick_fixture()
    if pick_result is None:
        return  # user cancelled

    fixture_name, branch_dia_ft, aff_height_ft = pick_result

    # Persist fixture selection
    script.set_envvar(ENVVAR_FIXTURE, fixture_name)

    script.set_envvar(ENVVAR_ACTIVE, True)
    script.toggle_icon(True)

    pipe_filter = WaterPipeFilter()

    try:
        while True:
            try:
                # Click 1: pick main pipe
                ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    pipe_filter,
                    "Pick CW/HW/HWC/NG main pipe  |  {}  (ESC to exit)"
                    .format(fixture_name)
                )
                main_pipe = doc.GetElement(ref.ElementId)
                click1    = ref.GlobalPoint

                if main_pipe is None:
                    forms.alert("Could not read selected pipe. Try again.")
                    continue

                # Click 2: pick destination
                click2 = uidoc.Selection.PickPoint(
                    "Pick destination point  |  {}  (ESC to cancel)"
                    .format(fixture_name)
                )

                with revit.Transaction("Pipe Takeoff"):
                    build_takeoff(
                        main_pipe, click1, click2,
                        branch_dia_ft, aff_height_ft
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