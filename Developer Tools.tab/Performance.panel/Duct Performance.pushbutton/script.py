# -*- coding: utf-8 -*-
"""
Duct Performance.pushbutton/script.py  --  HVAC Phase 1

Colors ductwork in a copied floor plan view and flags what needs attention.
Mains and branches are judged by two different methods:

  MAIN    (2+ terminals downstream)  velocity + friction against the firm
          design limits, editable in the dialog. Unchanged behaviour,
          including the purple oversized suggestion.

  BRANCH  (exactly 1 terminal downstream)  size match against the published
          diffuser tables in lib/diffuser_tables.py instead. A correctly-
          selected auto-sizing diffuser is picked on static pressure and NC,
          not on duct velocity, so velocity-checking a run that only ever
          feeds that one diffuser is redundant and produces false flags. What
          actually matters on a branch is whether the diffuser is big enough
          for its own CFM and whether the duct is big enough for the diffuser.
          Slot diffusers have no published neck table, so a slot branch is
          judged off the auto-sized slot family: its inlet diameter and its
          Sizing Warning (Colin, 2026-10-22).

The diffuser/duct height clearance check still applies to both, and still
overrides either verdict - it is an installability problem, not a sizing one.

VAV System (dialog checkbox, plan approved 2026-10-22, lib/vav_checks.py):
supply equipment entered mid tree through its inlet is a VAV box. The supply
between the unit and the box inlets is medium pressure (MP): always a main,
judged against the MP limits, never purple. After the box, as above. Flex
upstream of a box and a diffuser on an MP main are red. External static
prices each box and reports the pressure left at it.

Which columns appear in the output tables is chosen in the settings dialog;
the same selection drives both the drafting-view schedule placed on the sheet
and the console table.

Run HVAC Diagnose first to verify the network and CFM values, then run this.

IronPython 2.7 / pyRevit  --  no f-strings, no walrus, no nonlocal.
"""

import os
import sys
import math
import re
import traceback
import datetime

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from pyrevit import script, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector, Transaction,
    BuiltInCategory, BuiltInParameter, ViewSheet, ViewType,
    Viewport, ViewDuplicateOption, ElementId, XYZ, ViewPlan, BrowserOrganization,
    PlanViewPlane, PlanViewRange, Category,
    OverrideGraphicSettings, Color,
    TextNote, TextNoteOptions, TextNoteType,
    Line, ViewDrafting, ViewFamilyType, ViewFamily,
    FamilySymbol, StorageType,
    FilledRegion, FilledRegionType, CurveLoop
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

from System.Windows import (
    Window, WindowStartupLocation, Thickness,
    HorizontalAlignment, VerticalAlignment, SizeToContent, TextWrapping,
    SystemParameters
)
from System.Windows.Controls import (
    Grid, Label, TextBox, Button, StackPanel,
    ColumnDefinition, RowDefinition, Orientation,
    Separator, TextBlock, CheckBox, Expander, ScrollViewer,
    ScrollBarVisibility, ListBox, SelectionMode
)
from System.Windows.Media import SolidColorBrush, Colors
from System.Windows import FontWeights
from System.Windows import GridLength
from System.Windows import Visibility

doc    = __revit__.ActiveUIDocument.Document
uidoc  = __revit__.ActiveUIDocument
output = script.get_output()

# Add lib/ to path
_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

import hvac_graph
import fitting_tables
import ashrae_fittings
import revit_helpers
import rja_drafting
import vav_checks
import console_report
import duct_perf_console
from revit_helpers import eid_int

# ── SMACNA colors ─────────────────────────────────────────────────────────────
GREEN  = Color(0,   200,  0)
YELLOW = Color(255, 215,  0)
RED    = Color(210,  40, 40)
PURPLE = Color(140,  40, 200)
GRAY   = Color(160, 160, 160)

_COLOR_MAP = {'GREEN': GREEN, 'YELLOW': YELLOW, 'RED': RED, 'PURPLE': PURPLE, 'GRAY': GRAY}


# ── output table columns ───────────────────────────────────────────────────────
# ONE definition list drives BOTH output tables: the drafting-view schedule
# placed on the sheet and the console table. They are rendered by different
# code (detail lines + TextNotes vs. an HTML table) and have
# drifted apart before, so the column set, the headers and the cell values all
# come from here and from _row_cells() rather than being written out twice.
#
#   view_w     column width in feet at the schedule view's 1:1 scale
#              (= printed inches / 12)
#   console_w  column width in monospace characters (plain text console
#              fallback, used only when output.print_html fails)
#   default_on starting checkbox state in the settings dialog (selectable
#              columns only - see _FIXED_COLUMN_KEYS)
#
# Some columns are always in the tables and have no checkbox; the rest are
# optional. Defaults for the optional ones are the firm's chosen starting
# state, not a stored preference.
_COLUMN_DEFS = [
    # key          header                              view_w  console_w  default_on
    ('num',       '#',                                  0.050,    3,  True),
    ('status',    'Status',                             0.120,    7,  True),
    ('reason',    'Reason',                             0.245,   34,  True),
    ('role',      'Duct Role',                          0.150,   18,  True),
    ('size',      'Size',                               0.100,    9,  True),
    ('required',  'Required Size',                      0.140,   14,  True),
    ('cfm',       'Total CFM through Duct',             0.200,   22,  False),
    ('fpm',       'Actual FPM',                         0.120,   11,  True),
    ('fric',      'Actual Fric (iwc/100)',              0.180,   21,  True),
    ('length',    'Length (ft)',                        0.090,   11,  False),
]

# Always shown, no checkbox. '#' is also the number printed in the keynote
# circles placed in the view, so it stays in the table to keep that
# correlation.
_FIXED_COLUMN_KEYS = frozenset(
    ['num', 'status', 'reason', 'role', 'size', 'required', 'fpm', 'fric'])


def _optional_columns():
    """The _COLUMN_DEFS entries that get a checkbox, in definition order."""
    return [c for c in _COLUMN_DEFS if c[0] not in _FIXED_COLUMN_KEYS]


def _selected_columns(selected_keys):
    """The _COLUMN_DEFS entries to render: every fixed column plus the
    optional ones the user checked, in definition order.

    Both renderers call this so column ORDER is defined once too, not just
    the set - a dialog that returned an unordered set would otherwise let the
    two tables order themselves differently.
    """
    return [c for c in _COLUMN_DEFS
            if c[0] in _FIXED_COLUMN_KEYS or c[0] in selected_keys]


# ── velocity settings dialog ───────────────────────────────────────────────────
# ── settings dialog, themed XAML version ──────────────────────────────────────
# The dialog the user sees. The code-built version further down
# (_show_velocity_settings_dialog_legacy) is kept ONLY as a fallback: if the XAML
# dialog raises anything while building or loading, show_velocity_settings_dialog
# calls it so the tool can never fail to launch because of styling.
#
# Defaults live here once, for both versions.
_DV_ROWS = [
    ('Supply Air',   800,  0.08),
    ('Return Air',   600,  0.05),
    ('Exhaust Air',  600,  0.05),
    ('Outside Air',  600,  0.05),
    ('Transfer Air', 400,  0.05),
]
_DV_DEFAULT_TOL_PCT = 10      # yellow band: this % above max before red
_DV_DEFAULT_SAFETY_PCT = 10   # SP_LOSS_WORKSHEET's own last row

_SETTINGS_XAML = '''
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Duct Performance Settings"
        Width="680" MinWidth="560" MaxWidth="900"
        SizeToContent="Height" ResizeMode="CanResizeWithGrip"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        Background="White" FontFamily="Segoe UI" FontSize="12"
        Foreground="{DynamicResource RjaTextBrush}"
        UseLayoutRounding="True" SnapsToDevicePixels="True">
    @@RESOURCES@@
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        @@HEADER@@

        <ScrollViewer x:Name="scroll_main" Grid.Row="1"
                      VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
            <StackPanel Margin="16,4,16,12">

                <!-- Scope -->
                <Border Style="{DynamicResource RjaSectionHeader}" Margin="0,10,0,6">
                    <TextBlock Style="{DynamicResource RjaSectionTitle}" FontSize="13" FontWeight="SemiBold" Text="Scope"/>
                </Border>
                <CheckBox x:Name="cb_full_diag" Content="Report all ducts"
                          ToolTip="Unchecked, the report lists only ducts that fail or can be downsized (red, yellow, purple). Checked, every duct is listed and numbered in the plan."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,4"
                           Text="Unchecked lists only the ducts that fail or can be downsized (red, yellow, purple)."/>
                <CheckBox x:Name="cb_vav" Content="VAV System"
                          ToolTip="Supply mains between the unit and the VAV box inlets are checked as medium pressure (MP). Single duct and fan powered boxes alike."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,0"
                           Text="Boxes are found by position: supply equipment entered through its inlet. Ductwork after a box is checked as usual."/>
                <Expander x:Name="exp_vav" IsExpanded="False" Margin="20,4,0,0"
                          ToolTip="VAV limits and checks. Greyed out until VAV System is ticked.">
                    <Expander.Header>
                        <TextBlock FontWeight="SemiBold" Text="VAV"/>
                    </Expander.Header>
                    <Border Style="{DynamicResource RjaGroupBorder}" Margin="0,6,0,0">
                        <Grid x:Name="grid_vav">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                        </Grid>
                    </Border>
                </Expander>

                <!-- Main duct limits -->
                <Border Style="{DynamicResource RjaSectionHeader}" Margin="0,14,0,6">
                    <TextBlock Style="{DynamicResource RjaSectionTitle}" FontSize="13" FontWeight="SemiBold" Text="Main duct limits"/>
                </Border>
                <Grid x:Name="grid_limits">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                </Grid>
                <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                    <TextBlock Text="Yellow tolerance" VerticalAlignment="Center"/>
                    <TextBox x:Name="tb_tol" Width="56" Margin="8,0,6,0"
                             ToolTip="At or under max is green. Within tolerance is yellow. Past it is red. Must be above 0 and below 100."/>
                    <TextBlock Text="% above max before red" VerticalAlignment="Center"/>
                </StackPanel>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="0,4,0,0"
                           Text="At or under max is green, within tolerance is yellow, past it is red."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,0"
                           Text="Limits apply to main ducts. Branches are checked against the diffuser tables."/>

                <!-- Outputs -->
                <Border Style="{DynamicResource RjaSectionHeader}" Margin="0,14,0,6">
                    <TextBlock Style="{DynamicResource RjaSectionTitle}" FontSize="13" FontWeight="SemiBold" Text="Outputs"/>
                </Border>
                <Grid x:Name="grid_cols">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                </Grid>

                <!-- External static pressure, grouped under Outputs -->
                <CheckBox x:Name="cb_static" Content="Calculate total external static pressure" Margin="0,10,0,0"
                          ToolTip="The fan static along the most restrictive run outside the unit, supply path plus return path. See Calculation basis."/>
                <Border x:Name="pnl_static" Style="{DynamicResource RjaGroupBorder}" Margin="20,4,0,0">
                    <StackPanel>
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Safety factor" VerticalAlignment="Center"/>
                            <TextBox x:Name="tb_sf" Width="56" Margin="8,0,6,0"
                                     ToolTip="Percent added to the external static total. Does not affect the per duct checks. Range 0 to 100."/>
                            <TextBlock Text="% added to the external static total" VerticalAlignment="Center"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                            <CheckBox x:Name="cb_filter" Content="Include return filter" VerticalAlignment="Center"
                                      ToolTip="A known filter drop added once to the return path. Leave off when the filter is inside the unit and already deducted from its published ESP."/>
                            <TextBox x:Name="tb_filter" Width="56" Margin="8,0,6,0"
                                     ToolTip="Known filter drop in in. wc. Cannot be negative."/>
                            <TextBlock Text="in. wc known drop (MERV 8), added to the return total" VerticalAlignment="Center"/>
                        </StackPanel>
                        <StackPanel x:Name="pnl_vav_static" Margin="0,10,0,0" Visibility="Collapsed">
                            <TextBlock FontWeight="SemiBold" Text="VAV boxes"/>
                            <Grid x:Name="grid_vav_static" Margin="0,4,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                            </Grid>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <!-- Advanced -->
                <Expander x:Name="exp_adv" IsExpanded="False" Margin="0,16,0,0"
                          ToolTip="Overrides to the calculation basis. Standard values apply unless changed.">
                    <Expander.Header>
                        <TextBlock FontSize="13" FontWeight="SemiBold" Text="Advanced"/>
                    </Expander.Header>
                    <Border Style="{DynamicResource RjaGroupBorder}" Margin="0,6,0,0">
                        <StackPanel>
                            <TextBlock x:Name="tb_adv_note" Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap"
                                       Text="Overrides apply to every duct in the run. Leave at standard values unless the project needs otherwise."/>
                            <Expander x:Name="exp_ctab" IsExpanded="False" Margin="0,8,0,0"
                                      Header="Fitting loss coefficient tables (2001 ASHRAE Fundamentals)"
                                      ToolTip="Each fitting's C is looked up from its table at that fitting's own geometry and flows. Open a fitting to see its table. Read only.">
                                <StackPanel x:Name="pnl_ctab" Margin="16,4,0,4"/>
                            </Expander>
                            <Expander x:Name="exp_c" IsExpanded="False" Margin="0,4,0,0"
                                      Header="Fallback C values (RJA SP loss worksheet)"
                                      ToolTip="Used only where a fitting's geometry or flows cannot be read; the report flags each one as FALLBACK. Cannot be negative.">
                                <Grid x:Name="grid_c" Margin="16,4,0,4">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                </Grid>
                            </Expander>
                            <Expander x:Name="exp_comp" IsExpanded="False" Margin="0,4,0,0"
                                      Header="Component pressure drops (in. wc each)"
                                      ToolTip="Defaults from the component table. Cannot be negative.">
                                <Grid x:Name="grid_comp" Margin="16,4,0,4">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                </Grid>
                            </Expander>
                            <Expander x:Name="exp_basis" IsExpanded="False" Margin="0,4,0,0"
                                      Header="Calculation basis for friction"
                                      ToolTip="Must be greater than 0.">
                                <StackPanel Margin="16,4,0,4">
                                    <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,4"
                                               Text="Applies to every duct in the run, not one system. Air density also drives velocity pressure (fitting losses)."/>
                                    <Grid x:Name="grid_basis">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                    </Grid>
                                </StackPanel>
                            </Expander>
                        </StackPanel>
                    </Border>
                </Expander>

                <!-- Calculation basis reference -->
                <Expander x:Name="exp_ref" IsExpanded="False" Margin="0,8,0,0">
                    <Expander.Header>
                        <TextBlock FontSize="13" FontWeight="SemiBold" Text="Calculation basis"/>
                    </Expander.Header>
                    <Grid x:Name="grid_ref" Margin="16,6,0,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                    </Grid>
                </Expander>

            </StackPanel>
        </ScrollViewer>

        <Border Grid.Row="2" Background="{DynamicResource RjaPanelBrush}"
                BorderBrush="{DynamicResource RjaRuleBrush}" BorderThickness="0,1,0,0"
                Padding="16,10,16,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="tb_error" Foreground="#C62828" TextWrapping="Wrap"
                           VerticalAlignment="Center" Margin="0,0,12,0"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button x:Name="btn_cancel" Content="Cancel" IsCancel="True"
                            Style="{DynamicResource RjaSecondaryButton}" MinWidth="88" Margin="0,0,8,0"/>
                    <Button x:Name="btn_ok" Content="OK" IsDefault="True"
                            Style="{DynamicResource RjaPrimaryButton}" MinWidth="88"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
'''


def _wpf_named(win, name):
    """The element carrying x:Name `name`. Raises if it is not there, so a
    missing name sends the caller to the fallback dialog instead of failing later."""
    el = getattr(win, name, None)
    if el is None:
        el = win.FindName(name)
    if el is None:
        raise RuntimeError('XAML element not found: ' + name)
    return el


def _settings_result(limits, tol_pct, selected_cols, full_diag, ext_static,
                     safety_pct, c_values, comp_values, calc_basis, filter_inwc,
                     vav, vav_settings):
    """The settings dialog's answer, a dict, built in ONE place for both the
    themed and the legacy dialog so main() reads the same keys from either.

    With VAV ticked, the MP limits are also added to limits under
    vav_checks.MP_KEY, so every limit lookup finds them in one dict.
    """
    limits = dict(limits)
    if vav:
        limits[vav_checks.MP_KEY] = (vav_settings['mp_fpm'], vav_settings['mp_fric'])
    return {
        'limits':        limits,
        'tol_pct':       tol_pct,
        'selected_cols': selected_cols,
        'full_diag':     full_diag,
        'ext_static':    ext_static,
        'safety_pct':    safety_pct,
        'c_values':      c_values,
        'comp_values':   comp_values,
        'calc_basis':    calc_basis,
        'filter_inwc':   filter_inwc,
        'vav':           bool(vav),
        'vav_settings':  vav_settings,
    }


def _show_velocity_settings_xaml(use_theme):
    """The XAML settings dialog. Returns the _settings_result dict or None (cancel).

    Raises if the window cannot be built, which show_velocity_settings_dialog
    catches and answers with the next simpler dialog.
    """
    import ui_helpers
    from System.Windows import FontWeights as _FW

    xaml = (_SETTINGS_XAML
            .replace('@@RESOURCES@@', ui_helpers.rja_window_resources_xaml(use_theme))
            .replace('@@HEADER@@', ui_helpers.rja_header_xaml(
                'Duct Performance',
                'Checks duct sizing, velocity and friction against RJA design '
                'standards, and calculates external static pressure.')))
    win = forms.WPFWindow(xaml, literal_string=True)

    N = lambda name: _wpf_named(win, name)
    img_logo = N('img_logo')
    cb_full_diag = N('cb_full_diag')
    cb_vav = N('cb_vav')
    exp_vav = N('exp_vav')
    grid_vav = N('grid_vav')
    pnl_vav_static = N('pnl_vav_static')
    grid_vav_static = N('grid_vav_static')
    grid_limits = N('grid_limits')
    tb_tol = N('tb_tol')
    grid_cols = N('grid_cols')
    cb_static = N('cb_static')
    pnl_static = N('pnl_static')
    tb_sf = N('tb_sf')
    cb_filter = N('cb_filter')
    tb_filter = N('tb_filter')
    exp_adv = N('exp_adv')
    exp_c = N('exp_c')
    exp_comp = N('exp_comp')
    exp_basis = N('exp_basis')
    grid_c = N('grid_c')
    pnl_ctab = N('pnl_ctab')
    grid_comp = N('grid_comp')
    grid_basis = N('grid_basis')
    grid_ref = N('grid_ref')
    tb_error = N('tb_error')
    btn_ok = N('btn_ok')
    btn_cancel = N('btn_cancel')
    scroll_main = N('scroll_main')

    ui_helpers.rja_apply_logo(win, img_logo)
    try:
        win.MaxHeight = SystemParameters.WorkArea.Height * 0.9
    except Exception:
        win.MaxHeight = 800.0

    result = [None]
    all_boxes = []   # every TextBox that can be flagged invalid

    def _box(text, width, margin):
        tb = TextBox()
        tb.Text = text
        tb.Width = width
        tb.Margin = margin
        tb.VerticalAlignment = VerticalAlignment.Center
        tb.HorizontalAlignment = HorizontalAlignment.Left
        ui_helpers.rja_attach_clear_on_edit(tb)
        all_boxes.append(tb)
        return tb

    def _put(grid, el, col, row):
        Grid.SetColumn(el, col)
        Grid.SetRow(el, row)
        grid.Children.Add(el)

    def _add_row_def(grid):
        rd = RowDefinition()
        rd.Height = GridLength.Auto
        grid.RowDefinitions.Add(rd)

    # defaults into the fixed boxes
    tb_tol.Text = str(_DV_DEFAULT_TOL_PCT)
    tb_sf.Text = str(_DV_DEFAULT_SAFETY_PCT)
    tb_filter.Text = '%g' % fitting_tables.DEFAULT_RETURN_FILTER_INWC
    for b in (tb_tol, tb_sf, tb_filter):
        ui_helpers.rja_attach_clear_on_edit(b)
        all_boxes.append(b)
    cb_full_diag.IsChecked = False
    cb_vav.IsChecked = False
    cb_static.IsChecked = False
    cb_filter.IsChecked = False

    # ── main duct limits grid ───────────────────────────────────────────────
    _add_row_def(grid_limits)
    for col, head in ((0, 'System'), (1, 'Max velocity (FPM)'),
                      (2, 'Max friction (in. wc/100 ft)')):
        h = TextBlock()
        h.Text = head
        h.FontWeight = _FW.SemiBold
        h.Margin = Thickness(0, 0, 24, 2)
        _put(grid_limits, h, col, 0)
    vel_boxes = {}
    fric_boxes = {}
    for i, (sys_class, def_fpm, def_fric) in enumerate(_DV_ROWS):
        r = i + 1
        _add_row_def(grid_limits)
        lb = TextBlock()
        lb.Text = sys_class
        lb.VerticalAlignment = VerticalAlignment.Center
        _put(grid_limits, lb, 0, r)
        vb = _box(str(def_fpm), 80, Thickness(0, 2, 24, 2))
        vb.ToolTip = 'Maximum velocity for ' + sys_class + ' main ducts, in FPM. Must be greater than 0.'
        _put(grid_limits, vb, 1, r)
        vel_boxes[i] = vb
        fb = _box(str(def_fric), 80, Thickness(0, 2, 24, 2))
        fb.ToolTip = ('Maximum friction rate for ' + sys_class +
                      ' main ducts, in in. wc per 100 ft. Must be greater than 0.')
        _put(grid_limits, fb, 2, r)
        fric_boxes[i] = fb

    # ── outputs: optional column checkboxes, two columns ────────────────────
    _optional = _optional_columns()
    _ncols = 2
    _nrows = (len(_optional) + _ncols - 1) // _ncols
    for _ in range(max(_nrows, 1)):
        _add_row_def(grid_cols)
    col_boxes = {}
    for i, (key, header, _vw, _cw, default_on) in enumerate(_optional):
        cb = CheckBox()
        cb.Content = header
        cb.IsChecked = default_on
        cb.ToolTip = 'Optional column in the sheet table and the pyRevit window table.'
        _put(grid_cols, cb, i // _nrows, i % _nrows)
        col_boxes[key] = cb

    # ── advanced value rows (same defaults and stores as the legacy dialog) ─
    def _fill_value_grid(grid, rows, store, order, fmt):
        """rows: (key, label, extra, default). order collects (key, label)."""
        for r, (key, label, extra, default) in enumerate(rows):
            _add_row_def(grid)
            lb = TextBlock()
            lb.Text = label
            lb.Margin = Thickness(0, 2, 12, 2)
            lb.VerticalAlignment = VerticalAlignment.Center
            _put(grid, lb, 0, r)
            tb = _box(fmt % default, 64, Thickness(0, 2, 8, 2))
            _put(grid, tb, 1, r)
            store[key] = tb
            order.append((key, label))
            if extra:
                ex = ui_helpers.rja_hint_block(win, extra)
                ex.VerticalAlignment = VerticalAlignment.Center
                _put(grid, ex, 2, r)

    c_boxes = {}
    c_order = []
    _fill_value_grid(
        grid_c,
        [(k, lbl, 'ASHRAE ' + no, v) for k, lbl, no, v in fitting_tables.C_TABLE],
        c_boxes, c_order, '%.2f')
    # ── one collapsed row per fitting, opening to the ASHRAE table it reads
    #    (Colin, 2026-10-23). Filled on first open so the dialog stays quick.
    def _table_view(key):
        panel = StackPanel()
        panel.Margin = Thickness(0, 2, 0, 8)
        t = ashrae_fittings.table(key)
        title = TextBlock()
        title.Text = ashrae_fittings.table_label(key)
        title.FontWeight = _FW.SemiBold
        title.TextWrapping = TextWrapping.Wrap
        panel.Children.Add(title)
        if t is None:
            panel.Children.Add(ui_helpers.rja_hint_block(win, 'Not transcribed.'))
            return panel
        coef = t.get('coef') or 'C'
        if t['kind'] == 'const':
            panel.Children.Add(ui_helpers.rja_hint_block(
                win, '%s = %s' % (coef, t['C'])))
            return panel
        if t['kind'] == '3d':
            panel.Children.Add(ui_helpers.rja_hint_block(win, (
                '%s by %s (layers %s), %s (rows) and %s (columns). Three '
                'dimensional, see the book pages.') % (
                    coef, t['layers']['name'],
                    ', '.join('%g' % v for v in t['layers']['values']),
                    t['rows']['name'], t['cols']['name'])))
            return panel
        if t['kind'] == '1d':
            head = [t['x']['name']] + ['%g' % v for v in t['x']['values']]
            body = [[coef] + ['-' if v is None else '%.2f' % v for v in t['C']]]
        else:
            head = ['%s  x  %s' % (t['rows']['name'], t['cols']['name'])] + [
                '%g' % v for v in t['cols']['values']]
            body = [['%g' % rv] + ['' if v is None else '%.2f' % v for v in row]
                    for rv, row in zip(t['rows']['values'], t['C'])]
        g = Grid()
        g.Margin = Thickness(0, 2, 0, 0)
        for _c in range(len(head)):
            cd = ColumnDefinition()
            cd.Width = GridLength.Auto
            g.ColumnDefinitions.Add(cd)
        for r, row in enumerate([head] + body):
            _add_row_def(g)
            for c, text in enumerate(row):
                tb = TextBlock()
                tb.Text = text
                tb.FontSize = 11
                tb.Margin = Thickness(0, 0, 10, 0)
                if r == 0 or c == 0:
                    tb.FontWeight = _FW.SemiBold
                _put(g, tb, c, r)
        panel.Children.Add(g)
        return panel

    for label, family, keys, _fallback in ashrae_fittings.FITTING_ROWS:
        ex = Expander()
        ex.Header = '%s   (%s)' % (label, ', '.join(k.split('_')[0] for k in keys
                                                     if not k.endswith('_K')))
        ex.IsExpanded = False
        ex.Margin = Thickness(0, 2, 0, 2)
        ex.ToolTip = 'Model family: ' + family
        body = StackPanel()
        body.Margin = Thickness(18, 2, 0, 4)
        ex.Content = body

        def _fill(sender, e, body=body, keys=keys, family=family):
            if body.Children.Count:
                return
            body.Children.Add(ui_helpers.rja_hint_block(win, 'Model family: ' + family))
            seen = set()
            for k in keys:
                if k in seen:
                    continue
                seen.add(k)
                body.Children.Add(_table_view(k))
        ex.Expanded += _fill
        pnl_ctab.Children.Add(ex)

    comp_boxes = {}
    comp_order = []
    _fill_value_grid(
        grid_comp,
        [(k, lbl, 'in. wc', v) for k, lbl, v in fitting_tables.COMPONENT_TABLE],
        comp_boxes, comp_order, '%.3f')
    basis_boxes = {}
    basis_order = []
    _fill_value_grid(
        grid_basis,
        [('rigid_roughness', 'Duct roughness, rigid (ft)',
          '0.0003 galvanized, 0.005 interior insulated',
          hvac_graph.STANDARD_RIGID_ROUGHNESS_FT),
         ('flex_roughness', 'Duct roughness, flex (ft)',
          '0.003 flex, fully extended (ASHRAE)',
          hvac_graph.STANDARD_FLEX_ROUGHNESS_FT),
         ('air_density', 'Air density (lb/ft3)',
          '0.075 standard air, sea level',
          hvac_graph.STANDARD_AIR_DENSITY_LB_FT3)],
        basis_boxes, basis_order, '%g')

    # ── VAV rows: limits in the VAV section, box and static fields under the
    #    external static option (unit max ESP may be blank) ─────────────────
    vav_boxes = {}
    for grid, fields in ((grid_vav, vav_checks.VAV_LIMIT_FIELDS),
                         (grid_vav_static, vav_checks.VAV_STATIC_FIELDS)):
        for r, (key, label, unit, default, _kind) in enumerate(fields):
            _add_row_def(grid)
            lb = TextBlock()
            lb.Text = label
            lb.Margin = Thickness(0, 2, 12, 2)
            lb.VerticalAlignment = VerticalAlignment.Center
            _put(grid, lb, 0, r)
            tb = _box(vav_checks.format_default(default), 64, Thickness(0, 2, 8, 2))
            _put(grid, tb, 1, r)
            vav_boxes[key] = tb
            ex = ui_helpers.rja_hint_block(win, unit)
            ex.VerticalAlignment = VerticalAlignment.Center
            _put(grid, ex, 2, r)

    # ── calculation basis reference (condensed) ─────────────────────────────
    ref_rows = [
        ('Limits', 'Checked on main ducts only. Branches are judged against the '
                   'diffuser capacity tables.'),
        ('Friction', 'Darcy Weisbach with the Altshul Tsal friction factor '
                     '(explicit approximation to Colebrook White), checked against '
                     'the RJA SP loss worksheet.'),
        ('Air density', '%g lb/ft3, standard air at sea level, not altitude '
                        'corrected. Editable under Advanced.'
                        % hvac_graph.STANDARD_AIR_DENSITY_LB_FT3),
        ('Roughness', '%g ft galvanized, %g ft flex (flex detected by category), '
                      'plus one flex bend per flex run. Editable under Advanced.'
                      % (hvac_graph.STANDARD_RIGID_ROUGHNESS_FT,
                         hvac_graph.STANDARD_FLEX_ROUGHNESS_FT)),
        ('Fittings', 'C x Pv, with C looked up per fitting from the 2001 ASHRAE '
                     'Fundamentals Ch. 34 tables at its own angle, radius, aspect, '
                     'area and flow ratios (tables shown under Advanced). Round '
                     'elbows: die stamped r/D 1.5 (CD3-1), 1 D corrected by CD3-12. '
                     'Rect mitered elbows: no vanes to 45 deg (CR3-6), vanes above '
                     '(CR3-9). Transitions at the section o Pv, taps at the common '
                     'section Pv. RJA SP loss worksheet C where geometry is '
                     'unreadable. One flex bend (C 3.3, ASHRAE Duct Systems Design '
                     'Guide) per flex run.'),
        ('External static', 'The most restrictive run (index run) from the unit to '
                            'a terminal, supply path plus return path. Includes duct '
                            'friction, fittings, balancing dampers, the diffuser and '
                            'the safety factor. Excludes the coil and cabinet (already '
                            'in the published ESP), fire and backdraft dampers, and the '
                            'filter unless Include return filter is ticked.'),
    ] + vav_checks.basis_rows()
    for r, (lab, val) in enumerate(ref_rows):
        _add_row_def(grid_ref)
        a = TextBlock()
        a.Text = lab
        a.FontWeight = _FW.SemiBold
        a.Margin = Thickness(0, 2, 12, 2)
        a.VerticalAlignment = VerticalAlignment.Top
        _put(grid_ref, a, 0, r)
        v = ui_helpers.rja_hint_block(win, val)
        v.Margin = Thickness(0, 2, 0, 2)
        _put(grid_ref, v, 1, r)

    # ── grey out the external static group until its checkbox is ticked ─────
    def sync_enable(s=None, e=None):
        master = bool(cb_static.IsChecked)
        pnl_static.IsEnabled = master
        pnl_static.Opacity = 1.0 if master else 0.55
        tb_filter.IsEnabled = bool(cb_filter.IsChecked)
        # VAV box and static fields appear only with VAV and external static
        # both ticked (Colin, 2026-10-22).
        both = master and bool(cb_vav.IsChecked)
        pnl_vav_static.Visibility = Visibility.Visible if both else Visibility.Collapsed
    cb_static.Checked += sync_enable
    cb_static.Unchecked += sync_enable
    cb_filter.Checked += sync_enable
    cb_filter.Unchecked += sync_enable

    # The VAV section stays greyed out until VAV System is ticked.
    def sync_vav(s=None, e=None):
        on = bool(cb_vav.IsChecked)
        exp_vav.IsEnabled = on
        exp_vav.Opacity = 1.0 if on else 0.55
        if not on:
            exp_vav.IsExpanded = False
        sync_enable()
    cb_vav.Checked += sync_vav
    cb_vav.Unchecked += sync_vav
    sync_enable()
    sync_vav()

    # ── OK: validate every field, report inline, window stays open ──────────
    def _collect():
        """Returns (errors, values). errors: list of (box, message, expanders)."""
        errors = []

        def num(box, label, kind, expanders=()):
            try:
                v = float(box.Text)
                if math.isnan(v) or math.isinf(v):
                    raise ValueError('not finite')
            except (ValueError, TypeError):
                errors.append((box, label + ' must be a number.', expanders))
                return None
            if kind == 'open_pct' and not (0 < v < 100):
                errors.append((box, label + ' must be above 0 and below 100.', expanders))
                return None
            if kind == 'pct' and not (0 <= v <= 100):
                errors.append((box, label + ' must be between 0 and 100.', expanders))
                return None
            if kind == 'pos' and v <= 0:
                errors.append((box, label + ' must be greater than 0.', expanders))
                return None
            if kind == 'nonneg' and v < 0:
                errors.append((box, label + ' cannot be negative.', expanders))
                return None
            return v

        gpct = num(tb_tol, 'Yellow tolerance', 'open_pct')
        out = {}
        for i, (sys_class, _d1, _d2) in enumerate(_DV_ROWS):
            max_fpm = num(vel_boxes[i], sys_class + ' max velocity', 'pos')
            max_fric = num(fric_boxes[i], sys_class + ' max friction', 'pos')
            if max_fpm is not None and max_fric is not None:
                out[sys_class] = (max_fpm, max_fric)

        ext_static = bool(cb_static.IsChecked)
        if ext_static:
            safety_pct = num(tb_sf, 'Safety factor', 'pct')
        else:
            # Box is greyed out and not used. Still returned, as before; an
            # unusable value falls back to the default rather than blocking OK.
            try:
                safety_pct = float(tb_sf.Text)
                if math.isnan(safety_pct) or not (0 <= safety_pct <= 100):
                    safety_pct = float(_DV_DEFAULT_SAFETY_PCT)
            except (ValueError, TypeError):
                safety_pct = float(_DV_DEFAULT_SAFETY_PCT)

        c_values = {}
        for k, label in c_order:
            v = num(c_boxes[k], 'Fitting C value, ' + label, 'nonneg', (exp_adv, exp_c))
            if v is not None:
                c_values[k] = v
        comp_values = {}
        for k, label in comp_order:
            v = num(comp_boxes[k], label + ' pressure drop', 'nonneg', (exp_adv, exp_comp))
            if v is not None:
                comp_values[k] = v
        calc_basis = {}
        for k, label in basis_order:
            v = num(basis_boxes[k], label, 'pos', (exp_adv, exp_basis))
            if v is not None:
                calc_basis[k] = v

        filter_inwc = 0.0
        if ext_static and bool(cb_filter.IsChecked):
            f = num(tb_filter, 'Return filter drop', 'nonneg')
            if f is not None:
                filter_inwc = f

        vav = bool(cb_vav.IsChecked)
        vav_settings = vav_checks.default_settings()
        if vav:
            fields = [(f, (exp_vav,)) for f in vav_checks.VAV_LIMIT_FIELDS]
            if ext_static:
                fields += [(f, ()) for f in vav_checks.VAV_STATIC_FIELDS]
            for (key, label, _unit, _default, kind), exps in fields:
                v, err = vav_checks.parse_field(vav_boxes[key].Text, kind)
                if err:
                    errors.append((vav_boxes[key], label + ' ' + err + '.', exps))
                else:
                    vav_settings[key] = v

        selected_cols = set(k for k, cb in col_boxes.items() if bool(cb.IsChecked))
        full_diag = bool(cb_full_diag.IsChecked)
        values = _settings_result(out, gpct, selected_cols, full_diag,
                                  ext_static, safety_pct, c_values, comp_values,
                                  calc_basis, filter_inwc, vav, vav_settings)
        return errors, values

    def on_ok(s, e):
        try:
            for b in all_boxes:
                ui_helpers.rja_clear_error(b)
            tb_error.Text = ''
            errors, values = _collect()
            if errors:
                for box, _msg, exps in errors:
                    ui_helpers.rja_mark_error(box, exps)
                msg = errors[0][1]
                if len(errors) > 1:
                    msg += '  (%d more highlighted)' % (len(errors) - 1)
                tb_error.Text = msg
                try:
                    win.UpdateLayout()
                    errors[0][0].BringIntoView()
                    errors[0][0].Focus()
                except Exception:
                    pass
                return
            result[0] = values
            win.Close()
        except Exception as ex:
            # Never let a handler fault escape into ShowDialog.
            try:
                tb_error.Text = 'Unexpected error: ' + str(ex)
            except Exception:
                pass

    def on_cancel(s, e):
        win.Close()

    btn_ok.Click += on_ok
    btn_cancel.Click += on_cancel

    win.ShowDialog()
    return result[0]


def show_velocity_settings_dialog():
    """The settings dialog. Returns the _settings_result() dict, or None if
    cancelled.

    Tries the themed XAML dialog first, then the same dialog without the theme
    file, then the original code-built dialog, so the tool always opens. The
    meaning of every returned value is documented on
    _show_velocity_settings_dialog_legacy.
    """
    log = script.get_logger()
    for use_theme in (True, False):
        try:
            return _show_velocity_settings_xaml(use_theme)
        except Exception:
            log.warning('Themed settings dialog failed (theme=%s), falling back:\n%s'
                        % (use_theme, traceback.format_exc()))
    return _show_velocity_settings_dialog_legacy()


def _show_velocity_settings_dialog_legacy():
    """FALLBACK ONLY. Code-built WPF dialog — VAV System, per-system max velocity + friction,
    a yellow/red tolerance %, and which columns the output tables show.

    Returns the _settings_result() dict or None. Keys: limits
    ({sys_class: (max_fpm, max_friction_inwc)}, plus vav_checks.MP_KEY when VAV
    is ticked), tol_pct, selected_cols, full_diag, ext_static, safety_pct,
    c_values, comp_values, calc_basis, filter_inwc, vav, vav_settings
    ({vav_checks.VAV_FIELDS key: value}; unit_max_esp None means no check).

    filter_inwc is 0.0 unless "Include return filter" is ticked, else the known
    drop in in. wc, added once to the RETURN path total.

    c_values maps fitting_tables.C_TABLE keys to C, and comp_values maps
    COMPONENT_TABLE keys to in. wc. Both start at the published defaults and are
    only changed if the engineer opens the expander and edits one.

    calc_basis maps 'rigid_roughness' / 'flex_roughness' (ft) and 'air_density'
    (lb/ft3) to the Advanced settings values, passed to hvac_graph.set_calc_basis().

    The velocity and friction limits here apply to MAIN ducts only. Branch
    ducts (a run feeding exactly one terminal) are sized against the published
    diffuser tables instead and ignore every number on this dialog — see the
    module docstring. Slot-diffuser branches are the one exception and do use
    these limits.

    The run is always equipment-level (Supply Air + Return Air only, any
    mechanical equipment node's Outside Air branch skipped by
    hvac_graph.traverse()). The old disabled "Multi System (pending)" box is
    now the VAV System box (VAV plan, Colin 2026-10-22).

    Color bands:
      Green  : value <= max
      Yellow : max < value <= max * (1 + tol_pct/100)
      Red    : value > max * (1 + tol_pct/100)
      Purple : determined separately — installed duct could be swapped for a
               smaller standard size and still meet max FPM and max friction
               with no tolerance band (see _suggest_size_info).
    """
    # Defaults: firm design standard. Main ducts only — branches no longer
    # share these values (see the docstring above).
    ROWS = list(_DV_ROWS)
    DEFAULT_TOL_PCT = _DV_DEFAULT_TOL_PCT
    DEFAULT_SAFETY_PCT = _DV_DEFAULT_SAFETY_PCT
    # Component drops are NOT defined here: the rows below are built straight
    # from fitting_tables.COMPONENT_TABLE so there is exactly one place to
    # change a default. Provenance for each figure lives beside it there.

    result    = [None]
    vel_boxes  = {}   # row_idx -> TextBox (velocity)
    fric_boxes = {}   # row_idx -> TextBox (friction)
    gpct_box   = [None]

    WIN_WIDTH   = 720
    CONTENT_W   = WIN_WIDTH - 28 - 20   # win width minus outer margin minus a little slack

    win = Window()
    win.Title  = 'Duct Performance Settings'
    win.Width  = WIN_WIDTH
    win.SizeToContent = SizeToContent.Height
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    outer = StackPanel()
    outer.Margin = Thickness(14)

    def _section_title(text, top=0):
        """One bold section title. Defined once so every section matches."""
        tb = TextBlock()
        tb.Text = text
        tb.FontWeight = FontWeights.Bold
        tb.Margin = Thickness(0, top, 0, 6)
        outer.Children.Add(tb)
        return tb

    # 'System Settings' covers the two checkboxes below it (Full System and
    # Include Outside Air). Colin 2026-09-24: "will the first title including the
    # first two check marks being System Settings".
    _section_title('System Settings')

    # ── 1. Full System Diagnostic ───────────────────────────────────────────
    # Ahead of every other option on purpose: it changes what the options
    # below apply TO (every duct, not just the flagged ones), so it has to be
    # seen and decided first.
    cb_full_diag = CheckBox()
    cb_full_diag_text = TextBlock()
    cb_full_diag_text.Text = (
        'Full System.  Unchecked = only failing ducts '
        '(Red, Yellow, Purple).')
    cb_full_diag_text.TextWrapping = TextWrapping.Wrap
    cb_full_diag_text.Width = CONTENT_W - 20
    cb_full_diag.Content   = cb_full_diag_text
    cb_full_diag.IsChecked = False
    cb_full_diag.FontWeight = FontWeights.Bold
    cb_full_diag.Margin = Thickness(2, 0, 0, 2)
    outer.Children.Add(cb_full_diag)

    full_diag_sep = Separator()
    full_diag_sep.Margin = Thickness(0, 8, 0, 10)
    outer.Children.Add(full_diag_sep)

    # ── 2. VAV System, with its fields greyed out until ticked ─────────────
    cb_vav = CheckBox()
    cb_vav_text = TextBlock()
    cb_vav_text.Text = ('VAV System.  Supply mains between the unit and the box '
                        'inlets are checked as medium pressure (MP).')
    cb_vav_text.TextWrapping = TextWrapping.Wrap
    cb_vav_text.Width = CONTENT_W - 20
    cb_vav.Content = cb_vav_text
    cb_vav.IsChecked = False
    cb_vav.Margin  = Thickness(2, 0, 0, 2)
    outer.Children.Add(cb_vav)

    vav_boxes = {}

    def _vav_rows(panel, fields):
        for key, label, unit, default, _kind in fields:
            rp = StackPanel()
            rp.Orientation = Orientation.Horizontal
            rp.Margin = Thickness(0, 1, 0, 1)
            lb = TextBlock()
            lb.Text = label
            lb.Width = 250
            lb.VerticalAlignment = VerticalAlignment.Center
            rp.Children.Add(lb)
            tb = TextBox()
            tb.Text = vav_checks.format_default(default)
            tb.Width = 55
            tb.Margin = Thickness(4, 0, 6, 0)
            tb.VerticalAlignment = VerticalAlignment.Center
            rp.Children.Add(tb)
            ex = TextBlock()
            ex.Text = unit
            ex.Foreground = SolidColorBrush(Colors.DimGray)
            ex.VerticalAlignment = VerticalAlignment.Center
            rp.Children.Add(ex)
            panel.Children.Add(rp)
            vav_boxes[key] = tb

    vav_panel = StackPanel()
    vav_panel.Margin = Thickness(22, 2, 0, 2)
    _vav_rows(vav_panel, vav_checks.VAV_LIMIT_FIELDS)
    outer.Children.Add(vav_panel)

    # Box and static fields, placed under the external static option below and
    # shown only with VAV and external static both ticked (Colin, 2026-10-22).
    vav_static_panel = StackPanel()
    vav_static_panel.Margin = Thickness(22, 2, 0, 6)
    _vav_rows(vav_static_panel, vav_checks.VAV_STATIC_FIELDS)

    def _sync_vav(s=None, e=None):
        on = bool(cb_vav.IsChecked)
        vav_panel.IsEnabled = on
        vav_panel.Opacity = 1.0 if on else 0.55
        both = on and bool(cb_static.IsChecked)
        vav_static_panel.Visibility = Visibility.Visible if both else Visibility.Collapsed

    oa_sep = Separator()
    oa_sep.Margin = Thickness(0, 8, 0, 10)
    outer.Children.Add(oa_sep)

    # An overarching bold title rather than a collapsed Expander. Colin tried the
    # Expander and rejected it 2026-09-24: "i actually dont like the max velocity
    # drop per system bar i would still like an over arching title." The window
    # went from 520 to 720 wide and the rows got tighter instead, so the whole
    # dialog fits without hiding anything.
    _section_title('Max Velocity and Friction per System')

    intro = Label()
    intro.Content = 'Max velocity (FPM) and pressure drop (in. wc/100 ft) per system:'
    intro.Margin  = Thickness(0, 0, 0, 8)
    outer.Children.Add(intro)

    # 3-column grid: system | velocity | friction
    grid = Grid()
    for w in (200, 180, 220):
        cd = ColumnDefinition()
        cd.Width = GridLength(w)
        grid.ColumnDefinitions.Add(cd)
    for _ in range(len(ROWS) + 1):
        rd = RowDefinition()
        rd.Height = GridLength(26)
        grid.RowDefinitions.Add(rd)

    def _lbl(text, col, row):
        lb = Label()
        lb.Content = text
        lb.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(lb, col)
        Grid.SetRow(lb, row)
        grid.Children.Add(lb)

    _lbl('System Type',          0, 0)
    _lbl('Max Velocity (FPM)',   1, 0)
    _lbl('Max Friction (iwc/100)', 2, 0)

    for i, (sys_class, def_fpm, def_fric) in enumerate(ROWS):
        r = i + 1
        _lbl(sys_class, 0, r)
        for col, val, store in ((1, def_fpm, vel_boxes), (2, def_fric, fric_boxes)):
            tb = TextBox()
            tb.Text  = str(val)
            tb.Width = 80
            tb.Margin = Thickness(4, 4, 4, 4)
            tb.VerticalAlignment   = VerticalAlignment.Center
            tb.HorizontalAlignment = HorizontalAlignment.Left
            Grid.SetColumn(tb, col)
            Grid.SetRow(tb, r)
            grid.Children.Add(tb)
            store[i] = tb

    outer.Children.Add(grid)

    # Green threshold row
    gpct_panel = StackPanel()
    gpct_panel.Orientation = Orientation.Horizontal
    gpct_panel.Margin = Thickness(0, 10, 0, 0)

    gpct_lbl = Label()
    gpct_lbl.Content = 'Yellow tolerance:'
    gpct_lbl.VerticalAlignment = VerticalAlignment.Center
    gpct_panel.Children.Add(gpct_lbl)

    tb_gpct = TextBox()
    tb_gpct.Text  = str(DEFAULT_TOL_PCT)
    tb_gpct.Width = 45
    tb_gpct.Margin = Thickness(4, 0, 4, 0)
    tb_gpct.VerticalAlignment = VerticalAlignment.Center
    gpct_panel.Children.Add(tb_gpct)
    gpct_box[0] = tb_gpct

    gpct_suffix_inline = Label()
    gpct_suffix_inline.Content = '% above max before red'
    gpct_suffix_inline.VerticalAlignment = VerticalAlignment.Center
    gpct_panel.Children.Add(gpct_suffix_inline)
    outer.Children.Add(gpct_panel)

    gpct_suffix = TextBlock()
    gpct_suffix.Text = 'At or under max = green, within tolerance = yellow, past it = red.'
    gpct_suffix.TextWrapping = TextWrapping.Wrap
    gpct_suffix.Width = CONTENT_W
    gpct_suffix.Foreground = SolidColorBrush(Colors.DimGray)
    gpct_suffix.Margin = Thickness(0, 2, 0, 0)
    outer.Children.Add(gpct_suffix)


    # ── Column picker ──────────────────────────────────────────────────────
    # Drives BOTH output tables identically (see _COLUMN_DEFS). Only the
    # optional columns appear here; the rest are always in the tables.
    col_sep = Separator()
    col_sep.Margin = Thickness(0, 12, 0, 8)
    outer.Children.Add(col_sep)

    # ── 4. Options ──────────────────────────────────────────────────────────
    _section_title('Options')

    # Total external static pressure: what the fan must develop against
    # everything OUTSIDE the unit, along the single most restrictive run.
    #
    # This supersedes an earlier "sum the Friction Loss column" TOTAL row, which
    # was a stand-in built before fitting losses existed. Colin, 2026-09-24: "we
    # dont need the friction loss option anymore as we have figured out total
    # external pressure." Both that row and the Friction Loss column are gone.
    #
    # External static deliberately stops at the unit casing. Filter, coil and
    # cabinet losses are internal, they come off the manufacturer's cutsheet,
    # and there is nothing in a Revit model to derive them from (Colin,
    # 2026-09-24: "maybe i just need external static pressure not total static
    # pressure ... this i feel like can be found directly in cutsheet").
    cb_static = CheckBox()
    # Label says only what the option IS. Colin, 2026-09-24: "call the option
    # for total external static pressure just that not all the riff raff after
    # that can be included in the calculation basis." What it includes and
    # excludes is stated in Calculation Basis at the bottom of this dialog.
    cb_static_text = TextBlock()
    cb_static_text.Text = 'Total external static pressure'
    cb_static_text.TextWrapping = TextWrapping.Wrap
    cb_static_text.Width = CONTENT_W - 20
    cb_static.Content   = cb_static_text
    cb_static.IsChecked = False
    cb_static.Margin    = Thickness(2, 0, 0, 2)
    outer.Children.Add(cb_static)

    # Safety factor: the last row of RJA's SP_LOSS_WORKSHEET, a flat % on the
    # subtotal (its example uses 10%). Applies to the external static total
    # only, never to the per-duct friction checks.
    sf_panel = StackPanel()
    sf_panel.Orientation = Orientation.Horizontal
    sf_panel.Margin = Thickness(22, 0, 0, 6)
    sf_lbl = Label()
    sf_lbl.Content = 'Safety factor:'
    sf_lbl.VerticalAlignment = VerticalAlignment.Center
    sf_panel.Children.Add(sf_lbl)
    tb_sf = TextBox()
    tb_sf.Text  = str(DEFAULT_SAFETY_PCT)
    tb_sf.Width = 45
    tb_sf.Margin = Thickness(4, 0, 4, 0)
    tb_sf.VerticalAlignment = VerticalAlignment.Center
    sf_panel.Children.Add(tb_sf)
    sf_suffix = Label()
    sf_suffix.Content = '% added to the external static total'
    sf_suffix.VerticalAlignment = VerticalAlignment.Center
    sf_panel.Children.Add(sf_suffix)
    outer.Children.Add(sf_panel)

    # Return filter as a known drop on the return path total. Off by default
    # because the filter is normally inside the unit and already deducted from
    # its published ESP. Colin, 2026-10-08: "Include Return Filter", 0.14 known drop.
    filt_panel = StackPanel()
    filt_panel.Orientation = Orientation.Horizontal
    filt_panel.Margin = Thickness(22, 0, 0, 6)
    cb_filter = CheckBox()
    cb_filter_text = TextBlock()
    cb_filter_text.Text = 'Include return filter'
    cb_filter.Content = cb_filter_text
    cb_filter.IsChecked = False
    cb_filter.VerticalAlignment = VerticalAlignment.Center
    filt_panel.Children.Add(cb_filter)
    tb_filter = TextBox()
    tb_filter.Text  = '%g' % fitting_tables.DEFAULT_RETURN_FILTER_INWC
    tb_filter.Width = 45
    tb_filter.Margin = Thickness(8, 0, 4, 0)
    tb_filter.VerticalAlignment = VerticalAlignment.Center
    filt_panel.Children.Add(tb_filter)
    filt_suffix = Label()
    filt_suffix.Content = 'in. wc known drop (MERV 8), added to the return total'
    filt_suffix.VerticalAlignment = VerticalAlignment.Center
    filt_panel.Children.Add(filt_suffix)
    outer.Children.Add(filt_panel)

    outer.Children.Add(vav_static_panel)
    cb_vav.Checked += _sync_vav
    cb_vav.Unchecked += _sync_vav
    cb_static.Checked += _sync_vav
    cb_static.Unchecked += _sync_vav
    _sync_vav()

    # Fitting C values and component drops are GIVENS with an override, not
    # questions the dialog asks every run. Colin, 2026-09-24: "remove diffuser
    # and balancing drops inputs and have them as givens. Maybe put a drop down
    # menu for all C values under fittings losses and a drop down with input
    # values for components losses ... make them input for fittings as well."
    #
    # An Expander rather than a ComboBox: a ComboBox picks one of a list, and
    # what is wanted here is the whole table visible and editable at once, out
    # of the way until it is needed. Both start collapsed, so the normal run is
    # still one checkbox and a safety factor.
    def _value_expander(header, rows, store, width_label, fmt):
        """Collapsed panel of labelled value boxes. rows: (key, label, extra, default)."""
        exp = Expander()
        exp.Header = header
        exp.IsExpanded = False
        exp.Margin = Thickness(22, 2, 0, 6)
        inner = StackPanel()
        inner.Margin = Thickness(4, 4, 0, 2)
        for row in rows:
            key, label, extra, default = row
            rp = StackPanel()
            rp.Orientation = Orientation.Horizontal
            rp.Margin = Thickness(0, 1, 0, 1)
            lb = TextBlock()
            lb.Text = label
            lb.Width = width_label
            lb.TextWrapping = TextWrapping.NoWrap
            lb.VerticalAlignment = VerticalAlignment.Center
            rp.Children.Add(lb)
            tb = TextBox()
            tb.Text  = fmt % default
            tb.Width = 55
            tb.Margin = Thickness(4, 0, 6, 0)
            tb.VerticalAlignment = VerticalAlignment.Center
            rp.Children.Add(tb)
            if extra:
                ex = TextBlock()
                ex.Text = extra
                ex.Foreground = SolidColorBrush(Colors.DimGray)
                ex.VerticalAlignment = VerticalAlignment.Center
                rp.Children.Add(ex)
            inner.Children.Add(rp)
            store[key] = tb
        exp.Content = inner
        outer.Children.Add(exp)
        return exp

    c_boxes = {}
    _value_expander(
        'Fallback C values (RJA SP loss worksheet, used where geometry is unreadable)',
        [(k, lbl, 'ASHRAE ' + no, v) for k, lbl, no, v in fitting_tables.C_TABLE],
        c_boxes, 250, '%.2f')

    comp_boxes = {}
    _value_expander(
        'Component pressure drops (in. wc each)',
        [(k, lbl, 'in. wc', v) for k, lbl, v in fitting_tables.COMPONENT_TABLE],
        comp_boxes, 250, '%.3f')

    # The calculation basis itself, for the rare job that departs from it.
    # Colin, 2026-10-08: interior-insulated duct has eps 0.005, not 0.0003.
    # Collapsed and at the standard values, so a normal run never touches it.
    basis_boxes = {}
    adv_exp = _value_expander(
        'Advanced settings (calculation basis)',
        [('rigid_roughness', 'Duct roughness, rigid (ft)',
          u'0.0003 galvanized, 0.005 interior-insulated',
          hvac_graph.STANDARD_RIGID_ROUGHNESS_FT),
         ('flex_roughness', 'Duct roughness, flex (ft)',
          u'0.003 flex, fully extended (ASHRAE)',
          hvac_graph.STANDARD_FLEX_ROUGHNESS_FT),
         ('air_density', u'Air density (lb/ft³)',
          u'0.075 standard air, sea level',
          hvac_graph.STANDARD_AIR_DENSITY_LB_FT3)],
        basis_boxes, 250, '%g')
    adv_note = TextBlock()
    adv_note.Text = (u'Applies to EVERY duct in the run, not one system. Leave at the '
                     u'standard values unless the project needs otherwise. Air density '
                     u'also drives velocity pressure (fitting losses).')
    adv_note.TextWrapping = TextWrapping.Wrap
    adv_note.Width = CONTENT_W - 40
    adv_note.Foreground = SolidColorBrush(Colors.DimGray)
    adv_note.Margin = Thickness(0, 0, 0, 4)
    adv_exp.Content.Children.Insert(0, adv_note)

    col_grid = Grid()
    _COL_PICKER_COLS = 2
    for _ in range(_COL_PICKER_COLS):
        cd = ColumnDefinition()
        cd.Width = GridLength(CONTENT_W / float(_COL_PICKER_COLS))
        col_grid.ColumnDefinitions.Add(cd)
    # Ceiling division, written the Python 2 way (// on ints floors).
    _optional = _optional_columns()
    _col_rows = (len(_optional) + _COL_PICKER_COLS - 1) // _COL_PICKER_COLS
    for _ in range(_col_rows):
        rd = RowDefinition()
        rd.Height = GridLength(22)
        col_grid.RowDefinitions.Add(rd)

    col_boxes = {}   # column key -> CheckBox
    for i, (key, header, _vw, _cw, default_on) in enumerate(_optional):
        cb = CheckBox()
        cb_txt = TextBlock()
        cb_txt.Text = header
        cb_txt.TextWrapping = TextWrapping.NoWrap
        cb.Content   = cb_txt
        cb.IsChecked = default_on
        cb.Margin    = Thickness(2, 2, 2, 2)
        cb.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(cb, i // _col_rows)
        Grid.SetRow(cb, i % _col_rows)
        col_grid.Children.Add(cb)
        col_boxes[key] = cb

    outer.Children.Add(col_grid)

    # Assumptions / formula reference block
    sep = Separator()
    sep.Margin = Thickness(0, 12, 0, 8)
    outer.Children.Add(sep)

    _INFO_LBL_W = 150

    def _info_row(label_text, value_text):
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.Margin = Thickness(0, 1, 0, 1)
        lbl = TextBlock()
        lbl.Text = label_text
        lbl.Width = _INFO_LBL_W
        lbl.FontWeight = FontWeights.Bold
        lbl.Foreground = SolidColorBrush(Colors.DimGray)
        val = TextBlock()
        val.Text = value_text
        val.TextWrapping = TextWrapping.Wrap
        val.Width = CONTENT_W - _INFO_LBL_W
        val.Foreground = SolidColorBrush(Colors.DimGray)
        row.Children.Add(lbl)
        row.Children.Add(val)
        outer.Children.Add(row)

    _section_title('Calculation Basis')

    _info_row('Velocity / friction:', 'checked on MAIN ducts only. Branches are judged '
                                      'against the diffuser capacity tables instead.')
    _info_row('Pressure drop:',   u'Darcy-Weisbach:  ΔP/ft = f/Dh × ρV²/2g')
    _info_row('Friction factor:', u'Altshul-Tsal  (ASHRAE explicit approx. to Colebrook-White)')
    _info_row('Verified against:', u'RJA SP_LOSS_WORKSHEET, matches its duct rows to the printed digit')
    _info_row('Air density:',     u'0.0750 lb/ft³  (standard air, 68°F, SEA LEVEL, '
                                  u'not altitude-corrected). Editable under Advanced settings.')
    _info_row('Duct roughness:',  u'ε = 0.0003 ft galvanized, 0.003 ft flex '
                                  u'(detected by category, ASHRAE Table 1-2). '
                                  u'Editable under Advanced settings.')
    _info_row('Fitting losses:',  u'C × Pv, C from RJA SP_LOSS_WORKSHEET '
                                  u'(1985 ASHRAE fitting numbers), plus one flex '
                                  u'bend (C 3.3) per flex run. Return and exhaust '
                                  u'transitions priced in the airflow direction. '
                                  u'Editable above')
    _info_row('External static:', 'the single most restrictive run (index run) from '
                                  'the unit to a terminal, supply path + return path')
    _info_row('  includes:',      'duct friction, fittings (elbows, take-offs, '
                                  'transitions), balancing dampers, the diffuser, '
                                  'and the safety factor')
    _info_row('  excludes:',      'filter (unless Include return filter is ticked), coil '
                                  'and cabinet (already in the published ESP). Fire and '
                                  'backdraft dampers.')
    for _lab, _val in vav_checks.basis_rows():
        _info_row(_lab + ':', _val)

    # OK / Cancel
    btn_panel = StackPanel()
    btn_panel.Orientation = Orientation.Horizontal
    btn_panel.HorizontalAlignment = HorizontalAlignment.Right
    btn_panel.Margin = Thickness(0, 14, 0, 0)

    ok_btn = Button()
    ok_btn.Content = 'OK'
    ok_btn.Width   = 72
    ok_btn.Margin  = Thickness(0, 0, 8, 0)

    cancel_btn = Button()
    cancel_btn.Content = 'Cancel'
    cancel_btn.Width   = 72

    def on_ok(s, e):
        try:
            gpct = float(gpct_box[0].Text)
            if not (0 < gpct < 100):
                forms.alert('Yellow tolerance must be between 0 and 100.', title='Invalid Input')
                return
            out = {}
            for i, (sys_class, _, _) in enumerate(ROWS):
                max_fpm  = float(vel_boxes[i].Text)
                max_fric = float(fric_boxes[i].Text)
                if max_fpm <= 0 or max_fric <= 0:
                    forms.alert('All values must be greater than 0.', title='Invalid Input')
                    return
                out[sys_class] = (max_fpm, max_fric)
            vav = bool(cb_vav.IsChecked)
            vav_settings = vav_checks.default_settings()
            if vav:
                fields = vav_checks.VAV_LIMIT_FIELDS
                if bool(cb_static.IsChecked):
                    fields = vav_checks.VAV_FIELDS
                for key, label, _unit, _default, kind in fields:
                    v, err = vav_checks.parse_field(vav_boxes[key].Text, kind)
                    if err:
                        forms.alert('{} {}.'.format(label, err), title='Invalid Input')
                        return
                    vav_settings[key] = v
            selected_cols = set(k for k, cb in col_boxes.items() if bool(cb.IsChecked))
            full_diag = bool(cb_full_diag.IsChecked)
            ext_static = bool(cb_static.IsChecked)
            safety_pct = float(tb_sf.Text)
            if safety_pct < 0 or safety_pct > 100:
                forms.alert('Safety factor must be between 0 and 100.',
                            title='Invalid Input')
                return
            c_values = {}
            for k, box in c_boxes.items():
                v = float(box.Text)
                if v < 0:
                    forms.alert('Fitting C values cannot be negative.',
                                title='Invalid Input')
                    return
                c_values[k] = v
            comp_values = {}
            for k, box in comp_boxes.items():
                v = float(box.Text)
                if v < 0:
                    forms.alert('Component pressure drops cannot be negative.',
                                title='Invalid Input')
                    return
                comp_values[k] = v
            calc_basis = {}
            for k, box in basis_boxes.items():
                v = float(box.Text)
                if v <= 0:
                    forms.alert('Advanced settings (roughness, air density) must '
                                'be greater than 0.', title='Invalid Input')
                    return
                calc_basis[k] = v
            filter_inwc = 0.0
            if bool(cb_filter.IsChecked):
                filter_inwc = float(tb_filter.Text)
                if filter_inwc < 0:
                    forms.alert('The return filter drop cannot be negative.',
                                title='Invalid Input')
                    return
            result[0] = _settings_result(out, gpct, selected_cols, full_diag,
                                         ext_static, safety_pct, c_values,
                                         comp_values, calc_basis, filter_inwc,
                                         vav, vav_settings)
        except ValueError:
            forms.alert('Enter valid numbers for all fields.', title='Invalid Input')
            return
        win.Close()

    def on_cancel(s, e):
        win.Close()

    ok_btn.Click     += on_ok
    cancel_btn.Click += on_cancel
    btn_panel.Children.Add(ok_btn)
    btn_panel.Children.Add(cancel_btn)
    outer.Children.Add(btn_panel)

    # Bounded and scrollable rather than free-growing. SizeToContent.Height
    # still sizes the window to its content, but only up to MaxHeight, past
    # which the ScrollViewer takes over - so the OK/Cancel buttons can never end
    # up below the bottom of the screen no matter what gets added later.
    scroll = ScrollViewer()
    scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
    scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
    scroll.Content = outer
    win.Content = scroll
    try:
        # Usable desktop area, so the taskbar is excluded.
        win.MaxHeight = SystemParameters.WorkArea.Height * 0.9
    except Exception:
        win.MaxHeight = 800.0
    win.ShowDialog()
    return result[0]


# ── system picker (second screen) ────────────────────────────────────
#
# Replaces forms.SelectFromList so the screen can carry a SECOND action button.
# The equipment list only ever holds equipment found in the ACTIVE VIEW, so a
# system whose AHU sits on another floor, in a mechanical room that is not on
# this plan, or outside the view crop used to be unreachable. "Select a Duct
# System" lets the user point at the ductwork instead (Colin 2026-09-25): two
# clicks, either order, supply or return, no roles to declare.
#
# Picking a duct does NOT change how the run is rooted: build_network still
# walks to that duct's base equipment and roots there, so CFM direction and
# every downstream sum stay correct. It only changes how the system is NAMED
# to the tool.
class _DuctPickFilter(ISelectionFilter):
    """Restricts picking to rigid and flex duct, so nothing else is clickable."""

    def AllowElement(self, elem):
        # hvac_graph.is_duct() already covers rigid AND flex duct, and is the
        # same test the traversal uses, so the filter can never allow something
        # the traversal would then refuse to walk.
        try:
            return hvac_graph.is_duct(elem)
        except Exception:
            return False

    def AllowReference(self, ref, point):
        return False


def _pick_duct_system():
    """Pick two ducts, in any order, supply or return - no roles assigned.

    Colin 2026-09-25: "it should just be click two ducts no particular order it
    can be return or supply". Nothing downstream needs to know which is which:
    each pick is rooted at its own base equipment and the system class is read
    off the ductwork itself, so making the user label them would be busywork
    that could only be got wrong.

    Returns a list of picked duct elements (1 or 2), or [] if cancelled on the
    first pick. Esc on the SECOND pick is not a cancel, it means "just the one",
    which is a real case (a single ducted system, or a return air plenum with no
    ducted return to click).
    """
    try:
        filt = _DuctPickFilter()
    except Exception:
        filt = None
        output.print_md(':warning: Duct selection filter unavailable, picking '
                        'is unfiltered; click a duct, not a fitting.')

    picked = []
    prompts = [
        'Select a duct in the system to analyze (supply or return, either one)',
        'Select a second duct on the other system, or press Esc to run just the first',
    ]
    for i, prompt in enumerate(prompts):
        try:
            if filt is not None:
                ref = uidoc.Selection.PickObject(ObjectType.Element, filt, prompt)
            else:
                ref = uidoc.Selection.PickObject(ObjectType.Element, prompt)
        except Exception:
            break
        elem = doc.GetElement(ref.ElementId)
        if elem is None:
            break
        picked.append(elem)
        output.print_md('Duct {} picked: **{}** (id {})'.format(
            i + 1, _elem_name(elem), eid_int(elem.Id)))
    return picked


def _pick_side(duct):
    """'supply', 'return' or None, from the duct's System Type name.

    Exhaust counts as the return side, matching how the external static adds
    the worst of Return/Exhaust to Supply. None when the name says neither, so
    an unusual System Type name never produces a false warning.
    """
    try:
        name = (hvac_graph.duct_sys_class(duct) or '').lower()
    except Exception:
        return None
    if vav_checks.is_supply_class(name):
        return 'supply'
    if 'return' in name or 'exhaust' in name:
        return 'return'
    return None


def _warn_if_same_side(picked):
    """Fail-safe for the single-system route: one supply duct and one return
    duct make one system. Two of the same side cannot, so say so (Colin,
    2026-10-08). A warning, not a stop: the run still goes ahead."""
    if len(picked) != 2:
        return
    a, b = _pick_side(picked[0]), _pick_side(picked[1])
    if a is not None and a == b:
        output.print_md(
            ':warning: **Both picked ducts are {} ducts.** A single system needs '
            'one supply duct and one return duct, so the external static below '
            'will not be a complete supply + return total. Run again and pick '
            'one of each.'.format(a))


_PICKER_DUCT_TAG = '__select_duct_system__'

_PICKER_XAML = '''
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Systems to Visualize"
        Width="520" MinWidth="420" MaxWidth="760"
        SizeToContent="Height" ResizeMode="CanResizeWithGrip"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        Background="White" FontFamily="Segoe UI" FontSize="12"
        Foreground="{DynamicResource RjaTextBrush}"
        UseLayoutRounding="True" SnapsToDevicePixels="True">
    @@RESOURCES@@
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        @@HEADER@@

        <StackPanel Grid.Row="1" Margin="16,12,16,12">
            <TextBlock x:Name="tb_picker_hint" Style="{DynamicResource RjaHint}" FontSize="11"
                       TextWrapping="Wrap" Margin="0,0,0,8"/>
            <ListBox x:Name="lst" SelectionMode="Extended" MaxHeight="300"
                     BorderBrush="#D0D0D0" BorderThickness="1"/>
        </StackPanel>

        <Border Grid.Row="2" Background="{DynamicResource RjaPanelBrush}"
                BorderBrush="{DynamicResource RjaRuleBrush}" BorderThickness="0,1,0,0"
                Padding="16,10,16,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="tb_error" Foreground="#C62828" TextWrapping="Wrap"
                           VerticalAlignment="Center" Margin="0,0,12,0"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button x:Name="btn_cancel" Content="Cancel" IsCancel="True"
                            Style="{DynamicResource RjaSecondaryButton}" MinWidth="88" Margin="0,0,8,0"/>
                    <Button x:Name="btn_ok" Content="OK" IsDefault="True"
                            Style="{DynamicResource RjaPrimaryButton}" MinWidth="88"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
'''


def _show_system_picker_xaml(display_names, use_theme):
    """Themed picker: ONE list, duct system entry first, equipment below it.
    Returns ('equipment', [names]), ('ducts', None) or None. Raises if the
    window cannot be built (caller falls back)."""
    import ui_helpers
    from System.Windows.Controls import ListBoxItem
    from System.Windows import FontWeights as _FW

    xaml = (_PICKER_XAML
            .replace('@@RESOURCES@@', ui_helpers.rja_window_resources_xaml(use_theme))
            .replace('@@HEADER@@', ui_helpers.rja_header_xaml(
                'Duct Performance', 'Select the system to analyze.')))
    win = forms.WPFWindow(xaml, literal_string=True)

    N = lambda name: _wpf_named(win, name)
    lst = N('lst')
    tb_hint = N('tb_picker_hint')
    tb_error = N('tb_error')
    btn_ok = N('btn_ok')
    btn_cancel = N('btn_cancel')
    ui_helpers.rja_apply_logo(win, N('img_logo'))
    try:
        win.MaxHeight = SystemParameters.WorkArea.Height * 0.9
    except Exception:
        win.MaxHeight = 800.0

    result = [None]

    if display_names:
        tb_hint.Text = ('Select one or more units found in this view, or choose '
                        'Select Duct System if the unit is not on this plan.')
    else:
        tb_hint.Text = ('No mechanical equipment found in this view. Use Select '
                        'Duct System.')

    # First entry: the duct system route, with a one line hint under it.
    duct_item = ListBoxItem()
    duct_item.Tag = _PICKER_DUCT_TAG
    duct_panel = StackPanel()
    duct_title = TextBlock()
    duct_title.Text = 'Select Duct System (SA/RA)'
    duct_title.FontWeight = _FW.SemiBold
    duct_panel.Children.Add(duct_title)
    duct_hint = ui_helpers.rja_hint_block(
        win, 'Pick one supply duct and one return duct. Use when the unit is not '
             'on this plan.')
    duct_hint.Margin = Thickness(0, 2, 0, 0)
    duct_panel.Children.Add(duct_hint)
    duct_item.Content = duct_panel
    lst.Items.Add(duct_item)

    equip_items = []
    for n in display_names:
        it = ListBoxItem()
        it.Content = n
        it.Tag = n
        lst.Items.Add(it)
        equip_items.append(it)

    # Preselect: first unit if there are any (as before), else the duct entry.
    first = equip_items[0] if equip_items else duct_item
    first.IsSelected = True

    # The duct system entry and unit entries exclude each other: selecting one
    # kind clears the other.
    busy = [False]

    def _is_duct(item):
        return item.Tag == _PICKER_DUCT_TAG

    def on_selection_changed(s, e):
        if busy[0]:
            return
        busy[0] = True
        try:
            added = list(e.AddedItems)
            if any(_is_duct(i) for i in added):
                for i in list(lst.SelectedItems):
                    if not _is_duct(i):
                        lst.SelectedItems.Remove(i)
            elif added:
                for i in list(lst.SelectedItems):
                    if _is_duct(i):
                        lst.SelectedItems.Remove(i)
            tb_error.Text = ''
        finally:
            busy[0] = False
    lst.SelectionChanged += on_selection_changed

    def on_ok(s, e):
        try:
            sel = list(lst.SelectedItems)
            if not sel:
                tb_error.Text = 'Select at least one entry.'
                return
            if any(_is_duct(i) for i in sel):
                result[0] = ('ducts', None)
            else:
                result[0] = ('equipment', [str(i.Tag) for i in sel])
            win.Close()
        except Exception as ex:
            try:
                tb_error.Text = 'Unexpected error: ' + str(ex)
            except Exception:
                pass

    def on_cancel(s, e):
        win.Close()

    def on_loaded(s, e):
        try:
            first.Focus()
        except Exception:
            pass

    btn_ok.Click += on_ok
    btn_cancel.Click += on_cancel
    win.Loaded += on_loaded
    win.ShowDialog()
    return result[0]


def show_system_picker(display_names):
    """Second screen. Returns ('equipment', [names]), ('ducts', None) or None.

    display_names is the sorted list of equipment labels found in the active
    view; it may be empty, in which case only the duct system entry is shown.
    Falls back to the original two button picker if the themed one cannot be
    built.
    """
    log = script.get_logger()
    for use_theme in (True, False):
        try:
            return _show_system_picker_xaml(display_names, use_theme)
        except Exception:
            log.warning('Themed system picker failed (theme=%s), falling back:\n%s'
                        % (use_theme, traceback.format_exc()))
    return _show_system_picker_legacy(display_names)


def _show_system_picker_legacy(display_names):
    """FALLBACK ONLY. Original two button picker.

    Returns ('equipment', [names]), ('ducts', None) or None.
    """
    result = [None]

    win = Window()
    win.Title = 'Select Systems to Visualize'
    win.Width = 520
    win.SizeToContent = SizeToContent.Height
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    outer = StackPanel()
    outer.Margin = Thickness(14)

    head = TextBlock()
    head.Text = ('Equipment found in the active view. Select one or more, or '
                 'select a duct system instead if the equipment is not on this '
                 'plan.')
    head.TextWrapping = TextWrapping.Wrap
    head.Margin = Thickness(0, 0, 0, 10)
    outer.Children.Add(head)

    lst = ListBox()
    lst.SelectionMode = SelectionMode.Extended
    lst.MaxHeight = 300
    for n in display_names:
        lst.Items.Add(n)
    if display_names:
        lst.SelectedIndex = 0
    else:
        lst.IsEnabled = False
        empty = TextBlock()
        empty.Text = ('No mechanical equipment in this view. Use Select a Duct '
                      'System.')
        empty.TextWrapping = TextWrapping.Wrap
        empty.Margin = Thickness(0, 0, 0, 8)
        outer.Children.Add(empty)
    outer.Children.Add(lst)

    btn_panel = StackPanel()
    btn_panel.Orientation = Orientation.Horizontal
    btn_panel.HorizontalAlignment = HorizontalAlignment.Right
    btn_panel.Margin = Thickness(0, 14, 0, 0)

    duct_btn = Button()
    duct_btn.Content = 'Select a Duct System'
    duct_btn.Width   = 150
    duct_btn.Margin  = Thickness(0, 0, 8, 0)

    run_btn = Button()
    run_btn.Content = 'Run Duct Performance'
    run_btn.Width   = 128
    run_btn.Margin  = Thickness(0, 0, 8, 0)
    run_btn.IsEnabled = bool(display_names)

    cancel_btn = Button()
    cancel_btn.Content = 'Cancel'
    cancel_btn.Width   = 72

    def on_run(s, e):
        names = [str(i) for i in lst.SelectedItems]
        if not names:
            forms.alert('Select at least one system, or use Select a Duct '
                        'System.', title='Nothing Selected')
            return
        result[0] = ('equipment', names)
        win.Close()

    def on_ducts(s, e):
        result[0] = ('ducts', None)
        win.Close()

    def on_cancel(s, e):
        win.Close()

    run_btn.Click    += on_run
    duct_btn.Click   += on_ducts
    cancel_btn.Click += on_cancel
    btn_panel.Children.Add(duct_btn)
    btn_panel.Children.Add(run_btn)
    btn_panel.Children.Add(cancel_btn)
    outer.Children.Add(btn_panel)

    win.Content = outer
    win.ShowDialog()
    return result[0]


# ── helpers ────────────────────────────────────────────────────────────────────
_PRIORITY = {'RED': 4, 'YELLOW': 3, 'PURPLE': 2, 'GREEN': 1, 'GRAY': 0}


# Character budget for one line of the System Summary block. The block is drawn
# at the flagged-duct table's width, and TextNote.Create with no width
# constraint does NOT wrap, so a longer line runs off the right edge of the chart
# instead of folding. Sized just under the shortest lines already known to fit.
SUMMARY_WRAP_CHARS = 92


def _wrap_summary_lines(lines, width=SUMMARY_WRAP_CHARS):
    """Fold over-long summary lines at word boundaries, keeping their indent.

    Colin, 2026-09-24: the METHOD and NOT INCLUDED lines "are writing too far out
    and need to be condensed (entered another line) down so it fit within the
    chart."

    A continuation gets its parent's indent plus two, so a folded sentence still
    reads as one item rather than as a new entry. Applied to the whole list
    rather than just the two long lines, so this cannot silently regress the next
    time someone adds a wordy line.
    """
    out = []
    for line in lines:
        if len(line) <= width:
            out.append(line)
            continue
        indent = len(line) - len(line.lstrip(' '))
        cont   = ' ' * (indent + 2)
        cur    = ' ' * indent
        for word in line.split():
            add = word if not cur.strip() else ' ' + word
            if cur.strip() and len(cur) + len(add) > width:
                out.append(cur)
                cur = cont + word
            else:
                cur = cur + add
        if cur.strip():
            out.append(cur)
    return out


def _esp_text_line(item, inwc, detail=None):
    """One external static line of the drafting view summary.

    Item padded to 19 so every value starts in the same column, e.g.
    '      Duct friction      0.123 in. wc   (12 ducts, 140 ft)'.
    """
    line = '      {:<19}{:.3f} in. wc'.format(item, inwc)
    if detail:
        line += '   ({})'.format(detail)
    return line


def _esp_console_text(total_inwc, sides_text):
    """The highlighted external static total in the console."""
    return 'EXTERNAL STATIC PRESSURE = {:.3f} in. wc ({})'.format(
        float(total_inwc), sides_text)


def _emit_html(html, fallback):
    """Print one themed console section, or its plain text when print_html
    is missing or throws. The failure is logged, never silent."""
    try:
        output.print_html(html)
        return True
    except Exception as ex:
        script.get_logger().warning(
            'Duct Performance: output.print_html failed ({}: {}), plain text '
            'printed instead.'.format(type(ex).__name__, ex))
    try:
        fallback()
    except Exception as ex:
        script.get_logger().error(
            'Duct Performance: plain text fallback failed too ({}: {}).'.format(
                type(ex).__name__, ex))
    return False


def _root_class_airflow(root_id, all_children, all_terminals):
    """Total CFM and terminal count per system class, downstream of ONE root.

    Per equipment, deliberately: the external static report is grouped by unit,
    and an airflow total that silently spanned two RTUs would not match the
    static pressure printed beside it.

    Counted per REACHABLE TERMINAL, the same source the duct CFMs are built
    from (hvac_graph reads Flow only at OST_DuctTerminal leaves), so this total
    and the duct velocities cannot disagree. A terminal reachable from two roots
    is counted under both, which is correct: both fans move that air.

    Returns {sys_class: (total_cfm, terminal_count)}.
    """
    out = {}
    stack = [root_id]
    seen = set([root_id])
    while stack:
        nid = stack.pop()
        term = all_terminals.get(nid)
        if term is not None:
            cfm, sys_class, _family = term
            key = sys_class or 'Unknown'
            tot, cnt = out.get(key, (0.0, 0))
            out[key] = (tot + cfm, cnt + 1)
        for cid in all_children.get(nid, []):
            if cid not in seen:
                seen.add(cid)
                stack.append(cid)
    return out


def _critical_path_loss(all_root_ids, all_children, all_duct_results,
                        all_terminals, all_nodes, safety_pct=0.0,
                        c_values=None, comp_values=None,
                        duct_rooted_ids=None, return_filter_inwc=0.0,
                        vav_box_ids=None, vav_box_drop_inwc=None):
    """Worst fan-to-terminal path per system class: the index run, with fittings.

    Total external static pressure is a PATH, not a sum. Air leaving the fan
    takes one route to one terminal and the fan has to beat the worst route;
    shorter runs get dampered back. Colin, 2026-09-24: "we do not need to
    account for every branch just the most restrictive run."

    Follows RJA's SP_LOSS_WORKSHEET method:
      duct friction   Darcy-Weisbach + Altshul-Tsal, per segment
      fittings        C * Pv, C from fitting_tables (1985 ASHRAE numbers)
      safety factor   a flat % on the subtotal, the sheet's own last row

    TWO coefficients apply at a take-off and they are never both charged to the
    same air, because they are two different streams through one junction:
      - the run TURNS OFF through a tap        -> 0.98  (supply dovetail branch)
      - the run CONTINUES past a tap on a main -> 0.20  (main duct @ take-off; worksheet 0.28, lowered 2026-10-09 per ASHRAE)
    Over a whole index run you therefore accumulate 0.20 for every OTHER tap
    hanging off the mains it traverses, plus one 0.98 where it finally leaves.
    Colin (2026-09, when it was .28): "for each branch you should add the .28 as it restricts the main duct
    which may be the most restrictive."

    Revit models a tap as a 2-connector in-line stub on an UNBROKEN main, so
    there is no element on the main representing that 0.20 - which is exactly
    why that column sits at 0 in the filled worksheet. The taps hanging off each
    main are counted instead (hvac_graph.takeoff_child_ids).

    Components on the run are included too: a balancing damper's drop for each
    one found, and the diffuser's at the end. Those two are COUNTED from the
    model but VALUED in the dialog, because a damper's drop depends on how far it
    is throttled and a diffuser's is a cutsheet figure - neither is derivable
    from duct geometry.

    NOT INCLUDED, and by definition rather than as a gap: filter, coil and
    cabinet losses. Those are inside the unit and already deducted from the
    manufacturer's published ESP, so counting them here would double them. The one
    exception is the return filter, added as a known drop when the engineer ticks
    "Include return filter" (return_filter_inwc), for a filter the published ESP
    does not already cover.

    duct_rooted_ids: roots that are a picked duct because no unit was found. For
    those, the duct and fittings between the picked duct and the unit connection
    are priced and added to every path (see _unit_side_chain).
    Fire, smoke and backdraft dampers are not priced either - they have their own
    drops and must not inherit the balancing-damper figure - and any accessory on
    the run that is not a balancing damper is reported as uncounted.

    all_children / all_terminals / all_nodes are keyed by int element id.
    all_duct_results is keyed by the real Revit ElementId, so it is re-keyed by
    its own DuctResult.element_id here rather than assumed to match.

    Returns {root_id: {sys_class: {...}}}, i.e. keyed by EQUIPMENT FIRST.
    Each leaf carries friction_inwc, fitting_inwc, component_inwc,
    subtotal_inwc, total_inwc, duct_count, fitting_count, tap_bypass_count,
    length_ft, terminal_id, terminal_name, unpriced, truncated.

    Per equipment, NOT merged across equipment. Colin, 2026-09-24: "if i
    selected two equipments i need to total external pressure drops ... split per
    equipment." Every unit gets its own fan, so its own index run and its own
    external static; merging two AHUs would let the bigger one's run hide the
    smaller one's entirely.

    Mid tree equipment (VAV plan, Colin 2026-10-22): equipment the run passes
    THROUGH (anything but this walk's own root, so a box that is also a picked
    root is still priced in its AHU's walk) is a component. One in vav_box_ids
    is a VAV box, priced at vav_box_drop_inwc (default: the balancing damper
    value) under component key 'vav_box', and the velocity the next fitting
    sees is reset to the box's outlet duct. Any other mid tree equipment (VAV
    off, or a duct heater and the like) is charged under 'mid_equip' at the
    balancing damper value. Through any mid tree equipment the walk only
    continues on its own side (supply stays supply).
    Each leaf also carries 'boxes': {box id: (MP loss to the box inlet,
    same with the safety factor)}, the worst path to each box of that root.
    """
    dr_by_int_id = {}
    for dr in all_duct_results.values():
        dr_by_int_id[dr.element_id] = dr

    box_ids  = vav_box_ids or set()
    box_drop = (vav_box_drop_inwc if vav_box_drop_inwc is not None
                else fitting_tables.component_of('damper', comp_values))

    def _comp(key):
        """One component's drop: the VAV box value, mid tree equipment at the
        balancing damper value, else the component table."""
        if key == 'vav_box':
            return box_drop
        if key == 'mid_equip':
            return fitting_tables.component_of('damper', comp_values)
        return fitting_tables.component_of(key, comp_values)

    def _kind(n):
        """Node kind for vav_checks' graph helpers."""
        if n in dr_by_int_id:
            return 'duct'
        if n in all_terminals:
            return 'terminal'
        e = all_nodes.get(n)
        if e is None:
            return None
        if hvac_graph.is_equipment(e):
            return 'equipment'
        if hvac_graph.is_fitting(e):
            return 'fitting'
        if hvac_graph.is_accessory(e):
            return 'accessory'
        return None

    first_duct_cache = {}
    def _first_duct(start):
        """DuctResult at start, or the nearest one below it through fittings
        and accessories (vav_checks.nearest_ducts_down), or None."""
        if start not in first_duct_cache:
            if start in dr_by_int_id:
                found = dr_by_int_id[start]
            else:
                below = vav_checks.nearest_ducts_down(start, all_children, _kind)
                found = dr_by_int_id.get(below[0]) if below else None
            first_duct_cache[start] = found
        return first_duct_cache[start]

    def _side_below(start):
        """System class of the branch starting at start, or None."""
        if start in all_terminals:
            return all_terminals[start][1]
        d = _first_duct(start)
        return d.sys_class if d is not None else None

    # Cache per-duct take-off children and end-of-main, so a duct reached by
    # several candidate paths is not re-walked for each one.
    tap_cache = {}
    def _taps(nid):
        if nid not in tap_cache:
            tap_cache[nid] = hvac_graph.takeoff_child_ids(nid, all_nodes, all_children)
        return tap_cache[nid]

    cont_cache = {}
    def _continues(nid):
        if nid not in cont_cache:
            cont_cache[nid] = hvac_graph.duct_continues_past(nid, all_nodes, all_children)
        return cont_cache[nid]

    # ── ASHRAE 2001 lookups (Colin, 2026-10-23) ─────────────────────────────
    # Every elbow, transition and tap is priced from the 2001 Fundamentals
    # table for its own geometry and flows (ashrae_fittings). When the geometry
    # cannot be read, or no 2001 table fits, the RJA worksheet C is used and a
    # FALLBACK note is reported, so the total is never silently short.
    parent_of = {}
    for _p, _cs in all_children.items():
        for _c in _cs:
            parent_of.setdefault(_c, _p)

    def _is_round_duct(dr):
        return 'dia' in (getattr(dr, 'size', '') or '')

    def _round_d_in(dr):
        """Diameter of a round main in inches, or None."""
        if not _is_round_duct(dr) or not dr.area_ft2:
            return None
        return math.sqrt(4.0 * dr.area_ft2 / math.pi) * 12.0

    def _is_45_tap(fam):
        f = (fam or '').lower()
        return 'shoe' in f or 'bevel' in f or '45' in f

    def _tap_branch_area(tap_id):
        """Area ft2 of a tap's branch opening: the smaller of its connectors
        (a beveled tap widens toward the main), or None."""
        elem = all_nodes.get(tap_id)
        if elem is None:
            return None
        try:
            areas = [a for a, _r in hvac_graph.connector_areas_ft2(elem) if a > 0]
        except Exception:
            areas = []
        return min(areas) if areas else None

    def _branch_flow(tap_id):
        """(cfm, area ft2) of the branch a tap feeds, or (None, None).

        The cfm is every duct the branch splits into below the tap (a tap
        feeding a tee counts both legs), else the terminals right on it. The
        area is the tap's own branch opening, else the one branch duct's."""
        below = vav_checks.nearest_ducts_down(tap_id, all_children, _kind)
        ducts = [dr_by_int_id[d] for d in below if d in dr_by_int_id]
        q = sum([d.cfm for d in ducts], 0.0)
        if q <= 0.0:
            q = sum([all_terminals[ch][0] for ch in all_children.get(tap_id, [])
                     if ch in all_terminals], 0.0)
        if q <= 0.0:
            return None, None
        ab = _tap_branch_area(tap_id)
        if ab is None and len(ducts) == 1:
            ab = ducts[0].area_ft2
        return q, ab

    tapflow_cache = {}
    def _tap_flows(duct_id):
        """({tap id: (local common cfm, branch cfm, branch area, feet)}, ordered).

        On an unbroken Revit main the duct's cfm is the flow at its unit end;
        each tap sees that less the branches taken off ahead of it, so taps are
        put in order along the duct (hvac_graph.tap_positions_ft). When the
        order cannot be read every tap sees the duct's own cfm (NOTE)."""
        if duct_id in tapflow_cache:
            return tapflow_cache[duct_id]
        dr = dr_by_int_id[duct_id]
        taps = _taps(duct_id)
        pos = {}
        elem = all_nodes.get(duct_id)
        par = parent_of.get(duct_id)
        if elem is not None and par is not None and taps:
            try:
                pos = hvac_graph.tap_positions_ft(elem, par)
            except Exception:
                pos = {}
        ordered = bool(taps) and all(t in pos for t in taps)
        order = sorted(taps, key=lambda t: pos[t]) if ordered else list(taps)
        q = dr.cfm
        out = {}
        for t in order:
            qb, ab = _branch_flow(t)
            out[t] = (q, qb, ab, pos.get(t))
            if ordered and qb:
                q = max(q - qb, 0.0)
        tapflow_cache[duct_id] = (out, ordered)
        return tapflow_cache[duct_id]

    def _floor_return(loss, sys_class, what):
        """(loss, notes). ASHRAE prints negative converging junction C for
        the slower stream, a credit. On return and exhaust a tap loss is
        floored at zero for external static, a conservative estimate (Colin,
        2026-10-23); the report names each one floored."""
        if loss < 0.0 and not fitting_tables._is_supply(sys_class):
            return 0.0, ['%s %.3f in. wc from the 2001 table, floored at 0'
                         % (what, loss)]
        return loss, []

    tapprice_cache = {}
    def _tap_turn(tap_id):
        """(loss in. wc, label, notes) for the air turning off through tap_id,
        or None when no 2001 table fits.

        The printed Cb is referenced to the branch section (2001 Example 8,
        Eq. 35). ashrae_fittings converts it to Cc,b, referenced to the main,
        at the Ab/Ac the table was read at (Eq. 34), so ER5-2 / ER5-3 keep
        their printed Ab/Ac 0.5 geometry; the loss is Cc,b x the main's Pv
        (Colin, 2026-10-23: mains are designed at 600 fpm return and 800 fpm
        supply, low pressure). Return and exhaust losses floor at zero."""
        if tap_id in tapprice_cache:
            return tapprice_cache[tap_id]
        res = None
        main = parent_of.get(tap_id)
        dr = dr_by_int_id.get(main)
        elem = all_nodes.get(tap_id)
        if dr is not None and elem is not None and dr.area_ft2:
            flows, _ord = _tap_flows(main)
            qc, qb, ab, _ft = flows.get(tap_id, (dr.cfm, None, None, None))
            if qc and qb and ab:
                fam = hvac_graph.fitting_family_name(elem)
                cb, label, notes = ashrae_fittings.tap_branch_c(
                    fitting_tables._is_supply(dr.sys_class),
                    'round' if hvac_graph.fitting_is_round(elem) else 'rect',
                    'round' if _is_round_duct(dr) else 'rect',
                    _is_45_tap(fam), ab / dr.area_ft2, min(qb / qc, 1.0),
                    main_d_in=_round_d_in(dr))
                if cb is not None:
                    vc = qc / dr.area_ft2
                    loss, fl = _floor_return(
                        cb * hvac_graph.velocity_pressure_inwg(vc), dr.sys_class,
                        'tap %s branch loss' % tap_id)
                    res = (loss, '%s at main %.0f fpm' % (label, vc),
                           list(notes) + fl)
        tapprice_cache[tap_id] = res
        return res

    passprice_cache = {}
    def _tap_pass(duct_id, t):
        """(loss in. wc, table, C, notes, fallback) for the air continuing past
        tap t on main duct_id: Cc,s (converted from the printed Cs at the
        As/Ac the table was read at) x the main's Pv at the tap. On an
        unbroken main As = Ac; the Cs tables stop at As/Ac 0.9 and are read on
        that row (stated in the METHOD)."""
        key = (duct_id, t)
        if key in passprice_cache:
            return passprice_cache[key]
        dr = dr_by_int_id[duct_id]
        flows, _ord = _tap_flows(duct_id)
        qc, qb, ab, _ft = flows.get(t, (dr.cfm, None, None, None))
        elem = all_nodes.get(t)
        res = None
        if qc and qb and dr.area_ft2 and elem is not None:
            fam = hvac_graph.fitting_family_name(elem)
            qs = max(qc - qb, 0.0)
            cs, lbl, _n = ashrae_fittings.tap_main_c(
                fitting_tables._is_supply(dr.sys_class),
                'round' if hvac_graph.fitting_is_round(elem) else 'rect',
                'round' if _is_round_duct(dr) else 'rect',
                _is_45_tap(fam), 1.0, qs / qc,
                ab / dr.area_ft2 if ab else None, main_d_in=_round_d_in(dr))
            if cs is not None:
                vc = qc / dr.area_ft2
                one, fl = _floor_return(
                    cs * hvac_graph.velocity_pressure_inwg(vc), dr.sys_class,
                    'tap %s straight-through loss' % t)
                res = (one, lbl.split(' ')[0] + ' Cc,s', cs, fl, None)
        if res is None:
            c_ws = fitting_tables.takeoff_main_c(dr.sys_class, c_values)
            one = c_ws * hvac_graph.velocity_pressure_inwg(dr.fpm)
            res = (one, 'RJA worksheet', c_ws, [],
                   'FALLBACK: tap %s passed on duct %s, RJA worksheet C %.2f '
                   '(flows or branch area unreadable)' % (t, duct_id, c_ws))
        passprice_cache[key] = res
        return res

    def _taps_passed(duct_id, child_id):
        """(loss in. wc, count, [(table, C, loss)], report notes) for the taps
        on main duct_id the run passes before leaving through child_id. Only
        taps AHEAD of the exit are passed when the order is known; a run
        continuing on the main passes all."""
        taps = _taps(duct_id)
        flows, ordered = _tap_flows(duct_id)
        notes = []
        if child_id in taps:
            if ordered:
                mine = flows[child_id][3]
                passed = [t for t in taps if t != child_id and flows[t][3] < mine]
            else:
                passed = [t for t in taps if t != child_id]
                notes.append('NOTE: tap order on duct %s unreadable, every tap '
                             'priced at the full duct flow' % duct_id)
        else:
            passed = list(taps)
        loss, parts = 0.0, []
        for t in passed:
            one, table, c, n, fb = _tap_pass(duct_id, t)
            loss += one
            parts.append((table, c, one))
            notes.extend(['NOTE: ' + x for x in n])
            if fb:
                notes.append(fb)
        return loss, len(passed), parts, notes

    def _ashrae_fitting(elem, role, fam, sys_class, up_fpm, up_area):
        """(loss in. wc, label, notes, downstream area) from the 2001 tables,
        or None to fall back to the worksheet."""
        if role == 'elbow':
            g = hvac_graph.elbow_geometry(elem)
            if g is None:
                return None
            f = (fam or '').lower()
            if g['round']:
                if not g['d_in']:
                    return None
                rd = (g['r_in'] / g['d_in']) if g['r_in'] else None
                c, label, notes = ashrae_fittings.round_elbow_c(g['d_in'], rd, g['angle'])
            else:
                if not g['w_in'] or not g['h_in']:
                    return None
                hw = g['h_in'] / g['w_in']
                if 'radius' in f and 'miter' not in f:
                    if not g['r_in']:
                        return None
                    c, label, notes = ashrae_fittings.rect_smooth_elbow_c(
                        g['r_in'] / g['w_in'], hw, g['angle'])
                else:
                    c, label, notes = ashrae_fittings.rect_mitered_elbow_c(hw, g['angle'])
            if c is None:
                return None
            return (c * hvac_graph.velocity_pressure_inwg(up_fpm),
                    '%s, C %.2f' % (label, c), notes, None)
        if role == 'transition':
            ends, theta = hvac_graph.transition_geometry(elem)
            if not ends or theta is None or not up_area or not up_fpm:
                return None
            # Section 1 is the walk-upstream (unit side) end and o the far end,
            # for supply (air 1 -> o) and return (air o -> 1, 1 toward the fan)
            # alike, so the walk order IS the table order.
            e1, eo = sorted(ends, key=lambda e: abs(e[0] - up_area))
            c, label, notes = ashrae_fittings.transition_c(
                fitting_tables._is_supply(sys_class), e1[1], eo[1], e1[0], eo[0], theta)
            if c is None:
                return None
            vo = up_fpm * up_area / eo[0]
            return (c * hvac_graph.velocity_pressure_inwg(vo),
                    '%s, Co %.2f at %.0f fpm' % (label, c, vo), notes, eo[0])
        return None

    _COMP_LABEL = dict((k, lbl) for k, lbl, _v in fitting_tables.COMPONENT_TABLE)
    _C_IN_DETAIL = re.compile(r'\b(?:C|Cb|Co|Cs)\s+(-?\d+\.\d+)')
    _CC_IN_DETAIL = re.compile(r'Cc,[bs]\s+(-?\d+\.\d+)')
    _TABLE_ID = re.compile(r'^([A-Z]{2}\d+-\d+)')

    def _takeoff_key(elem, fam, detail, counted, note):
        """(row name, table, C, feet) grouping one fitting into the external
        static take-off table: fittings of one kind, one table and one C share
        a row, the way the WSFU take-off groups fixtures."""
        role = fitting_tables.classify_fitting(fam)
        f = (fam or '').lower()
        if role == 'elbow':
            ang = hvac_graph.fitting_angle_deg(elem)
            ang = 90.0 if ang is None else ang
            if hvac_graph.fitting_is_round(elem):
                kind = 'Round elbow'
            elif 'radius' in f and 'miter' not in f:
                kind = 'Rect radius elbow'
            else:
                kind = 'Rect mitered elbow'
            name = '%s, %.0f deg' % (kind, ang)
        elif role == 'transition':
            name = 'Transition'
        elif role == 'takeoff':
            name = 'Tap, into branch'
        elif role in ('tee', 'cross', 'wye', 'offset'):
            name = role.capitalize()
        else:
            name = fam or 'Fitting'
        m = _TABLE_ID.match(detail or '')
        if note and 'UNPRICED' in note and not note.startswith(('FALLBACK', 'NOTE')):
            table = 'UNPRICED'
        elif m:
            table = m.group(1)
        else:
            table = 'RJA worksheet'
        # A tap's take-off C is Cc,b: the value that multiplies the main Pv.
        cm = _CC_IN_DETAIL.search(detail or '') or _C_IN_DETAIL.search(detail or '')
        return (name, table, float(cm.group(1)) if cm else None, None)

    def _notes_text(notes):
        """Lookup notes (clamped read, blank cell, negative credit) as one
        NOTE line for the report, or None."""
        notes = [n for n in (notes or []) if n]
        return ('NOTE: ' + '; '.join(notes)) if notes else None

    def _price_fitting(elem, sys_class, up_fpm, up_area, against_flow=False,
                       nid=None):
        """(loss in. wc, counted 0/1, unpriced note or None, downstream area,
        detail text).

        One place that prices a fitting the path passes THROUGH, used by the main
        walk and by the unit-side chain below so the two can never disagree.
        Elbows, transitions and taps come from the ASHRAE 2001 tables first;
        the RJA worksheet prices anything else and anything unreadable (the
        note then starts FALLBACK).

        against_flow: the walk is going the opposite way to the air (the unit
        side chain of a picked supply duct). The 2001 transition tables are
        keyed to the walk order for the main walk only, so a reversed walk
        prices transitions on the worksheet, with its areas swapped so a return
        transition that grows toward the unit is priced as the contraction the
        air sees (Colin, 2026-10-21). The returned downstream area stays in
        WALK order, because the caller tracks area along the walk.
        """
        fam  = hvac_graph.fitting_family_name(elem)
        role = fitting_tables.classify_fitting(fam)
        if role == 'takeoff' and nid is not None:
            t = _tap_turn(nid)
            if t is not None:
                return t[0], 1, _notes_text(t[2]), None, t[1]
        reversed_walk = against_flow != fitting_tables.flows_toward_unit(sys_class)
        if role in ('elbow', 'transition') and not (role == 'transition' and reversed_walk):
            a = _ashrae_fitting(elem, role, fam, sys_class, up_fpm, up_area)
            if a is not None:
                return a[0], 1, _notes_text(a[2]), a[3], a[1]
        up_a, down_a = (None, None)
        flow_up, flow_down = (None, None)
        if role == 'transition':
            up_a, down_a = hvac_graph.transition_areas(elem, up_area)
            flow_up, flow_down = (down_a, up_a) if against_flow else (up_a, down_a)
        angle = hvac_graph.fitting_angle_deg(elem) if role == 'elbow' else None
        c, note = fitting_tables.fitting_c(
            role, sys_class,
            is_round=hvac_graph.fitting_is_round(elem),
            upstream_area_ft2=flow_up, downstream_area_ft2=flow_down,
            is_end_of_main=False, c=c_values, angle_deg=angle)
        fb = None
        if role in ('elbow', 'takeoff') or (role == 'transition' and not reversed_walk):
            fb = 'FALLBACK: %s: %s, RJA worksheet C %.2f (geometry or flows unreadable)' % (
                fam or '?', note, c)
        if c > 0.0:
            return (c * hvac_graph.velocity_pressure_inwg(up_fpm), 1, fb, down_a,
                    '%s, C %.2f worksheet' % (note, c))
        if 'UNPRICED' in note:
            return 0.0, 0, (fam or '?') + ': ' + note, down_a, 'UNPRICED'
        return 0.0, 0, None, down_a, note

    def _unit_side_chain(root_id):
        """Duct and fittings BEHIND a picked duct, on the unit side.

        With no unit on the ductwork the run is rooted at the picked duct, and the
        undirected walk hangs the stretch back toward the unit off the root as a
        side branch that reaches no terminal, so the path search never charged it.
        Colin, 2026-10-08: the fittings between the picked duct and the unit
        connection are ALWAYS part of the external static, so they are priced here
        and added to every path from this root.

        Only subtrees with no terminal and that do not start at a take-off count,
        so a tap hanging off the root is never mistaken for the unit side. The
        air in this stretch is the root's own flow, so velocity comes from the
        root's CFM over each duct's own area.

        Returns (friction, fitting loss, fittings priced, ducts, feet, unpriced).
        """
        root_dr = dr_by_int_id.get(root_id)
        if root_dr is None or not root_dr.fpm or root_dr.fpm <= 0:
            return (0.0, 0.0, 0, 0, 0.0, ())
        flow_cfm = root_dr.fpm * root_dr.area_ft2
        fric = fit = dlen = 0.0
        n_fit = n_duct = 0
        unpriced = ()
        for first in all_children.get(root_id, []):
            order = []
            todo = [first]
            has_term = False
            while todo:
                n = todo.pop()
                order.append(n)
                if n in all_terminals:
                    has_term = True
                    break
                todo.extend(all_children.get(n, []))
            if has_term:
                continue
            first_elem = all_nodes.get(first)
            if (first_elem is not None and hvac_graph.is_fitting(first_elem) and
                    fitting_tables.classify_fitting(
                        hvac_graph.fitting_family_name(first_elem)) == 'takeoff'):
                continue
            up_area = root_dr.area_ft2
            for n in order:
                elem = all_nodes.get(n)
                dr   = dr_by_int_id.get(n)
                if dr is not None:
                    fpm = flow_cfm / dr.area_ft2 if dr.area_ft2 else 0.0
                    d_h = getattr(dr, 'd_h_in', 0.0)
                    if d_h > 0 and fpm > 0:
                        fric += (hvac_graph.duct_friction_loss_per_100ft(
                                     fpm, d_h, getattr(dr, 'roughness_ft', None))
                                 * dr.length_ft / 100.0)
                    dlen   += dr.length_ft
                    n_duct += 1
                    up_area = dr.area_ft2
                elif elem is not None and hvac_graph.is_fitting(elem):
                    up_fpm = flow_cfm / up_area if up_area else 0.0
                    # This stretch is walked from the picked duct back TOWARD the
                    # unit, the reverse of the main walk, so supply air runs
                    # against it and return air runs with it.
                    loss, counted, note, down_a, _detail = _price_fitting(
                        elem, root_dr.sys_class, up_fpm, up_area,
                        against_flow=not fitting_tables.flows_toward_unit(
                            root_dr.sys_class))
                    fit   += loss
                    n_fit += counted
                    if note:
                        unpriced = unpriced + (note,)
                    if down_a:
                        up_area = down_a
        return (fric, fit, n_fit, n_duct, dlen, unpriced)

    best = {}              # root_id -> {sys_class: {...}}
    box_reach = {}         # root_id -> {box id: worst MP loss to the box inlet}
    truncated_roots = set()  # PER ROOT: one unit truncating must not label the
                             # others, now that results print per equipment.

    # Bound on total path steps. A simple-path search is exponential in the
    # worst case, and while duct networks are near-trees, a badly modelled one
    # with many merge points could in principle blow up. Hitting this is
    # reported, never silently swallowed.
    MAX_PATH_STEPS = 200000

    for root_id in all_root_ids:
        best_here = {}     # this equipment's own per-system-class winners
        boxes_here = {}    # box id -> worst MP loss reaching its inlet
        # (node, friction, fitting loss, ducts, fittings, taps bypassed, feet,
        #  upstream FPM, upstream area ft2, sys_class, unpriced notes,
        #  nodes already on THIS path)
        # A duct-rooted pick starts with the unit-side stretch already priced.
        pre = (0.0, 0.0, 0, 0, 0.0, ())
        if duct_rooted_ids and root_id in duct_rooted_ids:
            pre = _unit_side_chain(root_id)
        # pre = (friction, fitting loss, fittings, ducts, feet, unpriced)
        # Last field: was the previous duct on this path flex? One flex bend is
        # charged when a path ENTERS flex, so a flex run split into several
        # Revit segments still gets one bend, not one per segment.
        # Last field: the run item by item, (item, detail, fpm, loss in. wc),
        # printed under the summary so every number in the total can be traced.
        pre_items = ()
        if pre[1] or pre[0]:
            pre_items = (('Unit side stretch', '{} fitting(s), {} duct(s), {:.0f} ft'.format(
                pre[2], pre[3], pre[4]), None, pre[0] + pre[1],
                ('Unit side stretch (picked duct to unit)', 'mixed', None, pre[4])),)
        stack = [(root_id, pre[0], pre[1], (), pre[3], pre[2], 0, pre[4], 0.0,
                  0.0, None, pre[5], frozenset([root_id]), False, pre_items)]
        steps = 0
        while stack:
            steps = steps + 1
            if steps > MAX_PATH_STEPS:
                hvac_graph.log.warning('_critical_path_loss: path search truncated at %d '
                            'steps from root %s', MAX_PATH_STEPS, root_id)
                truncated_roots.add(root_id)
                break
            (nid, fric, fit, comp_keys, dcount, fcount, tcount, dlen,
             up_fpm, up_area, sys_class, unpriced, path, in_flex,
             items) = stack.pop()

            elem = all_nodes.get(nid)
            dr   = dr_by_int_id.get(nid)

            if dr is not None:
                fric      = fric + dr.friction_loss_inwc
                dlen      = dlen + dr.length_ft
                dcount    = dcount + 1
                up_fpm    = dr.fpm
                up_area   = dr.area_ft2
                sys_class = dr.sys_class
                items     = items + ((
                    'Duct {}'.format(nid),
                    '{}, {:.1f} ft'.format(getattr(dr, 'size', '?'),
                                           float(dr.length_ft)),
                    dr.fpm, dr.friction_loss_inwc,
                    ('Flex duct' if getattr(dr, 'is_flex', False) else 'Straight duct',
                     'Darcy-Weisbach', None, float(dr.length_ft))),)
                # Straight flex friction alone covers far less than one real
                # bend (ASHRAE Duct Systems Design Guide Ch. 2: flex bend C 0.9 to
                # 3.3). Colin, 2026-10-21: charge one bend per continuous flex
                # run at the duct's own velocity.
                is_flex_here = bool(getattr(dr, 'is_flex', False))
                if is_flex_here and not in_flex and dr.fpm:
                    bend   = (fitting_tables.flex_bend_c(c_values) *
                              hvac_graph.velocity_pressure_inwg(dr.fpm))
                    fit    = fit + bend
                    fcount = fcount + 1
                    items  = items + ((
                        'Flex bend', 'C {:.2f}'.format(
                            fitting_tables.flex_bend_c(c_values)),
                        dr.fpm, bend,
                        ('Flex bend, one per flex run', 'DSDG Ch. 2',
                         fitting_tables.flex_bend_c(c_values), None)),)
                in_flex = is_flex_here

            elif elem is not None and hvac_graph.is_fitting(elem):
                loss, counted, note, _down_a, detail = _price_fitting(
                    elem, sys_class, up_fpm, up_area,
                    against_flow=fitting_tables.flows_toward_unit(sys_class),
                    nid=nid)
                fit    = fit + loss
                fcount = fcount + counted
                if note:
                    unpriced = unpriced + (note,)
                _fam = hvac_graph.fitting_family_name(elem) or '?'
                items = items + ((
                    'Fitting {}'.format(nid),
                    '{}, {}'.format(_fam, detail),
                    up_fpm, loss,
                    _takeoff_key(elem, _fam, detail, counted, note)),)
                # Past a transition the air moves at its far end's velocity, so
                # the next fitting (or transition) is priced and oriented there.
                # Same airflow, new area.
                if _down_a and up_area and up_fpm:
                    up_fpm = up_fpm * up_area / _down_a
                    up_area = _down_a

            elif elem is not None and hvac_graph.is_accessory(elem):
                # Accessories are components, not fittings: the drop is a
                # cutsheet number the engineer enters and the model supplies only
                # the COUNT. Each kind gets its OWN value - a fire or smoke damper
                # must never inherit the balancing damper's, and anything
                # matching neither (a backdraft damper, say) is reported as
                # uncounted rather than priced wrong.
                fam  = hvac_graph.fitting_family_name(elem)
                akey = fitting_tables.classify_accessory(fam)
                if akey is not None:
                    comp_keys = comp_keys + (akey,)
                    items = items + ((
                        'Accessory {}'.format(nid), fam or '?', None,
                        fitting_tables.component_of(akey, comp_values),
                        (_COMP_LABEL.get(akey, akey), 'known drop', None, None)),)
                else:
                    unpriced = unpriced + (
                        (fam or '?') + ': accessory on the run, not priced',)

            box_here = False
            mid_equip = False
            # nid != root_id, not "not a picked root": a box picked alongside
            # its AHU is still priced inside the AHU's walk (review C3).
            if (dr is None and elem is not None and nid != root_id
                    and hvac_graph.is_equipment(elem)):
                mid_equip = True
                if nid in box_ids:
                    # What the MP supply costs to reach this box's inlet, before
                    # the box's own drop: the unit's max ESP less this is the
                    # pressure left at the box.
                    _reach = fric + fit + sum([_comp(_k) for _k in comp_keys], 0.0)
                    if _reach > boxes_here.get(nid, -1.0):
                        boxes_here[nid] = _reach
                    comp_keys = comp_keys + ('vav_box',)
                    items = items + ((
                        'VAV box {}'.format(nid), 'VAV box drop', None, box_drop,
                        ('VAV box', 'known drop', None, None)),)
                    box_here = True
                else:
                    # VAV off, or not a box: its own key, at the balancing
                    # damper value for now (VAV plan item 9, review S4).
                    comp_keys = comp_keys + ('mid_equip',)
                    items = items + ((
                        'Equipment {}'.format(nid),
                        'Mid tree equipment (as damper)', None, _comp('mid_equip'),
                        ('Mid tree equipment (as damper)', 'known drop', None, None)),)

            term = all_terminals.get(nid)
            if term is not None:
                # The diffuser the run ends at is itself a component drop.
                keys_here = comp_keys + ('diffuser',)
                comp_here = 0.0
                for _k in keys_here:
                    comp_here += _comp(_k)
                _cfm, term_class, _family = term
                key     = term_class or sys_class or 'Unknown'
                # Known drop added once to the RETURN path only (the "Include
                # Return Filter" option). Same on every return path, so it never
                # changes which run is the worst, only the total.
                filt    = return_filter_inwc if key == 'Return Air' else 0.0
                current = best_here.get(key)
                if current is None or (fric + fit + comp_here + filt) > (
                        current['friction_inwc'] + current['fitting_inwc'] +
                        current['component_inwc'] + current['filter_inwc']):
                    best_here[key] = {
                        'friction_inwc':    fric,
                        'fitting_inwc':     fit,
                        'component_inwc':   comp_here,
                        'filter_inwc':      filt,
                        'unit_side_fittings': pre[2],
                        'component_keys':   keys_here,
                        'duct_count':       dcount,
                        'fitting_count':    fcount,
                        'tap_bypass_count': tcount,
                        'length_ft':        dlen,
                        'terminal_id':      nid,
                        'terminal_name':    _family or '-',
                        'unpriced':         unpriced,
                        'items':            items + ((
                            'Terminal {}'.format(nid), _family or '-', None,
                            fitting_tables.component_of('diffuser', comp_values),
                            ('Diffuser / grille', 'known drop', None, None)),),
                    }

            for child_id in all_children.get(nid, []):
                # Per-PATH loop guard, not a global visited set. A single shared
                # visited set was a real bug: it marked a node seen on whichever
                # path reached it first, so where two routes merge (two AHUs
                # feeding a common duct, a ring main, or a mis-modelled shared
                # segment) the second route was thrown away at the merge and its
                # losses were never compared. Since this function exists to find
                # the WORST path, that silently under-reported external static.
                # Verified: a diamond graph reported 0.07 in. wc when the real
                # critical path was 0.52.
                if child_id in path:
                    continue
                # Through mid tree equipment the walk only goes on along its
                # own side: MP supply into a fan powered box never carries on
                # out through the box's ducted return (review C2).
                if mid_equip and not vav_checks.same_side(
                        sys_class, _side_below(child_id)):
                    continue
                extra_fit = fit
                extra_t   = tcount
                extra_items = items
                extra_unpriced = unpriced
                if dr is not None:
                    # Every tap on this main that the run passes before it
                    # leaves restricts it: Cs from the tap's 2001 table at its
                    # own flow split (worksheet C where flows are unreadable).
                    # A run leaving through a tap passes only the taps AHEAD
                    # of it on the duct when their order is known.
                    taps = _taps(nid)
                    _tl, n_by, _parts, _fbs = _taps_passed(nid, child_id)
                    if n_by > 0:
                        extra_fit = extra_fit + _tl
                        extra_t = extra_t + n_by
                        extra_unpriced = extra_unpriced + tuple(_fbs)
                        for _tbl, _c, _one in _parts:
                            extra_items = extra_items + ((
                                'Tap passed', 'on duct {}, {} {:.2f}'.format(nid, _tbl, _c),
                                dr.fpm, _one,
                                ('Tap passed on main (straight through)', _tbl, _c, None)),)
                    # A tap at the terminus of a main is the worksheet's "end of
                    # main branch" case and carries a different C, so the swing
                    # is applied here where the parent duct is known. A tap
                    # priced from its 2001 table already reads the end of a
                    # main through its flow split (Qb/Qc at 1.0), so it gets
                    # no swing.
                    if (child_id in taps and not _continues(nid)
                            and _tap_turn(child_id) is None):
                        ch = all_nodes.get(child_id)
                        if ch is not None:
                            end_c = fitting_tables.takeoff_c(
                                dr.sys_class, is_end_of_main=True, c=c_values)
                            typ_c = fitting_tables.takeoff_c(
                                dr.sys_class, c=c_values)
                            _sw = ((end_c - typ_c) *
                                   hvac_graph.velocity_pressure_inwg(dr.fpm))
                            extra_fit = extra_fit + _sw
                            extra_items = extra_items + ((
                                'End of main tap', 'C {:.2f} in place of {:.2f}'.format(
                                    end_c, typ_c),
                                dr.fpm, _sw,
                                ('End of main tap adjustment', 'RJA worksheet', None, None)),)
                c_fpm, c_area = up_fpm, up_area
                if box_here:
                    # Past a box the air moves at the box OUTLET duct's
                    # velocity, not the inlet runout's, so a fitting right on
                    # the box discharge is priced at the right Pv.
                    _od = _first_duct(child_id)
                    if _od is not None:
                        c_fpm, c_area = _od.fpm, _od.area_ft2
                stack.append((child_id, fric, extra_fit, comp_keys, dcount,
                              fcount, extra_t, dlen, c_fpm, c_area,
                              sys_class, extra_unpriced,
                              path | frozenset([child_id]), in_flex,
                              extra_items))

        if best_here:
            best[root_id] = best_here
        if boxes_here:
            box_reach[root_id] = boxes_here

    result = {}
    sf = 1.0 + safety_pct / 100.0
    for root_id, per_class in best.items():
        out = {}
        boxes = dict((bid, (v, v * sf))
                     for bid, v in box_reach.get(root_id, {}).items())
        for sys_class, b in per_class.items():
            sub = (b['friction_inwc'] + b['fitting_inwc'] + b['component_inwc'] +
                   b['filter_inwc'])
            b['subtotal_inwc'] = sub
            b['total_inwc']    = sub * sf
            b['truncated']     = root_id in truncated_roots
            b['boxes']         = boxes
            out[sys_class]     = b
        result[root_id] = out
    return result


def _vav_box_report(root_ids, vav_by_root, root_labels, critical, vav_set):
    """The VAV BOXES block of the System Summary, as structured data: per unit,
    each box and (external static on) the MP loss to its inlet and the
    pressure left at it.

    Pressure left = unit max external static less the MP loss WITH the safety
    factor, the same basis as the external static total it is compared to.
    Only shown when the engineer entered the unit's max (VAV plan item 8).

    Every value is worked out once here. _vav_summary_lines() flattens the
    'lines' of each record into the drafting view text; the console renders
    the other fields as a table.

    Returns {'title', 'basis', 'minimum', 'lines' (block header lines), 'units': [{'label',
    'lines', 'boxes': [{'box', 'lines', 'loss', 'loss_sf', 'left',
    'status'}], 'mp_terminals', 'mp_terminal_text'}]}. status is None
    (external static off), 'NOT REACHED', 'NOT CHECKED', 'OK' or 'LOW'.
    """
    umax = vav_set.get('unit_max_esp')
    minimum = float(vav_set['min_box_inwc'])
    basis = 'MP supply {:.0f} FPM, {:.2f} in. wc per 100 ft'.format(
        float(vav_set['mp_fpm']), float(vav_set['mp_fric']))
    title = 'VAV BOXES ({})'.format(basis)
    report = {'title': title, 'basis': basis, 'lines': ['', title], 'units': [],
              'minimum': minimum}

    for rid in root_ids:
        vs = vav_by_root.get(rid)
        if vs is None or not vs.boxes:
            continue
        label = root_labels.get(rid, 'id {}'.format(rid))
        unit = {'label': label, 'lines': ['  === {} ==='.format(label)],
                'boxes': [], 'mp_terminals': 0, 'mp_terminal_text': None}
        report['units'].append(unit)
        boxes = {}
        if critical is not None:
            for leaf in (critical.get(rid) or {}).values():
                boxes = leaf.get('boxes') or {}
                break
        for box in vs.boxes:
            rec = {'box': box, 'lines': ['    Box {}'.format(box)],
                   'loss': None, 'loss_sf': None, 'left': None, 'status': None}
            unit['boxes'].append(rec)
            if critical is None:
                continue
            reach = boxes.get(box)
            if reach is None:
                rec['status'] = 'NOT REACHED'
                rec['lines'].append(
                    '      MP loss to inlet: not reached by the index run search')
                continue
            rec['loss'], rec['loss_sf'] = reach[0], reach[1]
            text = '      MP loss to inlet {:.3f} in. wc ({:.3f} with safety factor)'.format(
                reach[0], reach[1])
            left = vav_checks.pressure_left(umax, reach[1])
            rec['left'] = left
            if left is None:
                rec['status'] = 'NOT CHECKED'
                text += ', pressure left not checked (no unit max entered)'
            else:
                text += ', pressure left {:.3f} in. wc'.format(left)
                if vav_checks.pressure_low(left, minimum):
                    rec['status'] = 'LOW'
                    text += ', LOW: under the {:.2f} in. wc minimum'.format(minimum)
                else:
                    rec['status'] = 'OK'
            rec['lines'].append(text)
        if vs.mp_terminals:
            unit['mp_terminals'] = len(vs.mp_terminals)
            unit['mp_terminal_text'] = (
                '{} diffuser(s) on an MP main with no box upstream '
                '(flagged red on the duct feeding each).'.format(len(vs.mp_terminals)))
    return report


def _vav_summary_lines(report):
    """Flatten a _vav_box_report() into the drafting view text lines."""
    lines = list(report['lines'])
    for unit in report['units']:
        lines.extend(unit['lines'])
        for rec in unit['boxes']:
            lines.extend(rec['lines'])
        if unit['mp_terminal_text']:
            lines.append('    ' + unit['mp_terminal_text'])
    return lines


# Minimum transition clearance between a duct and a genuinely smaller real
# downstream neighbour, inches.
_CLEARANCE_MARGIN_IN = 2.0


def _fails_height_clearance(dr, downstream_height_in):
    """True if this duct fails the diffuser/duct height clearance rule.

    Only fires where the real downstream neighbour is actually SMALLER than
    this duct (a genuine size reduction) and the step down is too shallow to
    build a transition into — an equal or larger downstream connection is a
    continuous run and needs no margin at all.

    Shared by the main and the branch labelling paths deliberately: this is a
    hard installability problem, orthogonal to whichever sizing method judged
    the duct, and it overrides both of them. Keeping it in one function is
    what stops the two paths' clearance behaviour drifting apart.
    """
    my_height = hvac_graph.effective_height_in(dr.elem)
    if my_height is None or downstream_height_in is None:
        return False
    return (downstream_height_in < my_height
            and my_height < downstream_height_in + _CLEARANCE_MARGIN_IN)


def _limits_of(dr, custom_limits):
    """(max_fpm, max_friction) for one duct: the one lookup every check uses
    (vav_checks.limits_for), MP limits for a VAV medium pressure duct."""
    return vav_checks.limits_for(dr.sys_class, custom_limits,
                                 hvac_graph.FIRM_DEFAULTS,
                                 getattr(dr, 'is_mp', False))


def _duct_label(dr, custom_limits, tol_pct, downstream_height_in=None):
    """Worst of velocity and friction checks (one-sided, unchanged), then
    downgraded to PURPLE if the duct is GREEN but oversized (a smaller
    standard size would still satisfy max FPM and max friction), then
    overridden to RED regardless of the above if the duct fails the
    diffuser/duct height clearance rule (its height doesn't clear its real
    downstream neighbor's height by the required 2" margin, wherever that
    neighbor is actually smaller — a hard installability problem, checked
    independent of velocity/friction).

    MAIN DUCTS ONLY as of the branch/diffuser split, with one exception: a
    branch feeding a round sidewall diffuser also comes through here (no table
    yet). Slot branches are judged off the slot family since 2026-10-22.
    Behaviour is unchanged from before the split — see _branch_duct_label()
    for what every other branch gets instead, and _label_duct() for the
    dispatch between them.

      Green  : cfm <= max_cap
      Yellow : max_cap < cfm <= max_cap * (1 + tol_pct/100)
      Red    : cfm > max_cap * (1 + tol_pct/100), OR fails height clearance
      Purple : GREEN AND a smaller standard size exists (see _suggest_size_info)

    Returns (label, max_cap_cfm, reason) — reason is 'Velocity', 'Friction',
    'Oversized', 'Diffuser/Duct Clearance', or '' (GREEN/GRAY, nothing to flag).

    A VAV medium pressure duct (dr.is_mp) is judged against the MP limits and
    is never PURPLE: MP supply is sized for the boxes' minimum inlet pressure,
    so a low velocity there is not an oversize to report.
    """
    max_fpm, max_friction = _limits_of(dr, custom_limits)
    tol_fac = tol_pct / 100.0

    # Velocity check (CFM vs capacity at max FPM)
    if dr.cfm <= 0 or dr.area_ft2 <= 0:
        vel_label = 'GRAY'
        max_cap   = 0.0
    else:
        max_cap = max_fpm * dr.area_ft2
        red_cap = max_cap * (1.0 + tol_fac)
        if dr.cfm <= max_cap:
            vel_label = 'GREEN'
        elif dr.cfm <= red_cap:
            vel_label = 'YELLOW'
        else:
            vel_label = 'RED'

    # Friction check
    if dr.friction_per_100ft <= 0 or max_friction <= 0:
        fric_label = 'GRAY'
    else:
        red_fric = max_friction * (1.0 + tol_fac)
        if dr.friction_per_100ft <= max_friction:
            fric_label = 'GREEN'
        elif dr.friction_per_100ft <= red_fric:
            fric_label = 'YELLOW'
        else:
            fric_label = 'RED'

    fric_wins = _PRIORITY.get(fric_label, 0) > _PRIORITY.get(vel_label, 0)
    combined  = fric_label if fric_wins else vel_label
    reason    = ('Friction' if fric_wins else 'Velocity') if combined in ('YELLOW', 'RED') else ''

    # GREEN ducts get downgraded to PURPLE if a smaller standard size — both
    # width and height at least one nominal size down — would still satisfy
    # both max FPM and max friction. i.e. it's oversized.
    if combined == 'GREEN' and not getattr(dr, 'is_mp', False):
        _, sugg_area = _suggest_shrink_info(dr, custom_limits, downstream_height_in)
        if sugg_area is not None:
            combined = 'PURPLE'
            reason   = 'Oversized'

    # Diffuser/duct height clearance — hard override, independent of
    # everything above.
    if _fails_height_clearance(dr, downstream_height_in):
        combined = 'RED'
        reason   = 'Diffuser/Duct Clearance'

    return combined, max_cap, reason


# ── branch labelling (diffuser-table method) ──────────────────────────────────

_ROLE_LABELS = {
    'ceiling':  'Branch (Ceiling)',
    'sidewall': 'Branch (Sidewall)',
    'slot':     'Branch (Slot)',
    'round_sidewall': 'Branch (Round Sidewall)',
}


def _duct_role_label(is_branch, branch_kind):
    """Display string for the Duct Role column.

    The diffuser category is carried into the label rather than just
    'Main'/'Branch' because it is the single most useful thing to see when a
    branch verdict looks wrong: it says which table the row was judged
    against, or that no table could be chosen at all.
    """
    if not is_branch:
        return 'Main'
    return _ROLE_LABELS.get(branch_kind, 'Branch (Unknown)')


def _branch_duct_label(dr, terminal_elem, terminal_cfm, downstream_height_in=None):
    """Label a branch duct by diffuser-table size match instead of velocity.

    All of the actual sizing logic lives in hvac_graph.branch_diffuser_check();
    this only adds the height-clearance override that applies to every duct
    regardless of how it was judged.

    Returns (label, max_cap_cfm, reason, BranchDiffuserResult). max_cap_cfm is
    always 0.0 — a branch has no velocity capacity figure, and the callers
    that carry that slot only ever read the label out of it. Status can be
    GREEN, RED or GRAY (converted to RED by _label_duct) but never YELLOW or PURPLE: there is no tolerance band
    on a size match, and a branch is never checked for oversize (the
    auto-sizing diffuser family sets the branch size).
    """
    res = hvac_graph.branch_diffuser_check(
        terminal_elem, terminal_cfm, hvac_graph.duct_installed_size_in(dr.elem))
    label  = res.status
    reason = res.reason
    if _fails_height_clearance(dr, downstream_height_in):
        label  = 'RED'
        reason = 'Diffuser/Duct Clearance'
    return label, 0.0, reason, res


def _label_duct(dr, custom_limits, tol_pct, downstream_height_in, branch_ctx):
    """Label one duct GREEN / YELLOW / RED / PURPLE, or GRAY.

    GRAY means the duct could not be judged (no airflow reaching it, no size,
    unreadable diffuser, no table). That is nearly always a broken or
    disconnected system, so it is shown gray in the view and explained in the
    legend rather than passed or failed. Gray ducts are not in the flagged
    table.
    """
    label, cap, reason, branch_res = _label_duct_checked(
        dr, custom_limits, tol_pct, downstream_height_in, branch_ctx)
    if label == 'GRAY' and not reason:
        if dr.cfm <= 0:
            reason = 'No airflow data'
        elif dr.area_ft2 <= 0:
            reason = 'No duct size data'
        else:
            reason = 'Cannot check'
    return label, cap, reason, branch_res


def _label_duct_checked(dr, custom_limits, tol_pct, downstream_height_in, branch_ctx):
    """Send one duct to whichever check applies to it.

    branch_ctx is None for a main, or (diffuser_category, terminal_elem,
    terminal_cfm) for a branch — built once per duct in main() from the
    downstream-terminal map, so the graph is only walked once.

    Returns (label, max_cap_cfm, reason, branch_result_or_None). The fourth
    element is None for anything judged the ductulator way, and is what tells
    the output tables where the Required Size came from. It does NOT change
    whether velocity/friction print: every duct prints its real numbers.
    May return GRAY.
    """
    if branch_ctx is None:
        label, cap, reason = _duct_label(dr, custom_limits, tol_pct, downstream_height_in)
        return label, cap, reason, None

    kind, terminal_elem, terminal_cfm = branch_ctx

    if terminal_elem is not None and not hvac_graph.is_auto_sizing_diffuser(terminal_elem):
        # Not an RJA auto-sizing diffuser: nothing sized the branch to it, so
        # the branch gets the main duct check (Colin, 2026-10-22).
        label, cap, reason = _duct_label(dr, custom_limits, tol_pct, downstream_height_in)
        return label, cap, reason, None

    if kind == 'round_sidewall':
        # Round sidewall diffusers have no table yet; Colin plans to add round
        # diffuser data later (2026-10-16). Fall back to the main check.
        # Slots are judged off the auto-sized family instead, in
        # hvac_graph.branch_diffuser_check (Colin, 2026-10-22).
        label, cap, reason = _duct_label(dr, custom_limits, tol_pct, downstream_height_in)
        return label, cap, reason, None

    if kind is None:
        # Connector geometry unreadable — no table can be chosen. Gray, not a
        # pass: an unverifiable branch must never look like a clean one.
        label  = 'GRAY'
        reason = 'Cannot classify diffuser'
        if _fails_height_clearance(dr, downstream_height_in):
            label  = 'RED'
            reason = 'Diffuser/Duct Clearance'
        return label, 0.0, reason, None

    return _branch_duct_label(dr, terminal_elem, terminal_cfm, downstream_height_in)


def _elem_name(elem):
    try:
        return elem.Name
    except Exception:
        return str(eid_int(elem.Id))


def _duct_size_label(elem):
    """Return readable size: '10"' for round/spiral, '18"x12"' for rectangular."""
    try:
        d = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if d is not None and d.AsDouble() > 0:
            return '{:.0f}"'.format(d.AsDouble() * 12.0)
        w = elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        h = elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if w and h and w.AsDouble() > 0 and h.AsDouble() > 0:
            return '{:.0f}"x{:.0f}"'.format(w.AsDouble() * 12.0, h.AsDouble() * 12.0)
    except Exception:
        pass
    return '?'


# Standard round sizes, the minimum dimension and the even-inch rounding are
# shared with Size Duct (lib/duct_sizing.py), so the Required Size this tool
# suggests is the size Size Duct would write.
from duct_sizing import ROUND_SIZES as _ROUND_SIZES          # noqa: E402
from duct_sizing import MIN_DUCT_DIM as _MIN_DUCT_DIM        # noqa: E402
from duct_sizing import snap_even as _snap_even              # noqa: E402


def _suggest_shrink_info(dr, custom_limits, downstream_height_in=None):
    """Return (label, area_ft2) for a SMALLER standard duct size than what's
    installed, satisfying both velocity AND friction limits — no tolerance
    band, straight against max FPM / max friction. Returns ('-', None) if no
    strictly-smaller standard size works.

    Used for the oversized/PURPLE suggestion only. Round/spiral: smaller
    standard diameter only. Rectangular: BOTH width and height must drop by
    at least one nominal (2") step from installed — a candidate that only
    shrinks one dimension (e.g. 12x12 -> 12x10) is not a valid suggestion,
    since that's a marginal height-only trim, not a genuine downsize. Only
    a candidate like 12x12 -> 10x10, where both dimensions are a nominal
    size smaller, qualifies. All dimensions snapped to 2" intervals.

    Never suggests below the firm's 6" absolute minimum (_ROUND_SIZES/
    _MIN_DUCT_DIM). If downstream_height_in is given (the effective height of
    whatever this duct's real downstream neighbor is, from
    hvac_graph.max_downstream_height_in), also never suggests a candidate
    that would leave less than the required 2" clearance margin wherever
    downstream_height_in is actually smaller than the candidate — a
    same-size or larger downstream connection needs no transition margin.
    """
    max_fpm, max_friction = _limits_of(dr, custom_limits)
    if dr.cfm <= 0 or max_fpm <= 0:
        return '-', None

    def _clears(candidate_in):
        if downstream_height_in is None:
            return True
        if downstream_height_in >= candidate_in:
            return True  # no reduction at this candidate, no margin needed
        return candidate_in >= downstream_height_in + 2.0

    try:
        d = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if d is not None and d.AsDouble() > 0:
            installed_d = _snap_even(d.AsDouble() * 12.0)
            for std_d in _ROUND_SIZES:
                if std_d >= installed_d:
                    break
                if not _clears(std_d):
                    continue
                area_ft2 = math.pi * (std_d / 24.0) ** 2
                vel      = dr.cfm / area_ft2
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, float(std_d))
                if vel <= max_fpm and fric <= max_friction:
                    return '{}"'.format(std_d), area_ft2
            return '-', None

        w_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        h_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if w_param and h_param and w_param.AsDouble() > 0 and h_param.AsDouble() > 0:
            w_in = _snap_even(w_param.AsDouble() * 12.0)
            h_in = _snap_even(h_param.AsDouble() * 12.0)
            # Both width and height must be at least one nominal (2") step
            # below what's installed — a candidate that only shrinks one
            # dimension isn't a genuine downsize suggestion.
            best = None
            for new_w in range(_MIN_DUCT_DIM, w_in - 2 + 1, 2):
                min_h = max(_snap_even(new_w / 4.0), _MIN_DUCT_DIM)
                max_h = min(h_in - 2, new_w * 4)
                for new_h in range(min_h, max_h + 1, 2):
                    if not _clears(new_h):
                        continue
                    area_ft2 = new_w * new_h / 144.0
                    vel      = dr.cfm / area_ft2
                    d_h      = 4.0 * new_w * new_h / (2.0 * (new_w + new_h))
                    fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h)
                    if vel <= max_fpm and fric <= max_friction:
                        if best is None or area_ft2 < best[1]:
                            best = ('{}"x{}"'.format(new_w, new_h), area_ft2)
            if best is not None:
                return best
    except Exception:
        pass
    return '-', None


def _suggest_size_info(dr, custom_limits, downstream_height_in=None):
    """Return (label, area_ft2) for the smallest standard duct size satisfying
    both velocity AND friction limits — no tolerance band, straight against
    max FPM / max friction. area_ft2 is None when no standard size could be
    found (duct needs to grow past the largest round size, or past what
    width-expansion covers for rectangular).

    Used for the undersized/YELLOW/RED suggestion (grow direction) — a truly
    undersized duct can never satisfy at a smaller height than installed
    either (smaller height only makes velocity/friction worse), so searching
    the full range from the AR floor upward is safe and never suggests a
    shrink here.
    Round/spiral: iterates standard diameters smallest-first.
    Rectangular: keeps the installed width fixed, iterates height from the
    smallest AR-valid value upward. Expands width only if no height at the
    installed width satisfies both limits within AR 4:1.
    All suggested dimensions are snapped to 2" intervals (2, 4, 6, 8...) —
    odd sizes are not stocked/installed.
    Both constraints must be satisfied — takes the binding (larger) of the two requirements.

    Never suggests below the firm's 6" absolute minimum (_ROUND_SIZES/
    _MIN_DUCT_DIM), and (same as _suggest_shrink_info) never suggests a
    candidate that would leave less than the required 2" clearance margin
    wherever downstream_height_in is actually smaller than the candidate.
    """
    max_fpm, max_friction = _limits_of(dr, custom_limits)
    if dr.cfm <= 0 or max_fpm <= 0:
        return '-', None

    def _clears(candidate_in):
        if downstream_height_in is None:
            return True
        if downstream_height_in >= candidate_in:
            return True
        return candidate_in >= downstream_height_in + 2.0

    try:
        d = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if d is not None and d.AsDouble() > 0:
            # Round / spiral — first standard diameter satisfying both vel and friction
            for std_d in _ROUND_SIZES:
                if not _clears(std_d):
                    continue
                area_ft2 = math.pi * (std_d / 24.0) ** 2
                vel      = dr.cfm / area_ft2
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, float(std_d))
                if vel <= max_fpm and fric <= max_friction:
                    return '{}"'.format(std_d), area_ft2
            return '>{}"'.format(_ROUND_SIZES[-1]), None

        # Rectangular — keep width fixed, find smallest satisfying height
        w_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        h_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if w_param and h_param and w_param.AsDouble() > 0 and h_param.AsDouble() > 0:
            w_in = _snap_even(w_param.AsDouble() * 12.0)
            h_in = _snap_even(h_param.AsDouble() * 12.0)

            # AR 4:1 both ways bounds the height search at this fixed width:
            # too-flat (w > 4h) below, too-tall (h > 4w) above.
            min_h = max(_snap_even(w_in / 4.0), _MIN_DUCT_DIM)
            max_h = w_in * 4
            for new_h in range(min_h, max_h + 2, 2):
                if not _clears(new_h):
                    continue
                area_ft2 = w_in * new_h / 144.0
                vel      = dr.cfm / area_ft2
                d_h      = 4.0 * w_in * new_h / (2.0 * (w_in + new_h))
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h)
                if vel <= max_fpm and fric <= max_friction:
                    return '{}"x{}"'.format(w_in, new_h), area_ft2

            # No height at the installed width works within AR — expand width
            for new_w in range(int(w_in) + 2, int(w_in) + 60, 2):
                min_h2 = max(_snap_even(new_w / 4.0), _MIN_DUCT_DIM)
                max_h2 = new_w * 4
                for new_h in range(min_h2, max_h2 + 2, 2):
                    if not _clears(new_h):
                        continue
                    area_ft2 = new_w * new_h / 144.0
                    vel      = dr.cfm / area_ft2
                    d_h      = 4.0 * new_w * new_h / (2.0 * (new_w + new_h))
                    fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h)
                    if vel <= max_fpm and fric <= max_friction:
                        return '{}"x{}"'.format(new_w, new_h), area_ft2
    except Exception:
        pass
    return '-', None


def _suggest_size(dr, custom_limits, tol_pct, duct_label, downstream_height_in=None):
    """Formatted-string wrapper for schedule/table display. PURPLE (oversized)
    ducts get a shrink-only suggestion (both width and height at least one
    nominal size down); YELLOW/RED (undersized) ducts get the grow-capable
    suggestion (may need a wider duct)."""
    if duct_label == 'PURPLE':
        size, _ = _suggest_shrink_info(dr, custom_limits, downstream_height_in)
    else:
        size, _ = _suggest_size_info(dr, custom_limits, downstream_height_in)
    return size


def _row_cells(label, dr, reason, role, branch_res,
               custom_limits, tol_pct, downstream_height_in):
    """Every cell value for one flagged duct, keyed by _COLUMN_DEFS key.

    Built once per duct and handed to both renderers, so the drafting-view
    schedule and the console table cannot report different numbers for the
    same duct no matter which columns each is showing.

    '#' is deliberately NOT produced here. The two tables number their rows
    differently on purpose — the schedule's numbers match the keynote circles
    placed in the view (element-id order), while the console list is re-sorted
    worst-first — so each renderer supplies its own row index.

    Every duct prints its REAL velocity and friction, mains and branches
    alike. Nothing is ever N/A, and the limits are NOT repeated per row: the
    engineer set them in the dialog, and Status + Reason already say whether
    velocity or friction is what failed.

    A branch shows its numbers WITHOUT being judged on them. A branch's verdict
    comes from the diffuser tables only (Colin, 2026-09-24: "a branch at 1100
    FPM with a correct diffuser goes to green as the diffuser is the limiting
    factor ... there should never be a time that a diffuser works but the duct
    fails greatly").
    """
    if branch_res is not None:
        required = branch_res.required_size
    else:
        required = _suggest_size(dr, custom_limits, tol_pct, label, downstream_height_in)

    fpm_cell  = '{:.0f}'.format(float(dr.fpm))
    fric_cell = '{:.3f}'.format(float(dr.friction_per_100ft))

    return {
        'status':    label,
        'reason':    reason,
        'role':      role,
        'size':      _duct_size_label(dr.elem),
        'required':  required,
        'cfm':       '{:.0f}'.format(dr.cfm),
        'fpm':       fpm_cell,
        'fric':      fric_cell,
        'length':    '{:.1f}'.format(dr.length_ft),
    }


# ── Schedule table in a Drafting View ─────────────────────────────────────────
# Drawn in the shared RJA table look (lib/rja_drafting.py, Colin 2026-10-21:
# "update the hvac duct ones to use the new styling"), the same bands, shaded
# headers and banded rows as the Water Performance and WSFU sheets.

# Color key + meaning shown in the legend, in display order.
# Wording covers both check methods, since one view mixes mains and branches.
_LEGEND_ROWS = [
    ('GREEN',  'Main: within limit. Branch: diffuser and duct both correctly sized.'),
    ('PURPLE', 'Main only: oversized, a smaller standard size is available.'),
    ('YELLOW', 'Main only: approaching the limit.'),
    ('RED',    'Main: exceeds the limit. Branch: diffuser or duct undersized. '
               'Either: fails the diffuser/duct height clearance (see Reason).'),
    ('GRAY',   'No airflow data: no CFM reaches this element (broken or '
               'disconnected system).'),
]

# Columns read best left aligned; the rest are centred.
_LEFT_COLUMNS = ('reason', 'role')


def _build_summary_view(doc, summary_lines, flagged_rows, selected_cols,
                        source_sheet_num, tn_type_id, ts, fill_id,
                        show_r_marker_note=False, legend_notes=(),
                        esp_takeoffs=()):
    """Create a Drafting View with a System Summary block + flagged-duct table
    + color legend, in the RJA table look.

    flagged_rows: list of per-duct cell dicts from _row_cells(), already in
    the order they should appear, the same order the numbered keynote circles
    were placed in the plan view, so row 3 here is circle 3 there.
    selected_cols: set of _COLUMN_DEFS keys the user kept in the dialog.
    tn_type_id: kept for the caller; the RJA table text types are used.

    Returns (ViewDrafting, total_content_height_ft, table_width_ft), or
    (None, 0.0, 0.0) on failure. The width is returned because the column
    selection changes it on every run, and the viewport is positioned by its
    centre: the caller needs it to keep the table's left edge on its anchor.
    """
    try:
        sched_view = rja_drafting._new_drafting_view(
            doc, 'Duct Schedule - DV-{} - {}'.format(source_sheet_num, ts))
        if sched_view is None:
            return None, 0.0, 0.0
        st = rja_drafting._Style(doc)

        # Columns come from the shared _COLUMN_DEFS, filtered to the user's
        # selection: same list, same order, same headers as the console table.
        cols = _selected_columns(selected_cols)
        columns = [(c[1], c[2]) for c in cols]
        # sum(..., 0.0) forces float: sum([]) returns int 0 in Python 2.7
        total_w = sum([c[2] for c in cols], 0.0)
        left = tuple(i for i, c in enumerate(cols) if c[0] in _LEFT_COLUMNS)

        y = rja_drafting._band(
            st, sched_view, 0.0, 0.0, total_w,
            'DUCT PERFORMANCE - DV-{}'.format(source_sheet_num), None,
            rja_drafting.BAND_H)
        if summary_lines:
            y = rja_drafting._note_panel(st, sched_view, 0.0, y, total_w,
                                         summary_lines)
        y -= rja_drafting.GAP

        # External static take-off, one table per unit and system class, like
        # the WSFU take-off (Colin, 2026-10-23).
        for to in esp_takeoffs:
            y = rja_drafting._band(st, sched_view, 0.0, y, total_w,
                                   'EXTERNAL STATIC TAKE-OFF - ' + to['title'], None,
                                   rja_drafting.SUB_BAND_H)
            to_cols = [(h, total_w * f) for h, f in zip(
                duct_perf_console.TAKEOFF_HEADERS, (0.34, 0.24, 0.14, 0.12, 0.16))]
            to_rows = [list(r) for r in to['rows']]
            if to.get('filter') is not None:
                to_rows.append(['Return filter', 'known drop', '1', '-',
                                '{:.3f}'.format(float(to['filter']))])
            to_rows.append(['Subtotal', '', '', '', '{:.3f}'.format(float(to['subtotal']))])
            to_rows.append(['Safety factor', to.get('safety_detail') or '', '', '',
                            '{:.3f}'.format(float(to['safety']))])
            to_rows.append(['TOTAL', '', '', '', '{:.3f}'.format(float(to['total']))])
            y = rja_drafting._styled_table(st, sched_view, 0.0, y, to_cols, to_rows,
                                           totals_row=True, left_cols=(0, 1))
            y -= rja_drafting.GAP

        y = rja_drafting._band(st, sched_view, 0.0, y, total_w,
                               'DUCTS NEEDING ATTENTION', None,
                               rja_drafting.SUB_BAND_H)
        if flagged_rows:
            # '#' is this table's own row index: it matches the numbered
            # keynote circle placed on the same duct in the plan view.
            rows = [[str(ri + 1) if c[0] == 'num' else cells.get(c[0], '')
                     for c in cols] for ri, cells in enumerate(flagged_rows)]
            y = rja_drafting._styled_table(st, sched_view, 0.0, y, columns,
                                           rows, left_cols=left)
        else:
            y = rja_drafting._note_panel(
                st, sched_view, 0.0, y, total_w,
                ['Every judged duct is within its limits.'])
        y -= rja_drafting.GAP

        # Legend: the plan colour as a swatch, then its meaning, boxed.
        y = rja_drafting._band(st, sched_view, 0.0, y, total_w,
                               'COLOR LEGEND', None, rja_drafting.SUB_BAND_H)
        legend_top = y
        row_h = rja_drafting.ROW_H
        pad = rja_drafting.PAD
        swatch = 0.014
        for color_key, meaning in _LEGEND_ROWS:
            if st.region_type_id is not None:
                rja_drafting._legend_swatch(
                    doc, sched_view, st.region_type_id, pad * 2.0,
                    y - swatch - (row_h - swatch) / 2.0, swatch,
                    _COLOR_MAP[color_key], fill_id)
            rja_drafting._t(st, sched_view, swatch + pad * 3.0, y,
                            '{}: {}'.format(color_key.title(), meaning))
            y -= row_h
        if show_r_marker_note:
            rja_drafting._t(st, sched_view, swatch + pad * 3.0, y,
                            '(R): Most restrictive diffuser. It ends the index '
                            'run that sets the required external static '
                            'pressure for its equipment or system.')
            y -= row_h
        # Extra lines under the swatches, e.g. the VAV limits and sources.
        for note in legend_notes:
            rja_drafting._t(st, sched_view, swatch + pad * 3.0, y, note)
            y -= row_h
        rja_drafting._outline(st, sched_view, 0.0, legend_top, total_w, y)

        return sched_view, (0.0 - y), total_w

    except Exception:
        return None, 0.0, 0.0


# ── main ───────────────────────────────────────────────────────────────────────
def _view_name(view):
    """View name via its built-in parameter (never .Name in IronPython)."""
    try:
        p = view.get_Parameter(BuiltInParameter.VIEW_NAME)
        return (p.AsString() or '') if p is not None else ''
    except Exception:
        return ''


def _level_name(level):
    try:
        p = level.get_Parameter(BuiltInParameter.DATUM_TEXT)
        return (p.AsString() or '') if p is not None else ''
    except Exception:
        return ''


def _elem_level_id(elem):
    """Level an element is drawn on: a duct's Reference Level, else LevelId."""
    try:
        lvl = elem.ReferenceLevel
        if lvl is not None:
            return lvl.Id
    except Exception:
        pass
    try:
        lid = elem.LevelId
        if lid is not None and lid != ElementId.InvalidElementId:
            return lid
    except Exception:
        pass
    return None


def _sheets_by_view():
    """int view id -> the ViewSheet it is placed on."""
    out = {}
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        for vpid in sheet.GetAllViewports():
            vp = doc.GetElement(vpid)
            if vp is not None:
                out[eid_int(vp.ViewId)] = sheet
    return out


def _target_plan_views(active_view, level_ids, sheets_by_view):
    """One floor plan per level the traversed ductwork sits on.

    The view the tool was run from always covers its own level. Every other
    level gets the floor plan on that level that is placed on a sheet, ranked:
    same view template as the active view, then any 'RJA Mechanical' template,
    then anything else on a sheet. Demolition plans and earlier Duct
    Performance outputs are never picked (Colin, 2026-10-16: a roof plan run
    stopped at the RTU boot because the main level was never drawn).

    level_ids: int level id -> ElementId.
    Returns ([(level ElementId, view, source sheet or None)], [missing level ElementIds]).
    """
    targets = []
    active_lvl = None
    try:
        active_lvl = active_view.GenLevel.Id
    except Exception:
        pass
    sheet = sheets_by_view.get(eid_int(active_view.Id))
    targets.append((active_lvl, active_view, sheet))

    candidates = {}
    for v in FilteredElementCollector(doc).OfClass(ViewPlan):
        try:
            if v.IsTemplate or v.ViewType != ViewType.FloorPlan or v.GenLevel is None:
                continue
        except Exception:
            continue
        sh = sheets_by_view.get(eid_int(v.Id))
        if sh is None:
            continue
        name = _view_name(v).upper()
        if 'DEMO' in name or name.startswith('DUCT PERFORMANCE'):
            continue
        # Dependent views and anything Revit will not duplicate would throw
        # inside the transaction and roll back the whole run.
        try:
            if v.GetPrimaryViewId() != ElementId.InvalidElementId:
                continue
            if not v.CanViewBeDuplicated(ViewDuplicateOption.Duplicate):
                continue
        except Exception:
            continue
        # Mechanical plans only: a level with nothing but a plumbing or
        # lighting plan is reported as missing, never colored on that plan.
        rank = 0
        try:
            tid = v.ViewTemplateId
            if tid == active_view.ViewTemplateId and tid != ElementId.InvalidElementId:
                rank = 2
            elif tid != ElementId.InvalidElementId and \
                    _view_name(doc.GetElement(tid)).upper().startswith('RJA MECHANICAL'):
                rank = 1
        except Exception:
            pass
        if rank == 0:
            continue
        lid = eid_int(v.GenLevel.Id)
        best = candidates.get(lid)
        if best is None or rank > best[0]:
            candidates[lid] = (rank, v, sh)

    missing = []
    for lid, lvl_eid in sorted(level_ids.items()):
        if active_lvl is not None and lid == eid_int(active_lvl):
            continue
        found = candidates.get(lid)
        if found is None:
            missing.append(lvl_eid)
        else:
            targets.append((lvl_eid, found[1], found[2]))
    return targets, missing


def _open_view_range(view, base_lvl_id, low_lvl_id, top_z):
    """Make one copied plan show a system that spans several levels.

    The view's template is removed (the view keeps its current graphics; only
    the copy is touched), the cut plane goes 1 ft above the highest element in
    the system, the top is unlimited, and bottom and view depth go to the
    lowest level the system reaches. Roofs, floors and ceilings are hidden in
    the copy so the slabs between levels do not cover the ductwork below.
    Tested live on OLMC 2026-10-16: 186 ducts from the RTU on the roof down
    to the main level diffusers in one view.
    """
    base_lvl = doc.GetElement(base_lvl_id)
    view.ViewTemplateId = ElementId.InvalidElementId
    vr = view.GetViewRange()
    vr.SetLevelId(PlanViewPlane.TopClipPlane, PlanViewRange.Unlimited)
    vr.SetLevelId(PlanViewPlane.CutPlane, base_lvl_id)
    vr.SetOffset(PlanViewPlane.CutPlane, top_z - base_lvl.Elevation + 1.0)
    vr.SetLevelId(PlanViewPlane.BottomClipPlane, low_lvl_id)
    vr.SetOffset(PlanViewPlane.BottomClipPlane, 0.0)
    vr.SetLevelId(PlanViewPlane.ViewDepthPlane, low_lvl_id)
    vr.SetOffset(PlanViewPlane.ViewDepthPlane, 0.0)
    view.SetViewRange(vr)
    for bic in (BuiltInCategory.OST_Roofs, BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_Ceilings):
        try:
            cat = Category.GetCategory(doc, bic)
            if cat is not None and view.CanCategoryBeHidden(cat.Id):
                view.SetCategoryHidden(cat.Id, True)
        except Exception as ex:
            output.print_md(':warning: Could not hide {} in the copied plan: {}'.format(bic, ex))


def _file_like_source_sheet(new_sheet, source_sheet):
    """File the new sheet beside its source sheet. See
    revit_helpers.file_like_source_sheet."""
    err = revit_helpers.file_like_source_sheet(doc, new_sheet, source_sheet)
    if err:
        output.print_md(':warning: Could not file sheet {} in the sheet browser: {}'.format(
            new_sheet.SheetNumber, err))


def main():
    output.print_md('## Duct Performance')
    output.print_md('_Tip: run HVAC Diagnose first to verify CFM values and network._')
    output.print_md('---')

    # 1. Validate active view
    active_view = doc.ActiveView
    if active_view.ViewType != ViewType.FloorPlan:
        forms.alert(
            'Open a floor plan view first, then run Duct Performance.',
            title='Wrong View Type', exitscript=True
        )

    # 2. Velocity + friction settings dialog
    dialog_result = show_velocity_settings_dialog()
    if dialog_result is None:
        output.print_md('**Cancelled.**')
        return
    custom_limits = dialog_result['limits']
    tol_pct       = dialog_result['tol_pct']
    selected_cols = dialog_result['selected_cols']
    full_diag     = dialog_result['full_diag']
    ext_static    = dialog_result['ext_static']
    safety_pct    = dialog_result['safety_pct']
    c_values      = dialog_result['c_values']
    comp_values   = dialog_result['comp_values']
    calc_basis    = dialog_result['calc_basis']
    filter_inwc   = dialog_result['filter_inwc']
    vav_on        = dialog_result['vav']
    vav_set       = dialog_result['vav_settings']

    # Every run, even at the standard values, so nothing set by an earlier run
    # in the same session carries over.
    hvac_graph.set_calc_basis(air_density=calc_basis['air_density'],
                              rigid_roughness=calc_basis['rigid_roughness'],
                              flex_roughness=calc_basis['flex_roughness'])
    if (calc_basis['rigid_roughness'] != hvac_graph.STANDARD_RIGID_ROUGHNESS_FT
            or calc_basis['flex_roughness'] != hvac_graph.STANDARD_FLEX_ROUGHNESS_FT
            or calc_basis['air_density'] != hvac_graph.STANDARD_AIR_DENSITY_LB_FT3):
        output.print_md(
            '**Calculation basis overridden (Advanced settings):** rigid duct '
            'eps = {:g} ft, flex eps = {:g} ft, air density = {:g} lb/ft3.'.format(
                calc_basis['rigid_roughness'], calc_basis['flex_roughness'],
                calc_basis['air_density']))

    if vav_on:
        output.print_md(
            '**VAV System is on:** supply mains between the unit and the VAV box '
            'inlets are checked as medium pressure at {:.0f} FPM and {:.2f} in. '
            'wc/100 ft.'.format(float(vav_set['mp_fpm']), float(vav_set['mp_fric'])))
    if full_diag:
        output.print_md('**Full System Diagnostic is on** — the table below and the '
                        'sheet schedule will list every duct, and every duct will '
                        'get a numbered circle in the plan, not just flagged ones.')

    # 3. Find AHUs in active view and let user pick systems
    equip_in_view = list(FilteredElementCollector(doc, active_view.Id)
                         .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                         .WhereElementIsNotElementType())

    name_to_elem = {}
    if equip_in_view:
        # Build display name → element map (deduplicate names)
        for eq in equip_in_view:
            try:
                name = eq.Symbol.Family.Name + ' : ' + eq.Name
            except Exception:
                name = _elem_name(eq)
            key = name
            suffix = 2
            while key in name_to_elem:
                key = '{} ({})'.format(name, suffix)
                suffix += 1
            name_to_elem[key] = eq

        choice = show_system_picker(sorted(name_to_elem.keys()))
    else:
        # Nothing in this view to list. The picker still opens, with the list
        # empty and only the duct route live, so the run is not dead-ended.
        output.print_md('No mechanical equipment found in active view.')
        choice = show_system_picker([])

    if choice is None:
        output.print_md('**Cancelled.**')
        return

    mode, names = choice
    if mode == 'equipment':
        sel_elems = [name_to_elem[n] for n in names]
    else:
        sel_elems = _pick_duct_system()
        if not sel_elems:
            output.print_md('**Cancelled.**')
            return
        _warn_if_same_side(sel_elems)
        output.print_md(
            '_Rooted from the picked duct(s): the traversal still walks back '
            'to the base equipment for each one, so airflow direction and '
            'downstream CFM sums are unchanged._')

    # 4. Traverse each system and merge results
    all_duct_results = {}   # ElementId -> DuctResult  (worst-wins on overlap)
    all_nodes        = {}   # merged for fitting adjacency
    all_children     = {}   # merged
    all_root_ids     = []   # one per successfully traversed system
    root_labels      = {}   # root int_id -> equipment label, for per-unit output
    duct_rooted_ids  = set()  # roots that are a picked duct, not equipment (no unit connected)
    ahu_labels       = []
    ahu_totals       = []   # (ahu_name, ahu_id, discharge_cfm, terminal_count, other_class_cfm)
    all_terminals    = {}   # int_id -> (cfm, sys_class, family_name)  (dedup across AHUs)
    all_zero_terms   = set()
    all_missing_flow = set()
    nodes_by_root    = {}   # root int_id -> set of node ids (parallel unit check)
    vav_by_root      = {}   # root int_id -> vav_checks.VavSystem (VAV on only)

    for sel_elem in sel_elems:
        output.print_md('Traversing **{}** (id {})...'.format(
            _elem_name(sel_elem), eid_int(sel_elem.Id)))
        net = hvac_graph.build_network(sel_elem, doc, equipment_level=True)

        if net.errors:
            for e in net.errors:
                output.print_md('- :warning: {}'.format(e))
            continue

        # Two picked ducts on the same system re-root at the same AHU. Without
        # this the unit would be listed, totalled and summed twice.
        if eid_int(net.root.Id) in all_root_ids:
            output.print_md('  already covered by **{}** - skipped.'.format(
                root_labels.get(eid_int(net.root.Id), _elem_name(net.root))))
            continue

        ahu_labels.append('{} (id {})'.format(_elem_name(net.root), eid_int(net.root.Id)))

        # No unit anywhere on the picked ductwork: build_network rooted at the
        # duct itself. Remembered so a supply duct and a return duct picked this
        # way can be added into ONE system's external static further down.
        if str(net.ahu_method).startswith('fallback'):
            duct_rooted_ids.add(eid_int(net.root.Id))

        # VAV: boxes and the MP supply. Run before the warnings print so an
        # unreadable box direction is listed.
        if vav_on:
            vs = hvac_graph.analyze_vav(net)
            vav_by_root[eid_int(net.root.Id)] = vs
            output.print_md('  VAV: {} box(es), {} MP supply element(s).'.format(
                len(vs.boxes), len(vs.mp_ids)))
            if not vs.is_vav:
                output.print_md('  :warning: VAV System is ticked but no VAV box '
                                'was found on this unit\'s supply. Judged as a '
                                'constant volume system.')
        nodes_by_root[eid_int(net.root.Id)] = set(net.nodes.keys())

        if net.warnings:
            for w in net.warnings:
                output.print_md(':warning: {}'.format(w))

        output.print_md('  {} ducts  |  {} terminals'.format(
            len(net.duct_results), len(net.terminal_cfms)))

        root_id = eid_int(net.root.Id)
        all_root_ids.append(root_id)
        root_labels[root_id] = '{} (id {})'.format(
            _elem_name(net.root), eid_int(net.root.Id))
        ahu_totals.append((
            _elem_name(net.root), root_id,
            net.equipment_discharge_cfm(root_id), len(net.terminal_cfms),
            net.equipment_other_class_cfm(root_id)))

        all_nodes.update(net.nodes)
        all_children.update(net.children)
        all_zero_terms.update(net.zero_terminals)
        all_missing_flow.update(net.missing_flow)

        for nid, cfm in net.terminal_cfms.items():
            if nid in all_terminals:
                continue
            term_elem = net.nodes.get(nid)
            if term_elem is None:
                continue
            all_terminals[nid] = (
                cfm,
                hvac_graph.terminal_sys_class(term_elem),
                hvac_graph.terminal_family_name(term_elem))

        for eid, dr in net.duct_results.items():
            if eid not in all_duct_results:
                all_duct_results[eid] = dr
            else:
                # Keep worst label if duct appears in multiple networks
                existing = all_duct_results[eid]
                if _PRIORITY.get(dr.label, 0) > _PRIORITY.get(existing.label, 0):
                    all_duct_results[eid] = dr

    if not all_duct_results:
        output.print_md('**No duct results — check errors above.**')
        return

    output.print_md('**Total: {} systems  |  {} ducts**'.format(
        len(ahu_labels), len(all_duct_results)))

    # 4-VAV. Parallel units, the MP set, and the VAV red rules.
    #
    # Parallel: two picked units sharing any ductwork, or (VAV on) equipment
    # reached through its supply OUTLET. Either way that unit's external
    # static is not one fan against one path, so it is marked unreliable.
    parallel_roots = vav_checks.overlapping_roots(nodes_by_root)
    for rid, vs in vav_by_root.items():
        if vs.reversed:
            parallel_roots.add(rid)
    for rid in sorted(parallel_roots):
        output.print_md(':warning: **Parallel units on one main** at {}. Its '
                        'external static is unreliable.'.format(
                            root_labels.get(rid, 'id {}'.format(rid))))

    mp_ids = set()
    vav_box_ids = set()
    inlet_runout_ids = set()   # MP with exactly one box downstream (review W5)
    for vs in vav_by_root.values():
        mp_ids |= vs.mp_ids
        vav_box_ids |= set(vs.boxes)
        inlet_runout_ids |= vs.inlet_runout_ids
    for dr in all_duct_results.values():
        dr.is_mp = dr.element_id in mp_ids

    # int duct id -> RED reason. First rule to claim a duct keeps it, in this
    # order: flex at a box inlet, then a diffuser on an MP main. Applied after
    # the normal label, so it overrides it (these are installation faults,
    # like the clearance rule).
    vav_red = {}
    if vav_on:
        vav_red = vav_checks.red_reasons(
            vav_by_root.values(),
            set(dr.element_id for dr in all_duct_results.values()
                if dr.is_mp and getattr(dr, 'is_flex', False)))

    # 4a. Branch vs main, against the merged graph.
    #
    # A duct is a BRANCH when exactly one terminal is reachable downstream of
    # it, however many segments and fittings that tap-off is made of; a MAIN
    # when two or more are. Branches are sized against the diffuser tables,
    # mains keep the velocity/friction check.
    #
    # Computed per root and unioned rather than merged by overwrite: a duct
    # reachable from two systems genuinely feeds every terminal both systems
    # reach, so taking one root's answer and discarding the other's could turn
    # a real main into an apparent branch.
    #
    # KNOWN LIMITATION, left as-is deliberately (Colin, 2026-09-23): a piece of
    # equipment feeding only ONE terminal — no AHU/DOAS trunk splitting to
    # multiple VAV/VRF/FCU taps — has its entire run, trunk included,
    # classified as a single branch and gets no velocity/friction check at
    # all. The normal case (an AHU/DOAS main with multiple taps downstream)
    # is unaffected. Not special-cased for now — under development.
    term_ids = set(all_terminals.keys())
    downstream_terms = {}
    for rid in all_root_ids:
        partial = hvac_graph.compute_downstream_terminal_ids(
            rid, all_nodes, all_children, term_ids)
        for nid, tset in partial.items():
            if nid in downstream_terms:
                downstream_terms[nid] = downstream_terms[nid] | tset
            else:
                downstream_terms[nid] = tset

    branch_ctx    = {}   # ElementId -> (category, terminal_elem, terminal_cfm) or None
    duct_roles    = {}   # ElementId -> Duct Role display string
    kind_cache    = {}   # terminal int_id -> category  (many ducts share one terminal)
    role_counts   = {'Main': 0, 'Branch': 0}

    for eid in all_duct_results.keys():
        if all_duct_results[eid].is_mp:
            # VAV medium pressure: always a main. With two or more boxes
            # downstream it is judged against the MP limits; with one it is
            # that box's inlet runout, sized to the box (Colin 2026-10-22).
            branch_ctx[eid] = None
            duct_roles[eid] = (vav_checks.ROLE_INLET
                               if eid_int(eid) in inlet_runout_ids
                               else vav_checks.ROLE_MP)
            role_counts['Main'] += 1
            continue
        tset = downstream_terms.get(eid_int(eid), set())
        term_elem = None
        if len(tset) == 1:
            term_elem = all_nodes.get(list(tset)[0])
        if term_elem is None:
            # 0 terminals downstream (a dead-end run), 2+ (a real main), or a
            # terminal that somehow isn't in the merged node map — all keep
            # the pre-existing velocity/friction treatment.
            branch_ctx[eid] = None
            duct_roles[eid] = _duct_role_label(False, None)
            role_counts['Main'] += 1
            continue
        term_id = eid_int(term_elem.Id)
        if term_id not in kind_cache:
            kind_cache[term_id] = hvac_graph.classify_diffuser_branch(term_elem)
        kind = kind_cache[term_id]
        branch_ctx[eid] = (kind, term_elem, all_terminals.get(term_id, (0.0, '', ''))[0])
        duct_roles[eid] = _duct_role_label(True, kind)
        role_counts['Branch'] += 1

    output.print_md('Duct roles: **{} main**, **{} branch** '
                    '(branch = exactly one terminal downstream).'.format(
                        role_counts['Main'], role_counts['Branch']))

    # 4b. System summary — total flow to each equipment, diffuser counts/types
    grand_total_cfm = sum(cfm for cfm, _, _ in all_terminals.values())
    sys_totals = {}    # sys_class -> [total_cfm, count]
    type_totals = {}   # family_name -> [count, total_cfm]
    for cfm, sys_class, family in all_terminals.values():
        s = sys_totals.setdefault(sys_class, [0.0, 0])
        s[0] += cfm
        s[1] += 1
        t = type_totals.setdefault(family, [0, 0.0])
        t[0] += 1
        t[1] += cfm

    # The summary is built ONCE into two forms from the same values:
    # summary_lines is the drafting view text (its layout shall not change),
    # rep holds the same content as records for the themed console sections
    # (lib/duct_perf_console.py). Every number below is worked out once and
    # handed to both.
    rep = {'grand': [], 'equipment': [], 'by_class': [], 'types': [],
           'warnings': []}
    esp_rec = None       # External Static Pressure records, when computed
    vav_rep = None       # VAV Boxes records, when VAV is on with boxes
    run_detail = []      # Index Run Detail records, when computed
    summary_lines = ['SYSTEM SUMMARY']
    run_detail_lines = []   # filled only when external static is computed
    esp_takeoffs = []       # drafting view take-off tables, one per unit and class
    for name, aid, discharge_cfm, term_count, other_class_cfm in ahu_totals:
        summary_lines.append('Equipment: {} (id {})  -  {:.0f} CFM supply air discharge ({} diffusers)'.format(
            name, aid, discharge_cfm, term_count))
        others = []
        for cls in sorted(other_class_cfm.keys()):
            summary_lines.append('    + {:.0f} CFM {} feeding this equipment (not counted in discharge total)'.format(
                other_class_cfm[cls], cls))
            others.append('{:,.0f} CFM {}'.format(float(other_class_cfm[cls]), cls))
        rep['equipment'].append([
            '{} (id {})'.format(name, aid), '{:,.0f}'.format(float(discharge_cfm)),
            str(term_count), ', '.join(others) if others else '-'])
    summary_lines.append('Grand Total Airflow: {:.0f} CFM  |  {} Diffusers  |  {} Ducts'.format(
        grand_total_cfm, len(all_terminals), len(all_duct_results)))
    rep['grand'] = [
        ('Grand total airflow', '{:,.0f} CFM'.format(float(grand_total_cfm))),
        ('Diffusers', str(len(all_terminals))),
        ('Ducts', str(len(all_duct_results)))]
    for sys_class in sorted(sys_totals.keys()):
        s_cfm, s_cnt = sys_totals[sys_class]
        summary_lines.append('  {}: {:.0f} CFM  ({} diffusers)'.format(sys_class, s_cfm, s_cnt))
        rep['by_class'].append([sys_class, '{:,.0f}'.format(float(s_cfm)), str(s_cnt)])
    summary_lines.append('Diffuser Types:')
    for family in sorted(type_totals.keys()):
        t_cnt, t_cfm = type_totals[family]
        summary_lines.append('  {}  x{}  ({:.0f} CFM)'.format(family, t_cnt, t_cfm))
        rep['types'].append([family, str(t_cnt), '{:,.0f}'.format(float(t_cfm))])
    if all_zero_terms:
        _msg = '{} diffuser(s) with Flow = 0 (missing CFM data)'.format(len(all_zero_terms))
        summary_lines.append('WARNING: ' + _msg)
        rep['warnings'].append(_msg)
    if all_missing_flow:
        _msg = '{} diffuser(s) missing a Flow parameter entirely'.format(len(all_missing_flow))
        summary_lines.append('WARNING: ' + _msg)
        rep['warnings'].append(_msg)

    # Terminal ids of the most-restrictive diffuser per equipment/system class
    # — the one the index run (_critical_path_loss) ends at, i.e. the one that
    # actually sizes the fan's external static. Only populated when External
    # Static is on, since that is the only place this path search runs.
    # Marked in the view with a circled "R" (see transaction below), on top
    # of the diffuser's real pass/fail color — never replacing it.
    restrictive_terminal_ids = set()

    # Where each console section's text starts in summary_lines, so the plain
    # text fallback can print just that section.
    esp_start = len(summary_lines)
    critical = None
    if ext_static:
        critical = _critical_path_loss(
            all_root_ids, all_children, all_duct_results, all_terminals,
            all_nodes, safety_pct, c_values, comp_values,
            duct_rooted_ids=duct_rooted_ids, return_filter_inwc=filter_inwc,
            vav_box_ids=vav_box_ids,
            vav_box_drop_inwc=vav_set['box_drop'] if vav_on else None)
        summary_lines.append(
            'TOTAL EXTERNAL STATIC PRESSURE (index run, per RJA SP_LOSS_WORKSHEET)')
        esp_rec = {'units': [], 'combined': [], 'warnings': [],
                   'method': None, 'not_included': None}
        esp_by_root = {}

        # Grouped by EQUIPMENT, because every unit has its own fan and so its own
        # external static. Merging them would let a big unit's index run hide a
        # small one's completely.
        one_sided = []   # (root_id, total_inwc, side names, is_supply) for roots with one side only
        for root_id in all_root_ids:
            per_class = critical.get(root_id)
            if not per_class:
                continue
            unit_label = root_labels.get(root_id, 'id {}'.format(root_id))
            summary_lines.append('')
            summary_lines.append('  === {} ==='.format(unit_label))
            unit_rec = {'label': unit_label, 'classes': [], 'untraced': [],
                        'esp': None, 'one_sided': None, 'checks': [],
                        'warnings': []}
            esp_rec['units'].append(unit_rec)
            esp_by_root[root_id] = unit_rec

            airflow = _root_class_airflow(root_id, all_children, all_terminals)

            for sys_class in sorted(per_class.keys()):
                c = per_class[sys_class]
                restrictive_terminal_ids.add(c['terminal_id'])
                summary_lines.append(
                    '    {}:'.format(sys_class))
                a_cfm, a_cnt = airflow.get(sys_class, (0.0, 0))
                summary_lines.append(
                    '      Total airflow      {:,.0f} CFM      ({} terminals)'.format(
                        a_cfm, a_cnt))
                # Name what was actually counted and at what rate, so a
                # component left at 0.00 reads as a deliberate zero rather than
                # as something the tool failed to find.
                tally = {}
                for _k in c.get('component_keys', ()):
                    tally[_k] = tally.get(_k, 0) + 1
                parts = []
                for _key, _lbl, _dflt in fitting_tables.COMPONENT_TABLE:
                    if _key in tally:
                        parts.append('{} x{} @ {:.3f}'.format(
                            _lbl, tally[_key],
                            fitting_tables.component_of(_key, comp_values)))
                if 'vav_box' in tally:
                    parts.append('VAV box x{} @ {:.3f}'.format(
                        tally['vav_box'], float(vav_set['box_drop'])))
                if 'mid_equip' in tally:
                    parts.append('Mid tree equipment (as damper) x{} @ {:.3f}'.format(
                        tally['mid_equip'],
                        fitting_tables.component_of('damper', comp_values)))
                # (item, in. wc, detail): one list feeds the text lines and
                # the console table.
                esp_rows = [
                    ('Duct friction', c['friction_inwc'], '{} ducts, {:.0f} ft'.format(
                        c['duct_count'], c['length_ft'])),
                    ('Fitting losses', c['fitting_inwc'],
                     '{} fittings, {} taps restricting the main'.format(
                         c['fitting_count'], c['tap_bypass_count'])),
                    ('Components', c['component_inwc'],
                     '; '.join(parts) if parts else 'none on this run'),
                ]
                if c.get('filter_inwc', 0.0) > 0.0:
                    esp_rows.append(('Return filter', c['filter_inwc'],
                                     'known drop, Include Return Filter'))
                for _item, _val, _det in esp_rows:
                    summary_lines.append(_esp_text_line(_item, _val, _det))
                notes = []
                if c.get('unit_side_fittings', 0) > 0:
                    _msg = ('(Includes {} fitting(s) between the picked duct and '
                            'the unit connection.)'.format(c['unit_side_fittings']))
                    summary_lines.append('      ' + _msg)
                    notes.append(_msg)
                safety_inwc = c['total_inwc'] - c['subtotal_inwc']
                safety_det = '+{:.0f}%'.format(safety_pct)
                total_det = 'worst run ends at {}'.format(c['terminal_name'])
                summary_lines.append(_esp_text_line('Subtotal', c['subtotal_inwc']))
                summary_lines.append(_esp_text_line('Safety factor', safety_inwc, safety_det))
                summary_lines.append(_esp_text_line('TOTAL', c['total_inwc'], total_det))
                cls_warnings = []
                for u in c['unpriced']:
                    # A FALLBACK note is priced (RJA worksheet C), not missing.
                    _u = (u if (u.startswith('FALLBACK') or u.startswith('NOTE'))
                          else 'UNPRICED: ' + u)
                    summary_lines.append('      ' + _u)
                    cls_warnings.append(_u)
                # The take-off: every item on the index run grouped by kind,
                # table and C, the way the WSFU take-off lists fixtures (Colin,
                # 2026-10-23). Its rows add up to the subtotal less the filter.
                _to_rows, _to_loss = duct_perf_console.takeoff_rows(c.get('items', ()))
                _to_sum = sum(_to_loss, 0.0)
                _expect = c['subtotal_inwc'] - c.get('filter_inwc', 0.0)
                if abs(_to_sum - _expect) > 0.0005:
                    _msg = ('take-off rows add to {:.3f} in. wc, the run totals {:.3f}; '
                            'check the index run detail.'.format(_to_sum, _expect))
                    hvac_graph.log.warning('ESP take-off mismatch %s %s: %s',
                                           unit_label, sys_class, _msg)
                    cls_warnings.append(_msg[0].upper() + _msg[1:])
                _filt = c['filter_inwc'] if c.get('filter_inwc', 0.0) > 0.0 else None
                esp_takeoffs.append({
                    'title': '{} | {} | {}'.format(
                        unit_label, sys_class,
                        '{:,.0f} CFM'.format(float(a_cfm))),
                    'rows': _to_rows, 'filter': _filt,
                    'subtotal': c['subtotal_inwc'], 'safety': safety_inwc,
                    'safety_detail': safety_det, 'total': c['total_inwc'],
                    'total_detail': total_det})
                # Console only: the index run item by item, so each number in
                # the total can be traced to a duct, fitting or tap.
                run_title = '{} | {} | worst run ends at terminal {}'.format(
                    unit_label, sys_class, c['terminal_id'])
                run_items = []
                for _it in c.get('items', ()):
                    _lbl, _det, _fpm, _loss = _it[:4]
                    run_items.append((_lbl, '{:.0f}'.format(float(_fpm)) if _fpm else '-',
                                      float(_loss or 0.0), _det))
                run_sub_det = 'before the {:.0f}% safety factor'.format(float(safety_pct))
                run_detail.append({'title': run_title, 'items': run_items,
                                   'subtotal': float(c['subtotal_inwc']),
                                   'subtotal_detail': run_sub_det})
                run_detail_lines.append('')
                run_detail_lines.append('=== {} ==='.format(run_title))
                run_detail_lines.append('{:<18} {:>7} {:>9}  {}'.format(
                    'Item', 'FPM', 'in. wc', 'Detail'))
                for _lbl, _fpm_txt, _loss, _det in run_items:
                    run_detail_lines.append('{:<18} {:>7} {:>9.3f}  {}'.format(
                        _lbl, _fpm_txt, _loss, _det))
                run_detail_lines.append('{:<18} {:>7} {:>9.3f}  {}'.format(
                    'Subtotal', '', float(c['subtotal_inwc']), run_sub_det))
                if c.get('truncated'):
                    _msg = ('path search hit its step limit, so this '
                            'may not be the true worst run.')
                    summary_lines.append('      WARNING: ' + _msg)
                    cls_warnings.append('{}: {}{}'.format(
                        sys_class, _msg[0].upper(), _msg[1:]))
                unit_rec['classes'].append({
                    'sys_class': sys_class,
                    'airflow': '{:,.0f} CFM, {} terminals'.format(float(a_cfm), a_cnt),
                    'rows': esp_rows, 'subtotal': c['subtotal_inwc'],
                    'takeoff': _to_rows, 'filter': _filt,
                    'safety': safety_inwc, 'safety_detail': safety_det,
                    'total': c['total_inwc'], 'total_detail': total_det,
                    'notes': notes, 'warnings': cls_warnings})

            # The fan sees supply and return in series, so the two index runs add.
            # Which sides went in is NAMED rather than assumed: a plenum-return
            # job would otherwise get a one-sided number labelled as the whole.
            supply  = per_class.get('Supply Air')
            returns = [(k, per_class[k]) for k in ('Return Air', 'Exhaust Air')
                       if k in per_class]
            for sys_class in sorted(airflow.keys()):
                if sys_class in per_class:
                    continue
                a_cfm, a_cnt = airflow[sys_class]
                _msg = ('{}:  {:,.0f} CFM ({} terminals), but no index run was '
                        'traced - no static pressure for it.'.format(
                            sys_class, a_cfm, a_cnt))
                summary_lines.append('    ' + _msg)
                unit_rec['untraced'].append(_msg.replace(':  ', ': ', 1))

            unit_total = 0.0
            sides = []
            if supply is not None:
                unit_total += supply['total_inwc']
                sides.append('Supply Air')
            if returns:
                worst_k, worst_c = max(returns, key=lambda kv: kv[1]['total_inwc'])
                unit_total += worst_c['total_inwc']
                sides.append(worst_k)
            if supply is not None and returns:
                summary_lines.append(
                    '    EXTERNAL STATIC PRESSURE = {:.3f} in. wc   ({})'.format(
                        unit_total, ' + '.join(sides)))
                unit_rec['esp'] = _esp_console_text(unit_total, ' + '.join(sides))
            elif sides:
                # One side only. Held back, not warned about yet: whether it is
                # half of a system depends on what else was selected.
                one_sided.append((root_id, unit_total, sides, supply is not None))
            if (vav_on and sides and vav_set.get('unit_max_esp') is not None
                    and root_id in vav_by_root and vav_by_root[root_id].is_vav):
                _umax = float(vav_set['unit_max_esp'])
                _ok = unit_total <= _umax + 1e-9
                _msg = ('Unit max external static {:.3f} in. wc: computed {:.3f} in. wc '
                        '({}), {}.'.format(_umax, unit_total, ' + '.join(sides),
                                           'within it' if _ok else 'EXCEEDS IT'))
                summary_lines.append('    ' + _msg)
                unit_rec['checks'].append((_ok, _msg))
            if root_id in parallel_roots:
                _msg = 'Parallel units on one main. This external static is unreliable.'
                summary_lines.append('    WARNING: ' + _msg)
                unit_rec['warnings'].append(_msg)

        # One-sided roots. A picked supply duct and return duct that reach a
        # unit both re-root at it and are complete above. With NO unit
        # connected, each pick roots at its own duct, so one supply root plus one
        # return root IS one system, and its external static is the two index
        # runs added (Colin, 2026-10-08). Nothing to warn about there.
        # Only duct-rooted ones are paired, so a plenum-return unit and a
        # different unit's return can never be added together by mistake.
        sup_only = [o for o in one_sided if o[3] and o[0] in duct_rooted_ids]
        ret_only = [o for o in one_sided if not o[3] and o[0] in duct_rooted_ids]
        paired = set()
        if len(sup_only) == 1 and len(ret_only) == 1:
            s_o, r_o = sup_only[0], ret_only[0]
            paired.update([s_o[0], r_o[0]])
            _title = '{} + {}'.format(
                root_labels.get(s_o[0], 'id {}'.format(s_o[0])),
                root_labels.get(r_o[0], 'id {}'.format(r_o[0])))
            _sides = '{} + {}'.format(s_o[2][0], r_o[2][0])
            summary_lines.append('')
            summary_lines.append('  === {} ==='.format(_title))
            summary_lines.append(
                '    EXTERNAL STATIC PRESSURE = {:.3f} in. wc   ({})'.format(
                    s_o[1] + r_o[1], _sides))
            esp_rec['combined'].append({
                'title': _title,
                'text': _esp_console_text(s_o[1] + r_o[1], _sides)})
        leftover = [o for o in one_sided if o[0] not in paired]
        for _rid, _tot, _sides, _is_sup in leftover:
            summary_lines.append('')
            summary_lines.append('  === {} ==='.format(
                root_labels.get(_rid, 'id {}'.format(_rid))))
            summary_lines.append(
                '    EXTERNAL STATIC PRESSURE = {:.3f} in. wc   ({} only)'.format(
                    _tot, ' + '.join(_sides)))
            if _rid in esp_by_root:
                esp_by_root[_rid]['one_sided'] = _esp_console_text(
                    _tot, ' + '.join(_sides) + ' only')
        if len([o for o in leftover if o[0] in duct_rooted_ids]) > 1:
            _msg = ('more than one system was selected, so these one-sided '
                    'runs cannot be paired into a single external static. Pick one '
                    'supply duct and one return duct per system.')
            summary_lines.append('')
            summary_lines.append('  WARNING: ' + _msg)
            esp_rec['warnings'].append(_msg[0].upper() + _msg[1:])

        _method = ('METHOD: Darcy-Weisbach + Altshul-Tsal, eps {:g} ft rigid / '
                   '{:g} ft flex, air density {:g} lb/ft3. Fittings C x Pv, C looked '
                   'up per fitting from the 2001 ASHRAE Fundamentals Ch. 34 tables '
                   '(elbows CD3-1, CR3-1, CR3-6, CR3-9; transitions SR4/SD4/ER4/ED4 '
                   'at the section o Pv; taps SR5/SD5/ER5/ED5, printed Cb and Cs '
                   'converted to the main Pv at the geometry each table was read at, '
                   'Eq. 34, ER5-2 and ER5-3 at their printed Ab/Ac 0.5). Unbroken '
                   'main: As/Ac read on the 0.9 row. Return and exhaust tap losses '
                   'floored at 0. RJA SP_LOSS_WORKSHEET C where geometry is '
                   'unreadable (FALLBACK).'.format(
                       calc_basis['rigid_roughness'], calc_basis['flex_roughness'],
                       calc_basis['air_density']))
        if filter_inwc > 0.0:
            _not_inc = ('NOT INCLUDED: coil and cabinet, already in the published ESP. '
                        'Return filter IS included as a {:.3f} in. wc known drop. Fire and '
                        'backdraft dampers not priced.'.format(filter_inwc))
        else:
            _not_inc = ('NOT INCLUDED: filter, coil and cabinet, already in the published '
                        'ESP. Fire and backdraft dampers not priced.')
        summary_lines.append('')
        summary_lines.append('  ' + _method)
        summary_lines.append('  ' + _not_inc)
        esp_rec['method'] = _method
        esp_rec['not_included'] = _not_inc

    vav_start = len(summary_lines)
    if vav_on and any(vs.boxes for vs in vav_by_root.values()):
        vav_rep = _vav_box_report(
            all_root_ids, vav_by_root, root_labels, critical, vav_set)
        summary_lines.extend(_vav_summary_lines(vav_rep))

    # Folded once here so the drafting view keeps its fixed width text. The
    # console no longer prints this text except as the fallback below, per
    # section, from the same lines (folding is per line, so a slice folds the
    # same as the whole).
    raw_summary = summary_lines
    summary_lines = _wrap_summary_lines(summary_lines)

    # The legend names the MP limits and their sources when VAV is on.
    legend_notes = []
    if vav_on:
        legend_notes = _wrap_summary_lines(
            ['{}: {}'.format(lab, val) for lab, val in vav_checks.basis_rows(vav_set)])

    def _text_fallback(title, lines):
        def _print():
            if lines:
                output.print_md('### ' + title)
                output.print_code('\n'.join(lines))
        return _print

    _emit_html(duct_perf_console.system_summary(rep), _text_fallback(
        'System Summary', _wrap_summary_lines(raw_summary[1:esp_start])))
    if esp_rec is not None:
        _emit_html(duct_perf_console.external_static(esp_rec), _text_fallback(
            'External Static Pressure',
            _wrap_summary_lines(raw_summary[esp_start:vav_start])))
    if vav_rep is not None:
        _emit_html(duct_perf_console.vav_boxes(vav_rep), _text_fallback(
            'VAV Boxes', _wrap_summary_lines(raw_summary[vav_start:])))
    if run_detail:
        _emit_html(duct_perf_console.index_run_detail(run_detail), _text_fallback(
            'Index Run Detail', run_detail_lines))

    # 5. One plan per level the ductwork reaches, each with its source sheet.
    # The view the tool was run from covers its own level; a roof plan run
    # otherwise stopped at the RTU boot, because the main level ductwork was
    # checked but never drawn.
    sheets_by_view = _sheets_by_view()
    level_ids = {}
    for eid, dr in all_duct_results.items():
        lid = _elem_level_id(dr.elem)
        if lid is not None:
            level_ids[eid_int(lid)] = lid
    level_targets, missing_levels = _target_plan_views(active_view, level_ids, sheets_by_view)

    # The whole system goes on ONE plan (Colin, 2026-10-16): the lowest
    # level's mechanical plan, with its view range raised over every element
    # in the system when more than one level is involved, so the RTU on the
    # roof and the diffusers below are colored in the same view.
    def _lvl_elev(item):
        try:
            return doc.GetElement(item[0]).Elevation
        except Exception:
            return float('inf')
    level_targets.sort(key=_lvl_elev)
    base_lvl, base_view, base_sheet = level_targets[0]
    multi_level = len(level_targets) > 1
    targets = [(None, base_view, base_sheet)]

    system_top_z = None
    for _nid, _el in all_nodes.items():
        try:
            bb = _el.get_BoundingBox(None)
        except Exception:
            bb = None
        if bb is not None and (system_top_z is None or bb.Max.Z > system_top_z):
            system_top_z = bb.Max.Z

    # Lowest level any duct sits on: the view range bottom goes there, so a
    # level with no mechanical plan of its own is still drawn in this one.
    system_low_lvl = None
    for _lid in level_ids.values():
        lv = doc.GetElement(_lid)
        if lv is not None and (system_low_lvl is None or
                               lv.Elevation < doc.GetElement(system_low_lvl).Elevation):
            system_low_lvl = _lid
    if system_low_lvl is not None and base_lvl is not None and \
            doc.GetElement(system_low_lvl).Elevation < doc.GetElement(base_lvl).Elevation:
        multi_level = True

    output.print_md('Plan: **{}** (source sheet **{}**){}'.format(
        _view_name(base_view),
        base_sheet.SheetNumber if base_sheet is not None else 'none',
        ', view range opened to show every level of the system' if multi_level else ''))

    # 6. Title block and solid fill
    tb_list = list(FilteredElementCollector(doc)
                   .OfCategory(BuiltInCategory.OST_TitleBlocks)
                   .WhereElementIsElementType())
    tb_id   = tb_list[0].Id if len(tb_list) > 0 else ElementId.InvalidElementId
    fill_id = hvac_graph.solid_fill_pattern_id(doc)

    # 7. Marker offset (5/64" printed, x2) is worked out per copied plan from
    # that plan's own scale, inside the transaction below.

    # 8. Transaction: copy view → color overrides → FPM annotations → sheet
    t = Transaction(doc, 'Duct Performance')
    t.Start()
    ts = datetime.datetime.now().strftime('%Y-%m-%d-%H%M%S')

    try:
        # ElementId -> color label for every duct, fitting, accessory and
        # diffuser. Worked out once, then applied to every copied plan below.
        elem_labels = {}

        # Color overrides — worst of velocity check and friction check
        counts         = {'GREEN': 0, 'YELLOW': 0, 'RED': 0, 'PURPLE': 0, 'GRAY': 0}
        clearance_count = 0
        # eid -> (label, green_cap_cfm, reason) for fittings + annotations + schedule
        duct_labels  = {}
        # eid -> BranchDiffuserResult, or None for anything judged the
        # ductulator way. Only decides where Required Size comes from; every
        # duct prints its real velocity and friction either way.
        branch_results = {}
        # eid -> effective height (in) of this duct's real downstream neighbor,
        # precomputed once here so _duct_label/_suggest_size don't need graph access
        downstream_heights = {}

        for eid, dr in all_duct_results.items():
            downstream_h = hvac_graph.max_downstream_height_in(eid_int(eid), all_nodes, all_children)
            downstream_heights[eid] = downstream_h
            if dr.element_id in inlet_runout_ids:
                # A VAV inlet runout is sized to its box by the engineer: not
                # judged on the MP limits, never purple. Only the clearance
                # rule and the VAV red rules below can override it.
                label, green_cap, reason, branch_res = (
                    'GREEN', 0.0, vav_checks.REASON_INLET_OK, None)
                if _fails_height_clearance(dr, downstream_h):
                    label, reason = 'RED', 'Diffuser/Duct Clearance'
            else:
                label, green_cap, reason, branch_res = _label_duct(
                    dr, custom_limits, tol_pct, downstream_h, branch_ctx.get(eid))
            if dr.element_id in vav_red:
                label, reason = 'RED', vav_red[dr.element_id]
            duct_labels[eid]    = (label, green_cap, reason)
            branch_results[eid] = branch_res
            if reason == 'Diffuser/Duct Clearance':
                clearance_count += 1
            counts[label] = counts.get(label, 0) + 1
            elem_labels[eid] = label

        # Color fittings, accessories and diffusers by worst nearby duct color
        adj = {}
        for pid, cids in all_children.items():
            if pid not in adj:
                adj[pid] = []
            for cid in cids:
                adj[pid].append(cid)
                if cid not in adj:
                    adj[cid] = []
                adj[cid].append(pid)

        fitting_counts = {'GREEN': 0, 'YELLOW': 0, 'RED': 0, 'PURPLE': 0, 'GRAY': 0}

        for nid, elem in all_nodes.items():
            # Equipment too (Colin, 2026-10-16: color every connected piece,
            # duct, fitting, damper, diffuser, unit).
            if not (hvac_graph.is_fitting_or_accessory(elem)
                    or hvac_graph.is_terminal(elem)
                    or hvac_graph.is_equipment(elem)):
                continue
            # Worst color among the nearest ducts, walking through any
            # fittings/accessories in between (a takeoff next to an elbow has
            # no duct as a direct neighbour). A diffuser therefore takes the
            # color of the branch duct feeding it. One with no duct reachable
            # at all is a piece of system that isn't properly connected, so it
            # is GRAY.
            worst   = None
            seen    = set([nid])
            frontier = [nid]
            while frontier:
                nxt = []
                for cur in frontier:
                    for neighbor_id in adj.get(cur, []):
                        if neighbor_id in seen:
                            continue
                        seen.add(neighbor_id)
                        nb_elem = all_nodes.get(neighbor_id)
                        if nb_elem is None:
                            continue
                        if hvac_graph.is_duct(nb_elem):
                            nb_label = duct_labels.get(nb_elem.Id, ('GRAY', 0.0, ''))[0]
                            if worst is None or _PRIORITY.get(nb_label, 0) > _PRIORITY.get(worst, 0):
                                worst = nb_label
                        elif hvac_graph.is_fitting_or_accessory(nb_elem):
                            nxt.append(neighbor_id)
                if worst is not None:
                    break
                frontier = nxt
            if worst is None:
                worst = 'GRAY'
            elem_labels[elem.Id] = worst
            fitting_counts[worst] = fitting_counts.get(worst, 0) + 1

        # Numbered callout markers on yellow/red ducts (placed in view)
        tn_types = list(FilteredElementCollector(doc).OfClass(TextNoteType).ToElements())
        tn_type_id = tn_types[0].Id if tn_types else None

        # Collect flagged ducts in stable element-id order. Cell values are
        # built once here and keyed by element id so the schedule view and the
        # console table render the exact same content in a different order,
        # rather than each computing its own.
        flagged_items    = []
        row_cells_by_eid = {}
        for eid in sorted(duct_labels.keys(), key=lambda e: eid_int(e)):
            lbl, _, reason = duct_labels[eid]
            if not full_diag and lbl not in ('YELLOW', 'RED', 'PURPLE'):
                continue
            dr = all_duct_results.get(eid)
            if dr is None:
                continue
            cells = _row_cells(lbl, dr, reason, duct_roles.get(eid, 'Main'),
                               branch_results.get(eid), custom_limits, tol_pct,
                               downstream_heights.get(eid))
            if dr.element_id in inlet_runout_ids:
                cells['required'] = '-'   # sized to the box, no MP suggestion
            row_cells_by_eid[eid] = cells
            flagged_items.append((lbl, dr, cells))

        # Find keynote circle symbol — search by family name
        keynote_sym = None
        for fs in FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements():
            try:
                fn = fs.Family.Name
                if 'RJA - Keynote Symbol' in fn and 'Circle' in fn:
                    keynote_sym = fs
                    break
                if 'RJA - Keynote Symbol' in fn and 'Circle' in fs.Name:
                    keynote_sym = fs
                    break
            except Exception:
                pass

        if keynote_sym is not None and not keynote_sym.IsActive:
            keynote_sym.Activate()
            doc.Regenerate()

        ogs_by_label = {}
        for lbl_key, color in _COLOR_MAP.items():
            ogs = OverrideGraphicSettings()
            ogs.SetSurfaceForegroundPatternColor(color)
            if fill_id != ElementId.InvalidElementId:
                ogs.SetSurfaceForegroundPatternId(fill_id)
            ogs.SetProjectionLineColor(color)
            ogs_by_label[lbl_key] = ogs

        # A marker is only placed in the plan for the level its element is
        # on, so a main level duct never gets a circle drawn on the roof plan.
        def _on_level(elem, lvl_eid):
            if lvl_eid is None:
                return True
            lid = _elem_level_id(elem)
            return lid is None or lid == lvl_eid

        new_sheets = []
        used_keys  = set()
        for lvl_eid, src_view, src_sheet in targets:
            source_sheet_num = src_sheet.SheetNumber if src_sheet is not None else 'NoSheet'
            # Two plans off the same sheet would give a duplicate sheet number
            # and roll the whole run back, so the second one gets its level.
            if source_sheet_num in used_keys and lvl_eid is not None:
                source_sheet_num = '{}-{}'.format(
                    source_sheet_num, _level_name(doc.GetElement(lvl_eid)))
            n = 2
            base_key = source_sheet_num
            while source_sheet_num in used_keys:
                source_sheet_num = '{}-{}'.format(base_key, n)
                n += 1
            used_keys.add(source_sheet_num)
            mark_off_ft = (5.0 / (64.0 * 12.0)) * float(getattr(src_view, 'Scale', 48)) * 2.0

            # Copy floor plan
            new_vid  = src_view.Duplicate(ViewDuplicateOption.Duplicate)
            new_view = doc.GetElement(new_vid)
            base_name = 'Duct Performance - {} - {}'.format(source_sheet_num, ts)
            try:
                new_view.Name = base_name
            except Exception:
                new_view.Name = base_name + ' (2)'

            if multi_level and system_top_z is not None and base_lvl is not None:
                _open_view_range(new_view, base_lvl, system_low_lvl or base_lvl,
                                 system_top_z)

            for oid, lbl_key in elem_labels.items():
                new_view.SetElementOverrides(oid, ogs_by_label[lbl_key])

            # Numbered callouts, one sequence across every level so circle 3
            # is row 3 of the schedule whichever sheet it is on.
            for idx, (lbl, dr, _cells) in enumerate(flagged_items, 1):
                if not _on_level(dr.elem, lvl_eid):
                    continue
                try:
                    mid_pt = dr.elem.Location.Curve.Evaluate(0.5, True)
                    if keynote_sym is not None:
                        inst      = doc.Create.NewFamilyInstance(mid_pt, keynote_sym, new_view)
                        num_param = inst.LookupParameter('Label')
                        if num_param and not num_param.IsReadOnly:
                            if num_param.StorageType == StorageType.String:
                                num_param.Set(str(idx))
                            elif num_param.StorageType == StorageType.Integer:
                                num_param.Set(idx)
                    elif tn_type_id is not None:
                        opts = TextNoteOptions(tn_type_id)
                        TextNote.Create(doc, new_vid, mid_pt, '({})'.format(idx), opts)
                except Exception:
                    pass

            # Most-restrictive-diffuser markers — circled "R" next to the
            # terminal each equipment/system's index run ends at (see
            # restrictive_terminal_ids above). Placed on top of the diffuser's
            # real pass/fail color, never in place of it: a restrictive diffuser
            # that also fails its own check still shows red, with an R beside it.
            # Reuses the same keynote-circle family as the numbered callouts
            # above (falls back to plain '(R)' text the same way, including
            # where the family's Label param is an integer and can't hold a
            # letter).
            for tid in sorted(restrictive_terminal_ids):
                term_elem = all_nodes.get(tid)
                if term_elem is None or not _on_level(term_elem, lvl_eid):
                    continue
                try:
                    r_pt = term_elem.Location.Point
                except Exception:
                    continue
                if r_pt is None:
                    continue
                # Offset off the diffuser symbol itself so the marker doesn't
                # sit directly on top of it.
                mark_pt = XYZ(r_pt.X + mark_off_ft, r_pt.Y + mark_off_ft, r_pt.Z)
                try:
                    placed = False
                    if keynote_sym is not None:
                        inst      = doc.Create.NewFamilyInstance(mark_pt, keynote_sym, new_view)
                        num_param = inst.LookupParameter('Label')
                        if num_param and not num_param.IsReadOnly and num_param.StorageType == StorageType.String:
                            num_param.Set('R')
                            placed = True
                        else:
                            doc.Delete(inst.Id)
                    if not placed and tn_type_id is not None:
                        opts = TextNoteOptions(tn_type_id)
                        TextNote.Create(doc, new_vid, mark_pt, '(R)', opts)
                except Exception:
                    pass

            # Output sheet
            new_sheet             = ViewSheet.Create(doc, tb_id)
            new_sheet.SheetNumber = 'DV-{}-{}'.format(source_sheet_num, ts)
            # Sheet/View names share one uniqueness pool in Revit — new_view above
            # already claimed base_name verbatim, so reusing it here always collides.
            new_sheet.Name = base_name + ' (Sheet)'
            _file_like_source_sheet(new_sheet, src_sheet)
            new_sheets.append(new_sheet)

            # Place viewport
            # Fixed center, positioned by hand in Revit on the Elevation Volleyball
            # DV-M101 sheet, then read back and hardcoded here via Revit MCP.
            #
            # Both Viewport.Create's placement point AND an immediate SetBoxCenter
            # right after Create landed (2.282, 1.55) off the intended center,
            # same exact offset both times, confirmed via MCP across three
            # separate generated sheets. Root cause: the diagram view's crop
            # region isn't settled yet at that point in the transaction, so any
            # box-center call made before a Regenerate() computes against a
            # stale/uncommitted crop extent. Regenerate() first, then
            # SetBoxCenter, to get the box center actually applied.
            diagram_vp = Viewport.Create(doc, new_sheet.Id, new_vid, XYZ(0, 0, 0))
            doc.Regenerate()
            diagram_vp.SetBoxCenter(XYZ(1.201, 1.212, 0))
            rja_drafting.set_viewport_type_none(doc, diagram_vp)

            # System Summary + flagged-duct table + legend, placed as second
            # viewport on every sheet (the table covers all levels).
            if tn_type_id is not None:
                sched_view, content_h, total_w = _build_summary_view(
                    doc, summary_lines, [item[2] for item in flagged_items],
                    selected_cols, source_sheet_num, tn_type_id, ts, fill_id,
                    show_r_marker_note=bool(restrictive_terminal_ids),
                    legend_notes=legend_notes, esp_takeoffs=esp_takeoffs)
                if sched_view is not None:
                    # X fixed by hand in Revit (see diagram note above) and read
                    # back via Revit MCP. Y keeps the original bottom-anchored,
                    # grows-upward-with-content_h behavior, just re-anchored to
                    # match the hand-placed bottom edge on that same sheet.
                    #
                    # Viewport.Create positions by CENTER, so the table's left edge
                    # is what has to be pinned — a table of a different width
                    # centered on the old point would drift sideways off the
                    # anchor. The table width is no longer fixed (the column picker
                    # changes it every run), so the left edge is stored instead and
                    # the center derived from it. The literal below is the original
                    # hand-placed center of -1.014 minus half the then-fixed table
                    # width of 1.360, i.e. the same left edge as before.
                    _SCHED_LEFT_EDGE_X = -1.014 - 1.360 / 2.0
                    sched_x         = _SCHED_LEFT_EDGE_X + total_w / 2.0
                    bottom_margin_y = 1.546
                    sched_y  = bottom_margin_y + content_h / 2.0
                    sched_vp = Viewport.Create(doc, new_sheet.Id, sched_view.Id,
                                               XYZ(sched_x, sched_y, 0))
                    rja_drafting.set_viewport_type_none(doc, sched_vp)

        t.Commit()
    except Exception as ex:
        t.RollBack()
        output.print_md('**Error — transaction rolled back:** {}: {}'.format(
            type(ex).__name__, str(ex)))
        output.print_md('```\n{}\n```'.format(traceback.format_exc()))
        return

    # 9. Console report: flagged ducts, design limits and legend, sheets.
    #
    # Flagged duct list — RED first, then YELLOW, then PURPLE.
    # RED/YELLOW sorted by velocity descending (worst overrun first);
    # PURPLE sorted by velocity ascending (worst oversizing first).
    #
    # Deliberately a different order from the schedule view's table, which
    # stays in element-id order so its row numbers line up with the numbered
    # keynote circles in the plan. The cell CONTENT is the same objects built
    # once above, so only the ordering and the row index differ.
    flagged = []
    for eid, dr in all_duct_results.items():
        label, _, reason = duct_labels.get(eid, ('GRAY', 0.0, ''))
        if not full_diag and label not in ('RED', 'YELLOW', 'PURPLE'):
            continue
        cells = row_cells_by_eid.get(eid)
        if cells is None:
            continue
        fpm = dr.cfm / dr.area_ft2 if dr.area_ft2 > 0 else 0.0
        flagged.append((label, fpm, cells))

    _FLAG_RANK = {'RED': 0, 'YELLOW': 1, 'PURPLE': 2}

    def _flag_sort_key(item):
        label, fpm = item[0], item[1]
        rank = _FLAG_RANK.get(label, 3)
        return (rank, fpm if label == 'PURPLE' else -fpm)

    if flagged:
        flagged.sort(key=_flag_sort_key)

        # Same shared column definitions and same user selection as the
        # schedule view.
        cols = _selected_columns(selected_cols)
        duct_rows = []
        for idx, (label, fpm, cells) in enumerate(flagged, 1):
            duct_rows.append([
                str(idx) if col[0] == 'num' else cells.get(col[0], '')
                for col in cols])
        duct_title = 'All Ducts' if full_diag else 'Flagged Ducts'
        keys = [col[0] for col in cols]

        def _ducts_text():
            # Padded monospace text, console widths from _COLUMN_DEFS.
            def _fmt_row(cells):
                return '  '.join(str(c).ljust(col[3]) for c, col in zip(cells, cols))
            rows = [_fmt_row([col[1] for col in cols]),
                    _fmt_row(['-' * col[3] for col in cols])]
            rows.extend(_fmt_row(r) for r in duct_rows)
            output.print_md('### ' + duct_title)
            output.print_code('\n'.join(rows))

        _emit_html(duct_perf_console.duct_table(
            duct_title, [col[1] for col in cols], duct_rows,
            status_col=keys.index('status') if 'status' in keys else None,
            align=['r' if k in ('num', 'cfm', 'fpm', 'fric', 'length') else 'l'
                   for k in keys],
            widths=[col[3] for col in cols]), _ducts_text)

    # Design limits and legend. One set of rows feeds the themed tables and
    # the markdown fallback.
    limit_rows = []
    for sys_class in ('Supply Air', 'Return Air', 'Exhaust Air', 'Outside Air', 'Transfer Air'):
        mx_fpm, mx_fric = vav_checks.limits_for(
            sys_class, custom_limits, hvac_graph.FIRM_DEFAULTS)
        limit_rows.append([sys_class, '{:.0f} FPM'.format(float(mx_fpm)),
                           '{:.3f} iwc/100'.format(float(mx_fric))])
    vav_basis = []
    if vav_on:
        mx_fpm, mx_fric = vav_checks.limits_for(
            'Supply Air', custom_limits, hvac_graph.FIRM_DEFAULTS, is_mp=True)
        limit_rows.append(['Supply Air, MP (VAV, unit to box inlets)',
                           '{:.0f} FPM'.format(float(mx_fpm)),
                           '{:.3f} iwc/100'.format(float(mx_fric))])
        vav_basis = list(vav_checks.basis_rows(vav_set))
    legend_rows = [
        ['Green', str(counts.get('GREEN', 0)), str(fitting_counts.get('GREEN', 0)),
         'Main: within limit. Branch: diffuser and duct both correctly sized'],
        ['Purple', str(counts.get('PURPLE', 0)), str(fitting_counts.get('PURPLE', 0)),
         'Main only: oversized (low velocity, review for cost)'],
        ['Yellow', str(counts.get('YELLOW', 0)), str(fitting_counts.get('YELLOW', 0)),
         'Main only: approaching limit'],
        ['Red', str(counts.get('RED', 0)), str(fitting_counts.get('RED', 0)),
         'Main: exceeds limit. Branch: diffuser or duct undersized'],
        ['Gray', str(counts.get('GRAY', 0)), str(fitting_counts.get('GRAY', 0)),
         'No airflow data: broken or disconnected system (not in the flagged table)'],
    ]
    legend_rec = {
        'roles_note': ('Mains and branches were checked by different methods.',
                       'A duct is a branch when exactly one terminal is reachable '
                       'downstream of it, a main when two or more are.'),
        'limits_caption': ('Main ducts, design limits used (green at or under max, '
                           'yellow within {}% above max, red over max + {}%, purple '
                           'when a smaller standard size still fits)'.format(
                               int(tol_pct), int(tol_pct))),
        'limits': limit_rows,
        'vav_basis': vav_basis,
        'notes': [
            ('Branch ducts are checked against the diffuser sizing tables '
             '(MEP Design Standards, section 1), not against velocity or friction.',
             'An auto-sizing diffuser is selected on static pressure and NC rather '
             'than duct velocity, so velocity-checking a run that feeds only that one '
             'diffuser is redundant. Two things are checked instead: whether the '
             'diffuser\'s own neck/face is big enough for its own CFM, and whether the '
             'branch duct is at least as big as the diffuser connects with. Either one '
             'failing is red; there is no yellow band on a size match. Oversize is '
             'not checked on branches.'),
            ('Branches feeding a slot diffuser',
             'are judged off the auto-sized slot family: the branch duct shall be at '
             'least the slot inlet diameter, and a filled Sizing Warning on the slot '
             'is red. No neck table is published for slots.'),
            ('Branches feeding a diffuser that is not an RJA auto-sizing family',
             '(no Sizing Warning parameter) get the main velocity/friction check.'),
        ],
        'legend': legend_rows,
        'clearance': (clearance_count > 0,
                      'Diffuser/duct height clearance issues (flagged Red): {}'.format(
                          clearance_count)),
    }

    def _legend_text():
        output.print_md('### Design Limits and Legend')
        output.print_md('**{}** {}'.format(*legend_rec['roles_note']))
        output.print_md('**{}:**'.format(legend_rec['limits_caption']))
        output.print_md(console_report.markdown_table(
            ['System', 'Max Velocity', 'Max Friction'], limit_rows))
        if vav_basis:
            output.print_md('**VAV System.** ' + ' '.join(
                '{}: {}'.format(lab, val) for lab, val in vav_basis))
        for lead, text in legend_rec['notes']:
            output.print_md('**{}** {}'.format(lead, text))
        output.print_md(console_report.markdown_table(
            ['Color', 'Ducts', 'Fittings, Accessories & Diffusers', 'Meaning'],
            legend_rows))
        output.print_md('**{}**'.format(legend_rec['clearance'][1]))

    _emit_html(duct_perf_console.limits_and_legend(legend_rec), _legend_text)

    sheet_numbers = [s.SheetNumber for s in new_sheets]

    def _sheets_text():
        output.print_md('### Sheets Created')
        for num in sheet_numbers:
            output.print_md('Sheet **{}** created.'.format(num))

    _emit_html(duct_perf_console.sheets_created(sheet_numbers), _sheets_text)

    if new_sheets:
        uidoc.ActiveView = new_sheets[0]


main()
