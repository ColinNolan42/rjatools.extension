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
          Slot diffusers are the exception: the design standard publishes no
          duct/neck size breakpoints for them, so a slot branch falls back to
          the same velocity/friction check a main gets.

The diffuser/duct height clearance check still applies to both, and still
overrides either verdict - it is an installability problem, not a sizing one.

Which columns appear in the output tables is chosen in the settings dialog;
the same selection drives both the drafting-view schedule placed on the sheet
and the console table.

Run HVAC Diagnose first to verify the network and CFM values, then run this.

IronPython 2.7 / pyRevit  --  no f-strings, no walrus, no nonlocal.
"""

import os
import sys
import math
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
    Viewport, ViewDuplicateOption, ElementId, XYZ,
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

doc    = __revit__.ActiveUIDocument.Document
uidoc  = __revit__.ActiveUIDocument
output = script.get_output()

# Add lib/ to path
_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

import hvac_graph
import fitting_tables
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
# placed on the sheet and the console print_code table. They are rendered by
# different code (detail lines + TextNotes vs. padded monospace text) and have
# drifted apart before, so the column set, the headers and the cell values all
# come from here and from _row_cells() rather than being written out twice.
#
#   view_w     column width in feet at the schedule view's 1:1 scale
#              (= printed inches / 12)
#   console_w  column width in monospace characters
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
                <CheckBox x:Name="cb_oa" Content="Multi System (pending)" IsEnabled="False"
                          ToolTip="Pending. Checks several connected systems in one run."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,0"
                           Text="Pending, for AHU, DOAS, VAV and FPB systems."/>

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
                            <Expander x:Name="exp_c" IsExpanded="False" Margin="0,8,0,0"
                                      Header="Fitting loss coefficients (C)"
                                      ToolTip="Defaults from the RJA SP loss worksheet (1985 ASHRAE fitting numbers). Cannot be negative.">
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


def _show_velocity_settings_xaml(use_theme):
    """The XAML settings dialog. Returns the 11 element tuple or None (cancel).

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
    cb_oa = N('cb_oa')
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
    cb_oa.IsChecked = False
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
          '0.012 corrugated flex',
          hvac_graph.STANDARD_FLEX_ROUGHNESS_FT),
         ('air_density', 'Air density (lb/ft3)',
          '0.075 standard air, sea level',
          hvac_graph.STANDARD_AIR_DENSITY_LB_FT3)],
        basis_boxes, basis_order, '%g')

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
        ('Roughness', '%g ft galvanized, %g ft flex (flex detected by category). '
                      'Editable under Advanced.'
                      % (hvac_graph.STANDARD_RIGID_ROUGHNESS_FT,
                         hvac_graph.STANDARD_FLEX_ROUGHNESS_FT)),
        ('Fittings', 'C x Pv, with C from the RJA SP loss worksheet (1985 ASHRAE '
                     'fitting numbers). Editable under Advanced.'),
        ('External static', 'The most restrictive run (index run) from the unit to '
                            'a terminal, supply path plus return path. Includes duct '
                            'friction, fittings, balancing dampers, the diffuser and '
                            'the safety factor. Excludes the coil and cabinet (already '
                            'in the published ESP), fire and backdraft dampers, and the '
                            'filter unless Include return filter is ticked.'),
    ]
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
    cb_static.Checked += sync_enable
    cb_static.Unchecked += sync_enable
    cb_filter.Checked += sync_enable
    cb_filter.Unchecked += sync_enable
    sync_enable()

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

        include_oa = bool(cb_oa.IsChecked)
        selected_cols = set(k for k, cb in col_boxes.items() if bool(cb.IsChecked))
        full_diag = bool(cb_full_diag.IsChecked)
        values = (out, gpct, include_oa, selected_cols, full_diag,
                  ext_static, safety_pct, c_values, comp_values, calc_basis,
                  filter_inwc)
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
    """The settings dialog. Returns ({sys_class: (max_fpm, max_friction_inwc)},
    tol_pct, include_oa, selected_column_keys, full_diag, ext_static, safety_pct,
    c_values, comp_values, calc_basis, filter_inwc) or None if cancelled.

    Tries the themed XAML dialog first, then the same dialog without the theme
    file, then the original code-built dialog, so the tool always opens. The
    meaning of every returned element is documented on
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
    """FALLBACK ONLY. Code-built WPF dialog — Outside Air scope, per-system max velocity + friction,
    a yellow/red tolerance %, and which columns the output tables show.

    Returns ({sys_class: (max_fpm, max_friction_inwc)}, tol_pct, include_oa,
    selected_column_keys, full_diag, ext_static, safety_pct,
    c_values, comp_values, calc_basis, filter_inwc) or None.

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

    include_oa is False by default (equipment-level: Supply Air +
    Return Air only — any mechanical equipment node's Outside Air branch is
    skipped entirely by hvac_graph.traverse(), not sized/colored, for
    projects where the OA network isn't modeled to completion). Checking
    the Outside Air box switches to system-level: full traversal including
    OA, for VAV/FCU networks where OA is actually being traced.

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

    # ── 2. Outside Air: equipment-level (SA+RA only, default) vs ────────────
    # ── system-level (traces OA too) ─────────────────────────────────────
    cb_oa = CheckBox()
    cb_oa_text = TextBlock()
    cb_oa_text.Text = ('Multi System (AHU/DOAS/VAV/FPB).  PENDING, under '
                       'development.')
    cb_oa_text.TextWrapping = TextWrapping.Wrap
    cb_oa_text.Width = CONTENT_W - 20
    cb_oa.Content = cb_oa_text
    cb_oa.IsChecked = False
    cb_oa.IsEnabled = False
    cb_oa.Margin  = Thickness(2, 0, 0, 2)
    outer.Children.Add(cb_oa)

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
        'Fitting loss coefficients (C)',
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
          u'0.012 corrugated flex',
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
    _info_row('Duct roughness:',  u'ε = 0.0003 ft galvanized, 0.012 ft flex '
                                  u'(detected by category, ~1.8× the friction). '
                                  u'Editable under Advanced settings.')
    _info_row('Fitting losses:',  u'C × Pv, C from RJA SP_LOSS_WORKSHEET '
                                  u'(1985 ASHRAE fitting numbers), editable above')
    _info_row('External static:', 'the single most restrictive run (index run) from '
                                  'the unit to a terminal, supply path + return path')
    _info_row('  includes:',      'duct friction, fittings (elbows, take-offs, '
                                  'transitions), balancing dampers, the diffuser, '
                                  'and the safety factor')
    _info_row('  excludes:',      'filter (unless Include return filter is ticked), coil '
                                  'and cabinet (already in the published ESP). Fire and '
                                  'backdraft dampers.')

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
            include_oa = bool(cb_oa.IsChecked)
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
            result[0] = (out, gpct, include_oa, selected_cols, full_diag,
                         ext_static, safety_pct, c_values, comp_values, calc_basis,
                         filter_inwc)
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
    if 'supply' in name:
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
                        duct_rooted_ids=None, return_filter_inwc=0.0):
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
    """
    dr_by_int_id = {}
    for dr in all_duct_results.values():
        dr_by_int_id[dr.element_id] = dr

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

    def _price_fitting(elem, sys_class, up_fpm, up_area):
        """(loss in. wc, counted 0/1, unpriced note or None, downstream area).

        One place that prices a fitting the path passes THROUGH, used by the main
        walk and by the unit-side chain below so the two can never disagree.
        """
        fam  = hvac_graph.fitting_family_name(elem)
        role = fitting_tables.classify_fitting(fam)
        up_a, down_a = (None, None)
        if role == 'transition':
            up_a, down_a = hvac_graph.transition_areas(elem, up_area)
        c, note = fitting_tables.fitting_c(
            role, sys_class,
            is_round=hvac_graph.fitting_is_round(elem),
            upstream_area_ft2=up_a, downstream_area_ft2=down_a,
            is_end_of_main=False, c=c_values)
        if c > 0.0:
            return c * hvac_graph.velocity_pressure_inwg(up_fpm), 1, None, down_a
        if 'UNPRICED' in note:
            return 0.0, 0, (fam or '?') + ': ' + note, down_a
        return 0.0, 0, None, down_a

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
                    loss, counted, note, down_a = _price_fitting(
                        elem, root_dr.sys_class, up_fpm, up_area)
                    fit   += loss
                    n_fit += counted
                    if note:
                        unpriced = unpriced + (note,)
                    if down_a:
                        up_area = down_a
        return (fric, fit, n_fit, n_duct, dlen, unpriced)

    best = {}              # root_id -> {sys_class: {...}}
    truncated_roots = set()  # PER ROOT: one unit truncating must not label the
                             # others, now that results print per equipment.

    # Bound on total path steps. A simple-path search is exponential in the
    # worst case, and while duct networks are near-trees, a badly modelled one
    # with many merge points could in principle blow up. Hitting this is
    # reported, never silently swallowed.
    MAX_PATH_STEPS = 200000

    for root_id in all_root_ids:
        best_here = {}     # this equipment's own per-system-class winners
        # (node, friction, fitting loss, ducts, fittings, taps bypassed, feet,
        #  upstream FPM, upstream area ft2, sys_class, unpriced notes,
        #  nodes already on THIS path)
        # A duct-rooted pick starts with the unit-side stretch already priced.
        pre = (0.0, 0.0, 0, 0, 0.0, ())
        if duct_rooted_ids and root_id in duct_rooted_ids:
            pre = _unit_side_chain(root_id)
        # pre = (friction, fitting loss, fittings, ducts, feet, unpriced)
        stack = [(root_id, pre[0], pre[1], (), pre[3], pre[2], 0, pre[4], 0.0,
                  0.0, None, pre[5], frozenset([root_id]))]
        steps = 0
        while stack:
            steps = steps + 1
            if steps > MAX_PATH_STEPS:
                log.warning('_critical_path_loss: path search truncated at %d '
                            'steps from root %s', MAX_PATH_STEPS, root_id)
                truncated_roots.add(root_id)
                break
            (nid, fric, fit, comp_keys, dcount, fcount, tcount, dlen,
             up_fpm, up_area, sys_class, unpriced, path) = stack.pop()

            elem = all_nodes.get(nid)
            dr   = dr_by_int_id.get(nid)

            if dr is not None:
                fric      = fric + dr.friction_loss_inwc
                dlen      = dlen + dr.length_ft
                dcount    = dcount + 1
                up_fpm    = dr.fpm
                up_area   = dr.area_ft2
                sys_class = dr.sys_class

            elif elem is not None and hvac_graph.is_fitting(elem):
                loss, counted, note, _down_a = _price_fitting(
                    elem, sys_class, up_fpm, up_area)
                fit    = fit + loss
                fcount = fcount + counted
                if note:
                    unpriced = unpriced + (note,)

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
                else:
                    unpriced = unpriced + (
                        (fam or '?') + ': accessory on the run, not priced',)

            term = all_terminals.get(nid)
            if term is not None:
                # The diffuser the run ends at is itself a component drop.
                keys_here = comp_keys + ('diffuser',)
                comp_here = 0.0
                for _k in keys_here:
                    comp_here += fitting_tables.component_of(_k, comp_values)
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
                extra_fit = fit
                extra_t   = tcount
                if dr is not None:
                    # Every tap on this main that the run does NOT leave
                    # through restricts it, and gets the main-duct coefficient.
                    taps = _taps(nid)
                    n_by = len(taps) - (1 if child_id in taps else 0)
                    if n_by > 0:
                        extra_fit = extra_fit + (
                            fitting_tables.takeoff_main_c(dr.sys_class, c_values) *
                            hvac_graph.velocity_pressure_inwg(dr.fpm) * n_by)
                        extra_t = extra_t + n_by
                    # A tap at the terminus of a main is the worksheet's "end of
                    # main branch" case and carries a different C, so the swing
                    # is applied here where the parent duct is known.
                    if child_id in taps and not _continues(nid):
                        ch = all_nodes.get(child_id)
                        if ch is not None:
                            end_c = fitting_tables.takeoff_c(
                                dr.sys_class, is_end_of_main=True, c=c_values)
                            typ_c = fitting_tables.takeoff_c(
                                dr.sys_class, c=c_values)
                            extra_fit = extra_fit + (
                                (end_c - typ_c) *
                                hvac_graph.velocity_pressure_inwg(dr.fpm))
                stack.append((child_id, fric, extra_fit, comp_keys, dcount,
                              fcount, extra_t, dlen, up_fpm, up_area,
                              sys_class, unpriced,
                              path | frozenset([child_id])))

        if best_here:
            best[root_id] = best_here

    result = {}
    for root_id, per_class in best.items():
        out = {}
        for sys_class, b in per_class.items():
            sub = (b['friction_inwc'] + b['fitting_inwc'] + b['component_inwc'] +
                   b['filter_inwc'])
            b['subtotal_inwc'] = sub
            b['total_inwc']    = sub * (1.0 + safety_pct / 100.0)
            b['truncated']     = root_id in truncated_roots
            out[sys_class]     = b
        result[root_id] = out
    return result


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
    branch feeding a slot diffuser also comes through here, because the design
    standard publishes no duct/neck size table to check a slot branch against.
    Behaviour is unchanged from before the split — see _branch_duct_label()
    for what every other branch gets instead, and _label_duct() for the
    dispatch between them.

      Green  : cfm <= max_cap
      Yellow : max_cap < cfm <= max_cap * (1 + tol_pct/100)
      Red    : cfm > max_cap * (1 + tol_pct/100), OR fails height clearance
      Purple : GREEN AND a smaller standard size exists (see _suggest_size_info)

    Returns (label, max_cap_cfm, reason) — reason is 'Velocity', 'Friction',
    'Oversized', 'Diffuser/Duct Clearance', or '' (GREEN/GRAY, nothing to flag).
    """
    defaults = hvac_graph.FIRM_DEFAULTS.get(dr.sys_class, (600, 0.05))
    max_fpm, max_friction = custom_limits.get(dr.sys_class, defaults)
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
    if combined == 'GREEN':
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

    if kind == 'slot':
        # No published duct/neck size breakpoints exist for slot diffusers
        # (the standard gives capacity per slot length and open-slot count
        # instead), so there is nothing to size-match against. Fall back to
        # the main check rather than report the branch as unverifiable.
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


# Standard spiral/round sizes in 2" increments (inches)
_ROUND_SIZES = [6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36]
_MIN_DUCT_DIM = 6  # firm standard: never suggest below 6" for rect or spiral


def _snap_even(inches):
    """Round up to the nearest even (2" interval) duct dimension — odd sizes
    like 17" aren't stocked/installed; the next even size up (18") is used.

    Subtracts a small epsilon before ceiling so floating-point noise from the
    Revit feet<->inch round-trip (e.g. a true 10" duct stored internally as
    10.000000000000002") doesn't get bumped up to the next even size (12").
    Real fractional dimensions (e.g. 8.5") still round up correctly since the
    epsilon is far smaller than any real duct dimension tolerance.
    """
    n = int(math.ceil(inches - 1e-6))
    if n % 2 != 0:
        n += 1
    return max(n, 2)


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
    defaults = hvac_graph.FIRM_DEFAULTS.get(dr.sys_class, (600, 0.05))
    max_fpm, max_friction = custom_limits.get(dr.sys_class, defaults)
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
    defaults = hvac_graph.FIRM_DEFAULTS.get(dr.sys_class, (600, 0.05))
    max_fpm, max_friction = custom_limits.get(dr.sys_class, defaults)
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

def _hline(doc, view, x0, x1, y):
    doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(x0, y, 0.0), XYZ(x1, y, 0.0)))

def _vline(doc, view, x, y0, y1):
    doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(x, y0, 0.0), XYZ(x, y1, 0.0)))


def _legend_swatch(doc, view, filled_region_type_id, x0, y0, size, color, fill_id):
    """Draw a small solid-color square at (x0, y0), size x size (ft), colored
    the same way ducts/fittings are colored in the view (matching fill pattern)."""
    pts = [XYZ(x0, y0, 0.0), XYZ(x0 + size, y0, 0.0),
           XYZ(x0 + size, y0 + size, 0.0), XYZ(x0, y0 + size, 0.0)]
    loop = CurveLoop()
    for i in range(4):
        loop.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
    fr = FilledRegion.Create(doc, filled_region_type_id, view.Id, [loop])
    ogs = OverrideGraphicSettings()
    ogs.SetSurfaceForegroundPatternColor(color)
    if fill_id != ElementId.InvalidElementId:
        ogs.SetSurfaceForegroundPatternId(fill_id)
    view.SetElementOverrides(fr.Id, ogs)
    return fr


# Color key + meaning shown in the legend, in display order.
# Wording covers both check methods, since one view mixes mains and branches.
_LEGEND_ROWS = [
    ('GREEN',  'Main: within limit.  Branch: diffuser and duct both correctly sized'),
    ('PURPLE', 'Main only: oversized — smaller standard size available'),
    ('YELLOW', 'Main only: approaching limit'),
    ('RED',    'Main: exceeds limit.  Branch: diffuser or duct undersized.  '
               'Either: fails diffuser/duct height clearance (see Reason column)'),
    ('GRAY',   'No airflow data — no CFM reaches this element; broken or disconnected system'),
]


def _build_summary_view(doc, summary_lines, flagged_rows, selected_cols,
                        source_sheet_num, tn_type_id, ts, fill_id,
                        show_r_marker_note=False):
    """Create a Drafting View with a System Summary block + flagged-duct table
    + color legend.

    flagged_rows: list of per-duct cell dicts from _row_cells(), already in
    the order they should appear — the same order the numbered keynote circles
    were placed in the plan view, so row 3 here is circle 3 there.
    selected_cols: set of _COLUMN_DEFS keys the user kept in the dialog.

    Returns (ViewDrafting, total_content_height_ft, table_width_ft), or
    (None, 0.0, 0.0) on failure. The width is returned rather than hardcoded
    at the call site because the column selection now changes it on every run,
    and the viewport is positioned by its centre — the caller needs the width
    to keep the table's left edge on its hand-placed anchor.

    At scale 1:1, model feet = paper feet, so all dims below are paper inches / 12.
    """
    drafting_type_id = None
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements():
        if vft.ViewFamily == ViewFamily.Drafting:
            drafting_type_id = vft.Id
            break
    if drafting_type_id is None:
        return None, 0.0, 0.0

    try:
        sched_view = ViewDrafting.Create(doc, drafting_type_id)
        sched_view.Scale = 1
        base_name = 'Duct Schedule - DV-{} - {}'.format(source_sheet_num, ts)
        try:
            sched_view.Name = base_name
        except Exception:
            sched_view.Name = base_name + ' (2)'

        # ── Layout (ft at 1:1 = inches on paper / 12) ─────────────────────
        ox, oy = 0.0, 0.0   # top-left origin
        PAD     = 0.004      # text inset from cell edge (~1/24")
        HEAD_H  = 0.030      # header row height  (~3/8")
        ROW_H   = 0.022      # data row height    (~1/4")
        SUM_ROW_H = 0.018    # summary line height (~1/5")

        # Columns come from the shared _COLUMN_DEFS, filtered to the user's
        # selection — same list, same order, same headers as the console
        # table. The total width also sets the summary block width, and is
        # returned to the caller so the viewport can stay left-anchored.
        cols        = _selected_columns(selected_cols)
        col_headers = [c[1] for c in cols]
        col_widths  = [c[2] for c in cols]
        # sum(..., 0.0) forces float — sum([]) returns int 0 in Python 2.7
        total_w     = sum(col_widths, 0.0)

        opts = TextNoteOptions(tn_type_id)
        y_cursor = oy

        # ── System Summary block ────────────────────────────────────────────
        if summary_lines:
            sum_top = y_cursor
            for line in summary_lines:
                y_cursor -= SUM_ROW_H
            sum_bottom = y_cursor

            _hline(doc, sched_view, ox, ox + total_w, sum_top)
            _hline(doc, sched_view, ox, ox + total_w, sum_bottom)
            _vline(doc, sched_view, ox, sum_top, sum_bottom)
            _vline(doc, sched_view, ox + total_w, sum_top, sum_bottom)

            row_y = sum_top
            for line in summary_lines:
                row_y -= SUM_ROW_H
                TextNote.Create(doc, sched_view.Id,
                                XYZ(ox + PAD, row_y + SUM_ROW_H - PAD, 0.0),
                                line, opts)

        # ── Flagged ducts table ──────────────────────────────────────────────
        if flagged_rows:
            table_top = y_cursor
            total_h   = HEAD_H + ROW_H * len(flagged_rows)

            col_xs = [ox]
            for w in col_widths:
                col_xs.append(col_xs[-1] + w)

            row_tops = [table_top]
            row_tops.append(table_top - HEAD_H)
            for _ in range(len(flagged_rows)):
                row_tops.append(row_tops[-1] - ROW_H)

            for y in row_tops:
                _hline(doc, sched_view, ox, ox + total_w, y)
            for x in col_xs:
                _vline(doc, sched_view, x, table_top, table_top - total_h)

            for ci, header in enumerate(col_headers):
                TextNote.Create(doc, sched_view.Id,
                                XYZ(col_xs[ci] + PAD, table_top - PAD, 0.0),
                                header, opts)

            for ri, cells in enumerate(flagged_rows):
                row_y = row_tops[ri + 1] - PAD
                for ci, col in enumerate(cols):
                    # '#' is this table's own row index — it matches the
                    # numbered keynote circle placed on the same duct in the
                    # plan view, which is why it is not baked into the cell
                    # dict (the console table numbers the same ducts
                    # differently).
                    cell_text = str(ri + 1) if col[0] == 'num' else cells.get(col[0], '')
                    # Blank cells are real (e.g. no Reason on a passing row), and
                    # TextNote.Create rejects an empty string — leave the cell
                    # empty rather than write a placeholder into the drawing.
                    if not cell_text:
                        continue
                    TextNote.Create(doc, sched_view.Id,
                                    XYZ(col_xs[ci] + PAD, row_y, 0.0),
                                    cell_text, opts)
            y_cursor = table_top - total_h

        # ── Color legend ─────────────────────────────────────────────────────
        LEGEND_HDR_H = 0.026     # header row height (~5/16")
        LEGEND_ROW_H = 0.020     # legend row height (~1/4")
        SWATCH       = 0.014     # swatch square size (~1/6")

        frt_id = None
        for frt in FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements():
            frt_id = frt.Id
            break

        TextNote.Create(doc, sched_view.Id, XYZ(ox + PAD, y_cursor - PAD, 0.0), 'COLOR LEGEND', opts)
        y_cursor -= LEGEND_HDR_H

        if frt_id is not None:
            for color_key, meaning in _LEGEND_ROWS:
                color = _COLOR_MAP[color_key]
                sw_y0 = y_cursor - SWATCH - (LEGEND_ROW_H - SWATCH) / 2.0
                _legend_swatch(doc, sched_view, frt_id, ox + PAD, sw_y0, SWATCH, color, fill_id)
                TextNote.Create(doc, sched_view.Id,
                                XYZ(ox + PAD + SWATCH + PAD * 2.0, y_cursor - PAD, 0.0),
                                '{} — {}'.format(color_key.title(), meaning), opts)
                y_cursor -= LEGEND_ROW_H

        if show_r_marker_note:
            TextNote.Create(
                doc, sched_view.Id, XYZ(ox + PAD, y_cursor - PAD, 0.0),
                '(R) — Most restrictive diffuser: ends the index run that sets '
                'required external static pressure for its equipment/system',
                opts)
            y_cursor -= LEGEND_ROW_H

        return sched_view, (oy - y_cursor), total_w

    except Exception:
        return None, 0.0, 0.0


# ── main ───────────────────────────────────────────────────────────────────────
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
    (custom_limits, tol_pct, include_oa, selected_cols, full_diag,
     ext_static, safety_pct, c_values, comp_values, calc_basis,
     filter_inwc) = dialog_result

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

    if include_oa:
        output.print_md('Scope: **System-level (Supply, Return, Outside Air — upstream and downstream)**')
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

    for sel_elem in sel_elems:
        output.print_md('Traversing **{}** (id {})...'.format(
            _elem_name(sel_elem), eid_int(sel_elem.Id)))
        net = hvac_graph.build_network(sel_elem, doc, equipment_level=(not include_oa))

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

    summary_lines = ['SYSTEM SUMMARY']
    for name, aid, discharge_cfm, term_count, other_class_cfm in ahu_totals:
        summary_lines.append('Equipment: {} (id {})  -  {:.0f} CFM supply air discharge ({} diffusers)'.format(
            name, aid, discharge_cfm, term_count))
        for cls in sorted(other_class_cfm.keys()):
            summary_lines.append('    + {:.0f} CFM {} feeding this equipment (not counted in discharge total)'.format(
                other_class_cfm[cls], cls))
    summary_lines.append('Grand Total Airflow: {:.0f} CFM  |  {} Diffusers  |  {} Ducts'.format(
        grand_total_cfm, len(all_terminals), len(all_duct_results)))
    for sys_class in sorted(sys_totals.keys()):
        s_cfm, s_cnt = sys_totals[sys_class]
        summary_lines.append('  {}: {:.0f} CFM  ({} diffusers)'.format(sys_class, s_cfm, s_cnt))
    summary_lines.append('Diffuser Types:')
    for family in sorted(type_totals.keys()):
        t_cnt, t_cfm = type_totals[family]
        summary_lines.append('  {}  x{}  ({:.0f} CFM)'.format(family, t_cnt, t_cfm))
    if all_zero_terms:
        summary_lines.append('WARNING: {} diffuser(s) with Flow = 0 (missing CFM data)'.format(
            len(all_zero_terms)))
    if all_missing_flow:
        summary_lines.append('WARNING: {} diffuser(s) missing a Flow parameter entirely'.format(
            len(all_missing_flow)))

    # Terminal ids of the most-restrictive diffuser per equipment/system class
    # — the one the index run (_critical_path_loss) ends at, i.e. the one that
    # actually sizes the fan's external static. Only populated when External
    # Static is on, since that is the only place this path search runs.
    # Marked in the view with a circled "R" (see transaction below), on top
    # of the diffuser's real pass/fail color — never replacing it.
    restrictive_terminal_ids = set()

    if ext_static:
        critical = _critical_path_loss(
            all_root_ids, all_children, all_duct_results, all_terminals,
            all_nodes, safety_pct, c_values, comp_values,
            duct_rooted_ids=duct_rooted_ids, return_filter_inwc=filter_inwc)
        summary_lines.append(
            'TOTAL EXTERNAL STATIC PRESSURE (index run, per RJA SP_LOSS_WORKSHEET)')

        # Grouped by EQUIPMENT, because every unit has its own fan and so its own
        # external static. Merging them would let a big unit's index run hide a
        # small one's completely.
        one_sided = []   # (root_id, total_inwc, side names, is_supply) for roots with one side only
        for root_id in all_root_ids:
            per_class = critical.get(root_id)
            if not per_class:
                continue
            summary_lines.append('')
            summary_lines.append('  === {} ==='.format(
                root_labels.get(root_id, 'id {}'.format(root_id))))

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
                summary_lines.append(
                    '      Duct friction      {:.3f} in. wc   ({} ducts, {:.0f} ft)'.format(
                        c['friction_inwc'], c['duct_count'], c['length_ft']))
                summary_lines.append(
                    '      Fitting losses     {:.3f} in. wc   ({} fittings, {} taps '
                    'restricting the main)'.format(
                        c['fitting_inwc'], c['fitting_count'], c['tap_bypass_count']))
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
                summary_lines.append(
                    '      Components         {:.3f} in. wc   ({})'.format(
                        c['component_inwc'],
                        '; '.join(parts) if parts else 'none on this run'))
                if c.get('filter_inwc', 0.0) > 0.0:
                    summary_lines.append(
                        '      Return filter      {:.3f} in. wc   (known drop, '
                        'Include Return Filter)'.format(c['filter_inwc']))
                if c.get('unit_side_fittings', 0) > 0:
                    summary_lines.append(
                        '      (Includes {} fitting(s) between the picked duct and '
                        'the unit connection.)'.format(c['unit_side_fittings']))
                summary_lines.append(
                    '      Subtotal           {:.3f} in. wc'.format(c['subtotal_inwc']))
                summary_lines.append(
                    '      Safety factor      {:.3f} in. wc   (+{:.0f}%)'.format(
                        c['total_inwc'] - c['subtotal_inwc'], safety_pct))
                summary_lines.append(
                    '      TOTAL              {:.3f} in. wc   (worst run ends at {})'.format(
                        c['total_inwc'], c['terminal_name']))
                for u in c['unpriced']:
                    summary_lines.append('      UNPRICED: ' + u)
                if c.get('truncated'):
                    summary_lines.append(
                        '      WARNING: path search hit its step limit, so this '
                        'may not be the true worst run.')

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
                summary_lines.append(
                    '    {}:  {:,.0f} CFM ({} terminals), but no index run was '
                    'traced - no static pressure for it.'.format(
                        sys_class, a_cfm, a_cnt))

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
            elif sides:
                # One side only. Held back, not warned about yet: whether it is
                # half of a system depends on what else was selected.
                one_sided.append((root_id, unit_total, sides, supply is not None))

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
            summary_lines.append('')
            summary_lines.append('  === {} + {} ==='.format(
                root_labels.get(s_o[0], 'id {}'.format(s_o[0])),
                root_labels.get(r_o[0], 'id {}'.format(r_o[0]))))
            summary_lines.append(
                '    EXTERNAL STATIC PRESSURE = {:.3f} in. wc   ({} + {})'.format(
                    s_o[1] + r_o[1], s_o[2][0], r_o[2][0]))
        leftover = [o for o in one_sided if o[0] not in paired]
        for _rid, _tot, _sides, _is_sup in leftover:
            summary_lines.append('')
            summary_lines.append('  === {} ==='.format(
                root_labels.get(_rid, 'id {}'.format(_rid))))
            summary_lines.append(
                '    EXTERNAL STATIC PRESSURE = {:.3f} in. wc   ({} only)'.format(
                    _tot, ' + '.join(_sides)))
        if len([o for o in leftover if o[0] in duct_rooted_ids]) > 1:
            summary_lines.append('')
            summary_lines.append(
                '  WARNING: more than one system was selected, so these one-sided '
                'runs cannot be paired into a single external static. Pick one '
                'supply duct and one return duct per system.')

        summary_lines.append('')
        summary_lines.append(
            '  METHOD: Darcy-Weisbach + Altshul-Tsal, eps {:g} ft rigid / '
            '{:g} ft flex, air density {:g} lb/ft3. Fittings C x Pv per RJA '
            'SP_LOSS_WORKSHEET (1985 ASHRAE). '
            'Dovetail take-offs, rect elbows vaned.'.format(
                calc_basis['rigid_roughness'], calc_basis['flex_roughness'],
                calc_basis['air_density']))
        if filter_inwc > 0.0:
            summary_lines.append(
                '  NOT INCLUDED: coil and cabinet, already in the published ESP. '
                'Return filter IS included as a {:.3f} in. wc known drop. Fire and '
                'backdraft dampers not priced.'.format(filter_inwc))
        else:
            summary_lines.append(
                '  NOT INCLUDED: filter, coil and cabinet, already in the published '
                'ESP. Fire and backdraft dampers not priced.')

    # Folded once here so the drafting view and the console render identical
    # text. print_code, not print_md: markdown collapses leading whitespace, and
    # the per-equipment blocks are indented to show what belongs to which unit.
    summary_lines = _wrap_summary_lines(summary_lines)

    output.print_md('---')
    output.print_md('### System Summary')
    output.print_code('\n'.join(summary_lines[1:]))

    # 5. Find source sheet number
    source_sheet_num = 'NoSheet'
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        for vpid in sheet.GetAllViewports():
            vp = doc.GetElement(vpid)
            if vp is not None and vp.ViewId == active_view.Id:
                source_sheet_num = sheet.SheetNumber
                break

    output.print_md('Source sheet: **{}**'.format(source_sheet_num))

    # 6. Title block and solid fill
    tb_list = list(FilteredElementCollector(doc)
                   .OfCategory(BuiltInCategory.OST_TitleBlocks)
                   .WhereElementIsElementType())
    tb_id   = tb_list[0].Id if len(tb_list) > 0 else ElementId.InvalidElementId
    fill_id = hvac_graph.solid_fill_pattern_id(doc)

    # 7. Text height: aim for 5/64" printed size at the view's print scale
    view_scale = getattr(active_view, 'Scale', 48)
    text_h_ft  = (5.0 / (64.0 * 12.0)) * float(view_scale)

    # 8. Transaction: copy view → color overrides → FPM annotations → sheet
    t = Transaction(doc, 'Duct Performance')
    t.Start()
    ts = datetime.datetime.now().strftime('%Y-%m-%d-%H%M%S')

    try:
        # Copy floor plan
        new_vid  = active_view.Duplicate(ViewDuplicateOption.Duplicate)
        new_view = doc.GetElement(new_vid)
        base_name = 'Duct Performance - {} - {}'.format(source_sheet_num, ts)
        try:
            new_view.Name = base_name
        except Exception:
            new_view.Name = base_name + ' (2)'

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
            label, green_cap, reason, branch_res = _label_duct(
                dr, custom_limits, tol_pct, downstream_h, branch_ctx.get(eid))
            duct_labels[eid]    = (label, green_cap, reason)
            branch_results[eid] = branch_res
            if reason == 'Diffuser/Duct Clearance':
                clearance_count += 1
            counts[label] = counts.get(label, 0) + 1
            color = _COLOR_MAP[label]
            ogs   = OverrideGraphicSettings()
            ogs.SetSurfaceForegroundPatternColor(color)
            if fill_id != ElementId.InvalidElementId:
                ogs.SetSurfaceForegroundPatternId(fill_id)
            ogs.SetProjectionLineColor(color)
            new_view.SetElementOverrides(eid, ogs)

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
            if not (hvac_graph.is_fitting_or_accessory(elem)
                    or hvac_graph.is_terminal(elem)):
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
            color = _COLOR_MAP[worst]
            ogs   = OverrideGraphicSettings()
            ogs.SetSurfaceForegroundPatternColor(color)
            if fill_id != ElementId.InvalidElementId:
                ogs.SetSurfaceForegroundPatternId(fill_id)
            ogs.SetProjectionLineColor(color)
            new_view.SetElementOverrides(elem.Id, ogs)
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

        for idx, (lbl, dr, _cells) in enumerate(flagged_items, 1):
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
            if term_elem is None:
                continue
            try:
                r_pt = term_elem.Location.Point
            except Exception:
                continue
            if r_pt is None:
                continue
            # Offset off the diffuser symbol itself so the marker doesn't
            # sit directly on top of it.
            mark_pt = XYZ(r_pt.X + text_h_ft * 2.0, r_pt.Y + text_h_ft * 2.0, r_pt.Z)
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

        # System Summary + flagged-duct table + legend, placed as second viewport on sheet
        if tn_type_id is not None:
            sched_view, content_h, total_w = _build_summary_view(
                doc, summary_lines, [item[2] for item in flagged_items],
                selected_cols, source_sheet_num, tn_type_id, ts, fill_id,
                show_r_marker_note=bool(restrictive_terminal_ids))
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
                # No title/border on this viewport — use the "None" Viewport Type if
                # present. FilteredElementCollector on OST_Viewports + WhereElementIsElementType
                # misses this system-family type entirely (confirmed via MCP inspection).
                # Best-effort only: cosmetic, must never break the whole tool if this
                # lookup misbehaves (confirmed AttributeError on vp_type.Name here even
                # though the identical query works fine outside pyRevit's IronPython).
                try:
                    for vp in FilteredElementCollector(doc).OfClass(Viewport):
                        vp_type = doc.GetElement(vp.GetTypeId())
                        if vp_type is None:
                            continue
                        vp_type_name = vp_type.Name
                        if vp_type_name is not None and vp_type_name.strip().lower() == 'none':
                            sched_vp.ChangeTypeId(vp_type.Id)
                            break
                except Exception:
                    pass

        t.Commit()
    except Exception as ex:
        t.RollBack()
        output.print_md('**Error — transaction rolled back:** {}: {}'.format(
            type(ex).__name__, str(ex)))
        output.print_md('```\n{}\n```'.format(traceback.format_exc()))
        return

    # 9. Summary
    output.print_md('---')
    output.print_md('## Done')
    output.print_md('Sheet **{}** created.'.format(new_sheet.SheetNumber))
    output.print_md('')
    output.print_md('**Mains and branches were checked by different methods.** '
                    'A duct is a branch when exactly one terminal is reachable '
                    'downstream of it, a main when two or more are.')
    output.print_md('')
    output.print_md('**Main ducts — design limits used  (green ≤ max,  yellow = within {}% above max,  red > max+{}%,  purple = a smaller standard size still fits):**'.format(
        int(tol_pct), int(tol_pct)))
    output.print_md('| System | Max Velocity | Max Friction |')
    output.print_md('| --- | --- | --- |')
    for sys_class in ('Supply Air', 'Return Air', 'Exhaust Air', 'Outside Air', 'Transfer Air'):
        defaults = hvac_graph.FIRM_DEFAULTS.get(sys_class, (600, 0.05))
        mx_fpm, mx_fric = custom_limits.get(sys_class, defaults)
        output.print_md('| {} | {:.0f} FPM | {:.3f} iwc/100 |'.format(
            sys_class, mx_fpm, mx_fric))
    output.print_md('')
    output.print_md('**Branch ducts — checked against the diffuser sizing tables '
                    '(MEP Design Standards, section 1), not against velocity or friction.** '
                    'An auto-sizing diffuser is selected on static pressure and NC rather '
                    'than duct velocity, so velocity-checking a run that feeds only that one '
                    'diffuser is redundant. Two things are checked instead: whether the '
                    'diffuser\'s own neck/face is big enough for its own CFM, and whether the '
                    'branch duct is at least as big as the diffuser connects with. Either one '
                    'failing is red; there is no yellow band on a size match. Oversize is '
                    'not checked on branches.')
    output.print_md('')
    output.print_md('Branches feeding a **slot diffuser** fall back to the main '
                    'velocity/friction check above — the design standard publishes no '
                    'duct or neck size breakpoints for slot diffusers to check against.')
    output.print_md('')
    output.print_md('| Color | Ducts | Fittings, Accessories & Diffusers | Meaning |')
    output.print_md('| --- | --- | --- | --- |')
    output.print_md('| Green  | {} | {} | Main: within limit. Branch: diffuser and duct both correctly sized |'.format(
        counts.get('GREEN',  0), fitting_counts.get('GREEN',  0)))
    output.print_md('| Purple | {} | {} | Main only: oversized (low velocity, review for cost) |'.format(
        counts.get('PURPLE', 0), fitting_counts.get('PURPLE', 0)))
    output.print_md('| Yellow | {} | {} | Main only: approaching limit |'.format(
        counts.get('YELLOW', 0), fitting_counts.get('YELLOW', 0)))
    output.print_md('| Red    | {} | {} | Main: exceeds limit. Branch: diffuser or duct undersized |'.format(
        counts.get('RED',    0), fitting_counts.get('RED',    0)))
    output.print_md('| Gray   | {} | {} | No airflow data: broken or disconnected system (not in the flagged table) |'.format(
        counts.get('GRAY',   0), fitting_counts.get('GRAY',   0)))
    output.print_md('')
    output.print_md('**Diffuser/duct height clearance issues (flagged Red): {}**'.format(
        clearance_count))

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
        # schedule view — only the fixed monospace widths differ, since this
        # table is padded text rather than ruled cells.
        cols = _selected_columns(selected_cols)

        def _fmt_row(cells):
            return '  '.join(str(c).ljust(col[3]) for c, col in zip(cells, cols))

        header    = _fmt_row([col[1] for col in cols])
        separator = _fmt_row(['-' * col[3] for col in cols])
        rows      = [header, separator]

        for idx, (label, fpm, cells) in enumerate(flagged, 1):
            rows.append(_fmt_row([
                str(idx) if col[0] == 'num' else cells.get(col[0], '')
                for col in cols
            ]))

        output.print_md('')
        output.print_md('### All Ducts' if full_diag else '### Flagged Ducts')
        output.print_code('\n'.join(rows))

    uidoc.ActiveView = new_sheet


main()
