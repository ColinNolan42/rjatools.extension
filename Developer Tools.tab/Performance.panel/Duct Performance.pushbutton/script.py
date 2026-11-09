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
judged against the MP limits, purple when a smaller size meets them (Colin,
2026-11-02). After the box, as above. A diffuser on an MP main is red. External static
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
    FilledRegion, FilledRegionType, CurveLoop,
    ConnectorProfileType, ConnectorType
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
    Separator, TextBlock, CheckBox, RadioButton, Expander, ScrollViewer,
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
import fan_system_effect
import diffuser_tables
import revit_helpers
import rja_drafting
import vav_checks
import console_report
import duct_perf_console
import ui_helpers
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
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,6"
                           Text="One unit per run. The scope sets which equipment the next list shows."/>
                <RadioButton x:Name="rb_scope_rtu" GroupName="scope" Content="RTU"
                             ToolTip="A constant volume unit with supply and return: RTU, AHU, furnace or fan coil. VAV boxes and exhaust fans are not listed."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,0"
                           Text="Lists units only, no VAV boxes or exhaust fans."/>
                <RadioButton x:Name="rb_scope_vav" GroupName="scope" Content="VAV System" Margin="0,8,0,0"
                             ToolTip="The full system from the unit: supply mains to the VAV box inlets checked as medium pressure (MP), the boxes, low pressure after them and the return, with external static."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,0"
                           Text="Lists units only (RTU, AHU). Boxes are found by position: supply equipment entered through its inlet. Medium pressure limits are under Main duct limits."/>
                <RadioButton x:Name="rb_scope_vavbox" GroupName="scope" Content="VAV Box" Margin="0,8,0,0"
                             ToolTip="One VAV box: the low pressure supply after the box, checked like an RTU's supply side, with external static from the box outlet to the worst diffuser. No medium pressure checks and no fan system effect."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,0"
                           Text="Lists VAV boxes only. Only the low pressure ductwork after the box is checked. External static runs from the box outlet only, not the total system loss."/>
                <RadioButton x:Name="rb_scope_exhaust" GroupName="scope" Content="Exhaust System" Margin="0,8,0,0"
                             ToolTip="Lists exhaust fans only, and reads a ceiling exhaust fan (its only duct connection is Exhaust Air, out)."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,0"
                           Text="Lists exhaust fans only."/>

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
                <CheckBox x:Name="cb_full_diag" Content="Report all ducts"
                          ToolTip="Unchecked, the report lists only ducts that fail or can be downsized (red, yellow, purple). Checked, every duct is listed and numbered in the plan."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,6"
                           Text="Unchecked lists only the ducts that fail or can be downsized (red, yellow, purple)."/>
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
                        <StackPanel x:Name="pnl_filter" Orientation="Horizontal" Margin="0,8,0,0">
                            <CheckBox x:Name="cb_filter" Content="Include return filter" VerticalAlignment="Center"
                                      ToolTip="A known filter drop added once to the return path. Leave off when the filter is inside the unit and already deducted from its published ESP."/>
                            <TextBox x:Name="tb_filter" Width="56" Margin="8,0,6,0"
                                     ToolTip="Known filter drop in in. wc. Cannot be negative."/>
                            <TextBlock Text="in. wc known drop (MERV 8), added to the return total" VerticalAlignment="Center"/>
                        </StackPanel>
                        <CheckBox x:Name="cb_fse" Content="Include fan system effect" Margin="0,10,0,0"
                                  ToolTip="Prices the first elbow past each unit connection from the 2001 ASHRAE Fundamentals fan system effect tables and adds it before the safety factor. Transitions and taps there are flagged, not priced."/>
                        <StackPanel x:Name="pnl_fse" Margin="20,4,0,0">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="Fan type" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                <RadioButton x:Name="rb_fan_swsi" GroupName="fan_type" Content="SWSI" VerticalAlignment="Center" Margin="0,0,14,0"
                                             ToolTip="Outlet tables SR7-5 to SR7-8. Set by the Scope: Exhaust System."/>
                                <RadioButton x:Name="rb_fan_dwdi" GroupName="fan_type" Content="DWDI" VerticalAlignment="Center"
                                             ToolTip="Outlet tables SR7-9 to SR7-12. Set by the Scope: RTU and VAV System."/>
                            </StackPanel>
                            <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,0"
                                       Text="SWSI, single width single inlet: air enters one side of the wheel (exhaust fans, utility sets). DWDI, double width double inlet: air enters both sides (most RTU and AHU blowers)."/>
                            <StackPanel Orientation="Horizontal" Margin="0,6,0,0">
                                <TextBlock Text="Blast area ratio Ab/Ao" VerticalAlignment="Center"/>
                                <TextBox x:Name="tb_blast" Width="56" Margin="8,0,6,0"
                                         ToolTip="Fan blast area over outlet area, from the fan data when known. Range 0.4 to 1.0."/>
                                <TextBlock Text="0.4 to 1.0" VerticalAlignment="Center"/>
                            </StackPanel>
                            <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,0"
                                       Text="Ab = blast area, the discharge opening at the scroll cutoff. Ao = outlet area, H x W at the fan discharge. Both at the fan outlet, from the fan cut sheet, 0.4 to 1.0."/>
                            <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                                <TextBlock Text="Fan velocity" VerticalAlignment="Center"/>
                                <TextBox x:Name="tb_fan_fpm" Width="56" Margin="8,0,6,0"
                                         ToolTip="Fan velocity, fpm. Must be greater than 0."/>
                                <TextBlock Text="fpm" VerticalAlignment="Center"/>
                            </StackPanel>
                            <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,0"
                                       Text="The unit connection velocity (cfm over the connector area) is used when it is higher."/>
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
                        <StackPanel x:Name="pnl_exh_static" Margin="0,10,0,0" Visibility="Collapsed">
                            <TextBlock FontWeight="SemiBold" Text="Exhaust fans"/>
                            <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                                <CheckBox x:Name="cb_exh_backdraft" Content="Include backdraft damper" Width="190" VerticalAlignment="Center"
                                          ToolTip="Adds the backdraft damper drop once to an exhaust fan's run."/>
                                <TextBox x:Name="tb_backdraft" Width="56" Margin="0,0,6,0"
                                         ToolTip="Backdraft damper drop, in. wc. Cannot be negative."/>
                                <TextBlock Text="in. wc, added once where none is modeled" VerticalAlignment="Center"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal" Margin="0,6,0,0">
                                <CheckBox x:Name="cb_exh_cap" Content="Include roof / wall cap" Width="190" VerticalAlignment="Center"
                                          ToolTip="Adds the roof / wall cap drop once to every exhaust run."/>
                                <TextBox x:Name="tb_exh_cap" Width="56" Margin="0,0,6,0"
                                         ToolTip="Roof / wall cap drop, in. wc. Cannot be negative."/>
                                <TextBlock Text="in. wc, added once to every exhaust run" VerticalAlignment="Center"/>
                            </StackPanel>
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
                            <Expander x:Name="exp_dtab" IsExpanded="False" Margin="0,4,0,0"
                                      Header="Diffuser sizing tables (MEP Design Standards)"
                                      ToolTip="The capacities branch ducts are checked against. Reference only, read only.">
                                <StackPanel x:Name="pnl_dtab" Margin="16,4,0,4"/>
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


def _scope_of(vav, vavbox, exhaust):
    """The Scope key from the dialog's choice."""
    if vavbox:
        return 'vavbox'
    return 'vav' if vav else 'exhaust' if exhaust else 'rtu'


def _settings_result(limits, tol_pct, selected_cols, full_diag, ext_static,
                     safety_pct, c_values, comp_values, calc_basis, filter_inwc,
                     vav, vav_settings, fan_effect=None, exhaust_cap=False,
                     exhaust=False, exhaust_settings=None, scope=None):
    """The settings dialog's answer, a dict, built in ONE place for both the
    themed and the legacy dialog so main() reads the same keys from either.

    With VAV ticked, the MP limits are also added to limits under
    vav_checks.MP_KEY, so every limit lookup finds them in one dict.

    fan_effect: {'on', 'fan_type', 'blast_ratio'}; None (the legacy dialog,
    which has no fields for it) takes fan_system_effect.default_settings().
    exhaust_settings: {'backdraft' (Include backdraft damper)}; a missing
    'backdraft' is taken as off in main() (Colin, 2026-11-03).
    'scope': 'rtu', 'vav', 'vavbox' or 'exhaust' (Colin, 2026-11-03: one
    Scope choice). None derives it from vav / exhaust (no VAV Box).
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
        'exhaust_cap':   bool(exhaust_cap),
        'exhaust':       bool(exhaust),
        # Empty: no backdraft damper added.
        'exhaust_settings': dict(exhaust_settings or {}),
        'vav':           bool(vav),
        'scope':         scope or _scope_of(vav, False, exhaust),
        'vav_settings':  vav_settings,
        'fan_effect':    dict(fan_effect or fan_system_effect.default_settings()),
    }


def _dialog_vav_rows():
    """The VAV rows of the dialog's Calculation basis: VAV static only
    (Colin, 2026-11-03). The sheet legend and console keep every row."""
    return [r for r in vav_checks.basis_rows() if r[0] == 'VAV, static']


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
    rb_scope_rtu = N('rb_scope_rtu')
    rb_scope_vav = N('rb_scope_vav')
    rb_scope_vavbox = N('rb_scope_vavbox')
    rb_scope_exhaust = N('rb_scope_exhaust')
    pnl_exh_static = N('pnl_exh_static')
    cb_exh_backdraft = N('cb_exh_backdraft')
    tb_backdraft = N('tb_backdraft')
    tb_exh_cap = N('tb_exh_cap')
    pnl_vav_static = N('pnl_vav_static')
    grid_vav_static = N('grid_vav_static')
    grid_limits = N('grid_limits')
    tb_tol = N('tb_tol')
    grid_cols = N('grid_cols')
    cb_static = N('cb_static')
    pnl_static = N('pnl_static')
    tb_sf = N('tb_sf')
    cb_filter = N('cb_filter')
    pnl_filter = N('pnl_filter')
    cb_exh_cap = N('cb_exh_cap')
    tb_filter = N('tb_filter')
    cb_fse = N('cb_fse')
    pnl_fse = N('pnl_fse')
    rb_fan_dwdi = N('rb_fan_dwdi')
    rb_fan_swsi = N('rb_fan_swsi')
    tb_blast = N('tb_blast')
    tb_fan_fpm = N('tb_fan_fpm')
    exp_adv = N('exp_adv')
    exp_c = N('exp_c')
    exp_comp = N('exp_comp')
    exp_basis = N('exp_basis')
    grid_c = N('grid_c')
    pnl_ctab = N('pnl_ctab')
    pnl_dtab = N('pnl_dtab')
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
    _fse_def = fan_system_effect.default_settings()
    tb_blast.Text = '%g' % _fse_def['blast_ratio']
    cb_fse.IsChecked = bool(_fse_def['on'])
    rb_fan_swsi.IsChecked = _fse_def['fan_type'] == 'SWSI'
    rb_fan_dwdi.IsChecked = _fse_def['fan_type'] != 'SWSI'
    tb_fan_fpm.Text = '%g' % _fse_def['fan_fpm']
    tb_backdraft.Text = '%.2f' % fitting_tables.component_of('backdraft')
    tb_exh_cap.Text = '%.2f' % fitting_tables.component_of('cap')
    for b in (tb_tol, tb_sf, tb_filter, tb_blast, tb_fan_fpm, tb_backdraft, tb_exh_cap):
        ui_helpers.rja_attach_clear_on_edit(b)
        all_boxes.append(b)
    cb_full_diag.IsChecked = False
    rb_scope_rtu.IsChecked = True    # one unit per run (Colin, 2026-11-03)
    cb_static.IsChecked = False
    cb_filter.IsChecked = True       # default on (Colin, 2026-11-02)
    cb_exh_backdraft.IsChecked = False  # default off (Colin, 2026-11-03)
    cb_exh_cap.IsChecked = False     # off: an upblast roof fan has no cap

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
    # The VAV medium pressure supply row, shown only with VAV System ticked
    # (Colin, 2026-11-02: replaces the VAV dropdown under Scope).
    vav_boxes = {}
    mp_row = []
    r = len(_DV_ROWS) + 1
    _add_row_def(grid_limits)
    lb = TextBlock()
    lb.Text = 'Supply Air (medium pressure)'
    lb.VerticalAlignment = VerticalAlignment.Center
    lb.ToolTip = 'VAV supply mains from the unit to the box inlets.'
    _put(grid_limits, lb, 0, r)
    mp_row.append(lb)
    for col, (key, label, _unit, default, _kind) in zip((1, 2), vav_checks.VAV_LIMIT_FIELDS):
        tb = _box(vav_checks.format_default(default), 80, Thickness(0, 2, 24, 2))
        tb.ToolTip = label + ', ' + _unit + '. Must be greater than 0.'
        _put(grid_limits, tb, col, r)
        vav_boxes[key] = tb
        mp_row.append(tb)

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

    # Diffuser sizing tables, reference only (Colin, 2026-10-23). Straight from
    # diffuser_tables, the same data the branch check reads. Filled on expand.
    for d_title, d_basis, d_head, d_rows in diffuser_tables.reference_tables():
        dex = Expander()
        dex.Header = d_title
        dex.IsExpanded = False
        dex.Margin = Thickness(0, 2, 0, 2)
        dbody = StackPanel()
        dbody.Margin = Thickness(18, 2, 0, 4)
        dex.Content = dbody

        def _fill_d(sender, e, body=dbody, basis=d_basis, head=d_head, rows=d_rows):
            if body.Children.Count:
                return
            body.Children.Add(ui_helpers.rja_hint_block(win, basis))
            if not head:
                return
            g = Grid()
            g.Margin = Thickness(0, 2, 0, 0)
            for _c in range(len(head)):
                cd = ColumnDefinition()
                cd.Width = GridLength.Auto
                g.ColumnDefinitions.Add(cd)
            for r, row in enumerate([head] + rows):
                _add_row_def(g)
                for c, text in enumerate(row):
                    tb = TextBlock()
                    tb.Text = text
                    tb.FontSize = 11
                    tb.Margin = Thickness(0, 0, 16, 0)
                    if c > 0:
                        tb.HorizontalAlignment = HorizontalAlignment.Right
                    if r == 0:
                        tb.FontWeight = _FW.SemiBold
                    _put(g, tb, c, r)
            body.Children.Add(g)
        dex.Expanded += _fill_d
        pnl_dtab.Children.Add(dex)

    comp_boxes = {}
    comp_order = []
    _fill_value_grid(
        grid_comp,
        [(k, lbl, 'in. wc', v) for k, lbl, v in fitting_tables.COMPONENT_TABLE
         if k not in ('backdraft', 'cap')],    # in the Exhaust fans section
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

    # ── VAV box and static fields under the external static option (unit
    #    max ESP may be blank); the MP limits are the Main duct limits row ───
    for r, (key, label, unit, default, _kind) in enumerate(vav_checks.VAV_STATIC_FIELDS):
        _add_row_def(grid_vav_static)
        lb = TextBlock()
        lb.Text = label
        lb.Margin = Thickness(0, 2, 12, 2)
        lb.VerticalAlignment = VerticalAlignment.Center
        _put(grid_vav_static, lb, 0, r)
        tb = _box(vav_checks.format_default(default), 64, Thickness(0, 2, 8, 2))
        _put(grid_vav_static, tb, 1, r)
        vav_boxes[key] = tb
        ex = ui_helpers.rja_hint_block(win, unit)
        ex.VerticalAlignment = VerticalAlignment.Center
        _put(grid_vav_static, ex, 2, r)

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
                            'the safety factor. Fire / smoke and backdraft dampers are '
                            'known drops. An exhaust fan run with no backdraft damper '
                            'modeled pays one when Include backdraft damper is ticked '
                            '(never a ceiling exhaust fan, which has one built in). '
                            'Excludes the coil and cabinet (already in the published '
                            'ESP), the filter unless Include return filter is ticked, '
                            'and an exhaust roof / wall cap unless Include roof / wall '
                            'cap is ticked (an inline fan; an upblast roof fan has none).'),
        ('Fan system effect', '2001 ASHRAE Fundamentals Eq. 38 (p. 34.14). The first '
                              'disruption past each unit connection: within Le '
                              '(Eq. 36 or 37) of an outlet, within the table floor '
                              'distance (5 D or 5 H) of an inlet. Outlet elbow: SR7 '
                              'tables for the fan type, worst of positions A to D, at '
                              'Ab/Ao and L/Le, elbow priced separately. Inlet elbow: '
                              'ED7-2 round, ER7-1 rect (SWSI tables, an assumption for '
                              'DWDI), at r/D and L/D, elbow included. Velocity: the '
                              'larger of the default fan velocity and the connection '
                              'velocity, or the connection alone. Added before the '
                              'safety factor. Transitions are priced as fittings, '
                              'taps flagged, not priced. An elbow under 0.5 Le of an '
                              'outlet, or under 2 D of an inlet, is red with the '
                              'straight duct before it.'),
    ] + _dialog_vav_rows()
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
    filter_state = {'was_box': False, 'tick': None}

    def sync_enable(s=None, e=None):
        master = bool(cb_static.IsChecked)
        # VAV Box: external static from the box outlet only (Colin,
        # 2026-11-09). No return path, so the return filter is greyed out
        # and off, its tick given back on leaving VAV Box; no fan system
        # effect at a box, so its box is greyed out too.
        box = bool(rb_scope_vavbox.IsChecked)
        if box and not filter_state['was_box']:
            filter_state['was_box'] = True
            filter_state['tick'] = bool(cb_filter.IsChecked)
            cb_filter.IsChecked = False
        elif not box and filter_state['was_box']:
            filter_state['was_box'] = False
            if filter_state['tick'] is not None:
                cb_filter.IsChecked = filter_state['tick']
        cb_filter.IsEnabled = not box
        cb_fse.IsEnabled = not box
        pnl_static.IsEnabled = master
        pnl_static.Opacity = 1.0 if master else 0.55
        tb_filter.IsEnabled = bool(cb_filter.IsChecked) and not box
        fse_on = master and bool(cb_fse.IsChecked) and not box
        pnl_fse.IsEnabled = fse_on
        pnl_fse.Opacity = 1.0 if fse_on else 0.55
        # VAV box and static fields appear only with VAV and external static
        # both ticked (Colin, 2026-10-22).
        both = master and bool(rb_scope_vav.IsChecked)
        pnl_vav_static.Visibility = Visibility.Visible if both else Visibility.Collapsed
        # Exhaust fans, the same way: Exhaust System and external static
        # both ticked (Colin, 2026-11-02).
        exh = master and bool(rb_scope_exhaust.IsChecked)
        pnl_exh_static.Visibility = Visibility.Visible if exh else Visibility.Collapsed
        # An exhaust fan has no return path, so the return filter is not
        # asked for under Exhaust System (Colin, 2026-11-04).
        pnl_filter.Visibility = (Visibility.Collapsed if bool(rb_scope_exhaust.IsChecked)
                                 else Visibility.Visible)
        # The medium pressure supply row shows with VAV System alone.
        vis = Visibility.Visible if bool(rb_scope_vav.IsChecked) else Visibility.Collapsed
        for el in mp_row:
            el.Visibility = vis
    for cb in (cb_static, cb_filter, cb_fse, rb_scope_rtu, rb_scope_vav, rb_scope_vavbox,
               rb_scope_exhaust):
        cb.Checked += sync_enable
        cb.Unchecked += sync_enable
    sync_enable()

    # Fan type follows the Scope (Colin, 2026-11-03): RTU and VAV System
    # DWDI, Exhaust System SWSI. Set only when the Scope changes, so a manual
    # pick holds until the Scope changes again.
    def _scope_fan_type(s=None, e=None):
        swsi = bool(rb_scope_exhaust.IsChecked)
        rb_fan_swsi.IsChecked = swsi
        rb_fan_dwdi.IsChecked = not swsi
    for rb in (rb_scope_rtu, rb_scope_vav, rb_scope_vavbox, rb_scope_exhaust):
        rb.Checked += _scope_fan_type

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
        # Exhaust System and VAV Box: no return path, the filter box is
        # ignored (VAV Box, Colin, 2026-11-09).
        if (ext_static and bool(cb_filter.IsChecked)
                and not bool(rb_scope_exhaust.IsChecked)
                and not bool(rb_scope_vavbox.IsChecked)):
            f = num(tb_filter, 'Return filter drop', 'nonneg')
            if f is not None:
                filter_inwc = f

        exhaust = bool(rb_scope_exhaust.IsChecked)
        exhaust_settings = {'backdraft': bool(cb_exh_backdraft.IsChecked)}
        if exhaust and ext_static:
            # The two drops price modeled dampers and a ceiling fan's own cap
            # too, so they are read whether or not their box is ticked.
            bd = num(tb_backdraft, 'Backdraft damper drop', 'nonneg')
            if bd is not None:
                comp_values['backdraft'] = bd
            cap = num(tb_exh_cap, 'Roof / wall cap drop', 'nonneg')
            if cap is not None:
                comp_values['cap'] = cap

        fan_effect = fan_system_effect.default_settings()
        fan_effect['on'] = bool(cb_fse.IsChecked)
        fan_effect['fan_type'] = 'SWSI' if bool(rb_fan_swsi.IsChecked) else 'DWDI'
        # The fan velocity, the connection's when higher; no other choice in
        # the dialog (Colin, 2026-11-03).
        fan_effect['velocity_basis'] = 'default'
        if ext_static and fan_effect['on']:
            v, err = fan_system_effect.parse_blast_ratio(tb_blast.Text)
            if err:
                errors.append((tb_blast, 'Blast area ratio Ab/Ao ' + err + '.', ()))
            else:
                fan_effect['blast_ratio'] = v
            if fan_effect['velocity_basis'] == 'default':
                v, err = fan_system_effect.parse_fan_fpm(tb_fan_fpm.Text)
                if err:
                    errors.append((tb_fan_fpm, 'Fan velocity ' + err + '.', ()))
                else:
                    fan_effect['fan_fpm'] = v

        vav = bool(rb_scope_vav.IsChecked)
        vav_settings = vav_checks.default_settings()
        if vav:
            fields = [(f, ()) for f in vav_checks.VAV_LIMIT_FIELDS]
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
                                  calc_basis, filter_inwc, vav, vav_settings,
                                  fan_effect,
                                  ext_static and exhaust and bool(cb_exh_cap.IsChecked),
                                  exhaust, exhaust_settings,
                                  scope=_scope_of(bool(rb_scope_vav.IsChecked),
                                                  bool(rb_scope_vavbox.IsChecked),
                                                  exhaust))
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

    # ── 2. Scope, one unit per run (Colin, 2026-11-03): RTU, VAV System
    # (its fields greyed out until chosen) or Exhaust System ───────────────
    def _scope_radio(text, checked=False):
        rb = RadioButton()
        rb.GroupName = 'scope'
        tb = TextBlock()
        tb.Text = text
        tb.TextWrapping = TextWrapping.Wrap
        tb.Width = CONTENT_W - 20
        rb.Content = tb
        rb.IsChecked = checked
        rb.Margin = Thickness(2, 0, 0, 2)
        outer.Children.Add(rb)
        return rb

    rb_rtu = _scope_radio('RTU.  Lists units only, no VAV boxes or exhaust fans.', True)
    rb_vav = _scope_radio('VAV System.  Lists units only. The full system: supply mains '
                          'between the unit and the box inlets are checked as medium '
                          'pressure (MP).')
    rb_vavbox = _scope_radio('VAV Box.  Lists VAV boxes only. Only the low pressure '
                             'ductwork after the box is checked. External static runs '
                             'from the box outlet only, not the total system loss.')
    # Exhaust System (Colin, 2026-11-02): lists exhaust fans; the backdraft
    # drop is in the component drops below.
    rb_exhaust = _scope_radio('Exhaust System.  Lists exhaust fans only.')

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
        on = bool(rb_vav.IsChecked)
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

    # Return filter as a known drop on the return path total. Colin,
    # 2026-10-08: "Include Return Filter", 0.14 known drop. On by default
    # (Colin, 2026-11-02); untick it when the filter is inside the unit and
    # already deducted from its published ESP.
    filt_panel = StackPanel()
    filt_panel.Orientation = Orientation.Horizontal
    filt_panel.Margin = Thickness(22, 0, 0, 6)
    cb_filter = CheckBox()
    cb_filter_text = TextBlock()
    cb_filter_text.Text = 'Include return filter'
    cb_filter.Content = cb_filter_text
    cb_filter.IsChecked = True
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

    # Exhaust System and VAV Box: no return path, so the return filter is not
    # asked for (Colin, 2026-11-04; VAV Box 2026-11-09).
    def _sync_filter(s=None, e=None):
        on = not bool(rb_exhaust.IsChecked) and not bool(rb_vavbox.IsChecked)
        cb_filter.IsEnabled = on
        tb_filter.IsEnabled = on
    for _rb in (rb_exhaust, rb_vavbox):
        _rb.Checked += _sync_filter
        _rb.Unchecked += _sync_filter
    _sync_filter()

    # The roof or wall cap on an inline exhaust fan's discharge (Colin,
    # 2026-11-02), priced at the Wall / roof cap component drop.
    cb_exh_cap = CheckBox()
    cb_exh_cap_text = TextBlock()
    cb_exh_cap_text.Text = 'Include exhaust roof cap (Wall / roof cap drop on every exhaust run)'
    cb_exh_cap.Content = cb_exh_cap_text
    cb_exh_cap.Margin = Thickness(22, 0, 0, 6)
    outer.Children.Add(cb_exh_cap)

    outer.Children.Add(vav_static_panel)
    rb_vav.Checked += _sync_vav
    rb_vav.Unchecked += _sync_vav
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
    _info_row('Fitting losses:',  u'C × Pv, C per fitting from the 2001 ASHRAE '
                                  u'Fundamentals Ch. 34 tables by its geometry, '
                                  u'RJA SP_LOSS_WORKSHEET as fallback (editable '
                                  u'above), plus one flex bend (C 3.3) per flex run. '
                                  u'Return and exhaust transitions priced in the '
                                  u'airflow direction.')
    _info_row('External static:', 'the single most restrictive run (index run) from '
                                  'the unit to a terminal, supply path + return path')
    _info_row('  includes:',      'duct friction, fittings (elbows, take-offs, '
                                  'transitions), balancing dampers, the diffuser, '
                                  'and the safety factor')
    _info_row('  excludes:',      'filter (unless Include return filter is ticked), coil '
                                  'and cabinet (already in the published ESP), exhaust '
                                  'roof cap (unless Include exhaust roof cap is ticked).')
    for _lab, _val in _dialog_vav_rows():
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
            vav = bool(rb_vav.IsChecked)
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
            if (bool(cb_filter.IsChecked) and not bool(rb_exhaust.IsChecked)
                    and not bool(rb_vavbox.IsChecked)):
                filter_inwc = float(tb_filter.Text)
                if filter_inwc < 0:
                    forms.alert('The return filter drop cannot be negative.',
                                title='Invalid Input')
                    return
            result[0] = _settings_result(out, gpct, selected_cols, full_diag,
                                         ext_static, safety_pct, c_values,
                                         comp_values, calc_basis, filter_inwc,
                                         vav, vav_settings,
                                         # no fan fields here: the fan type
                                         # follows the Scope (Colin, 2026-11-03)
                                         fan_effect=dict(
                                             fan_system_effect.default_settings(),
                                             fan_type='SWSI' if bool(rb_exhaust.IsChecked)
                                             else 'DWDI'),
                                         exhaust_cap=ext_static and bool(cb_exh_cap.IsChecked),
                                         exhaust=bool(rb_exhaust.IsChecked),
                                         scope=_scope_of(bool(rb_vav.IsChecked),
                                                         bool(rb_vavbox.IsChecked),
                                                         bool(rb_exhaust.IsChecked)))
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
# Replaces forms.SelectFromList: one list of the Scope's equipment, one pick.
_PICKER_XAML = '''
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Equipment"
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
            <ListBox x:Name="lst" SelectionMode="Single" MaxHeight="300"
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


_PICKER_PICK_TAG = '__select_unit_in_model__'


def _show_system_picker_xaml(display_names, scope_label, use_theme):
    """Themed picker: one list of the scope's equipment, ONE pick (Colin,
    2026-11-03: one unit per run, no duct pick). The first entry, "Select
    unit in model", closes the list so the unit is clicked in the view.
    Returns ('equipment', [name]), ('pick', None) or None. Raises if the
    window cannot be built (caller falls back)."""
    import ui_helpers
    from System.Windows.Controls import ListBoxItem
    from System.Windows import FontWeights as _FW

    xaml = (_PICKER_XAML
            .replace('@@RESOURCES@@', ui_helpers.rja_window_resources_xaml(use_theme))
            .replace('@@HEADER@@', ui_helpers.rja_header_xaml(
                'Duct Performance', 'Select the unit to analyze.')))
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
    tb_hint.Text = ('Scope: {}. Select one unit found in this view, or Select unit '
                    'in model to click it. One unit per run.'.format(scope_label))

    # First entry: click the unit in the view (Colin, 2026-11-03).
    pick_item = ListBoxItem()
    pick_item.Tag = _PICKER_PICK_TAG
    pick_panel = StackPanel()
    pick_title = TextBlock()
    pick_title.Text = 'Select unit in model'
    pick_title.FontWeight = _FW.SemiBold
    pick_panel.Children.Add(pick_title)
    pick_hint = ui_helpers.rja_hint_block(
        win, 'Click the unit in the view. Only {} equipment can be picked.'.format(
            scope_label))
    pick_hint.Margin = Thickness(0, 2, 0, 0)
    pick_panel.Children.Add(pick_hint)
    pick_item.Content = pick_panel
    lst.Items.Add(pick_item)

    items = []
    for n in display_names:
        it = ListBoxItem()
        it.Content = n
        it.Tag = n
        lst.Items.Add(it)
        items.append(it)
    if items:
        items[0].IsSelected = True

    def on_selection_changed(s, e):
        tb_error.Text = ''
    lst.SelectionChanged += on_selection_changed

    def on_ok(s, e):
        try:
            sel = list(lst.SelectedItems)
            if not sel:
                tb_error.Text = 'Select a unit.'
                return
            if sel[0].Tag == _PICKER_PICK_TAG:
                result[0] = ('pick', None)
            else:
                result[0] = ('equipment', [str(sel[0].Tag)])
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
            if items:
                items[0].Focus()
        except Exception:
            pass

    btn_ok.Click += on_ok
    btn_cancel.Click += on_cancel
    win.Loaded += on_loaded
    win.ShowDialog()
    return result[0]


def show_system_picker(display_names, scope_label):
    """Second screen. Returns ('equipment', [name]), ('pick', None) (click
    the unit in the view) or None.

    display_names is the sorted list of the scope's equipment labels found
    in the active view, never empty (main() stops first when it is).
    scope_label names the Scope choice in the hint. Falls back to the
    original picker if the themed one cannot be built.
    """
    log = script.get_logger()
    for use_theme in (True, False):
        try:
            return _show_system_picker_xaml(display_names, scope_label, use_theme)
        except Exception:
            log.warning('Themed system picker failed (theme=%s), falling back:\n%s'
                        % (use_theme, traceback.format_exc()))
    return _show_system_picker_legacy(display_names, scope_label)


def _show_system_picker_legacy(display_names, scope_label):
    """FALLBACK ONLY. Original picker, one selection, plus a Select unit in
    model button.

    Returns ('equipment', [name]), ('pick', None) or None.
    """
    result = [None]

    win = Window()
    win.Title = 'Select Equipment'
    win.Width = 520
    win.SizeToContent = SizeToContent.Height
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    outer = StackPanel()
    outer.Margin = Thickness(14)

    head = TextBlock()
    head.Text = ('Scope: {}. Equipment found in the active view. Select one unit; '
                 'one unit per run.'.format(scope_label))
    head.TextWrapping = TextWrapping.Wrap
    head.Margin = Thickness(0, 0, 0, 10)
    outer.Children.Add(head)

    lst = ListBox()
    lst.SelectionMode = SelectionMode.Single
    lst.MaxHeight = 300
    for n in display_names:
        lst.Items.Add(n)
    if display_names:
        lst.SelectedIndex = 0
    outer.Children.Add(lst)

    btn_panel = StackPanel()
    btn_panel.Orientation = Orientation.Horizontal
    btn_panel.HorizontalAlignment = HorizontalAlignment.Right
    btn_panel.Margin = Thickness(0, 14, 0, 0)

    run_btn = Button()
    run_btn.Content = 'Run Duct Performance'
    run_btn.Width   = 128
    run_btn.Margin  = Thickness(0, 0, 8, 0)

    cancel_btn = Button()
    cancel_btn.Content = 'Cancel'
    cancel_btn.Width   = 72

    def on_run(s, e):
        if lst.SelectedItem is None:
            forms.alert('Select a unit.', title='Nothing Selected')
            return
        result[0] = ('equipment', [str(lst.SelectedItem)])
        win.Close()

    def on_cancel(s, e):
        win.Close()

    pick_btn = Button()
    pick_btn.Content = 'Select unit in model'
    pick_btn.Width   = 140
    pick_btn.Margin  = Thickness(0, 0, 8, 0)

    def on_pick(s, e):
        result[0] = ('pick', None)
        win.Close()

    run_btn.Click    += on_run
    pick_btn.Click   += on_pick
    cancel_btn.Click += on_cancel
    btn_panel.Children.Add(pick_btn)
    btn_panel.Children.Add(run_btn)
    btn_panel.Children.Add(cancel_btn)
    outer.Children.Add(btn_panel)

    win.Content = outer
    win.ShowDialog()
    return result[0]


# ── helpers ────────────────────────────────────────────────────────────────────
# Scope (Colin, 2026-11-03): one choice, one unit per run. Each lists only
# its kinds of equipment; a VAV box can still be picked alone under VAV System.
_SCOPE_LABEL = {'rtu': 'RTU', 'vav': 'VAV System', 'vavbox': 'VAV Box',
                'exhaust': 'Exhaust System'}
# Exhaust fan type as read from its connections (Colin, 2026-11-03).
_EXHAUST_FAN_KIND_TEXT = {
    'ceiling': 'Exhaust fan type: ceiling (outlet only): discharge run to its cap, '
               'outlet fan system effect.',
    'inline': 'Exhaust fan type: inline (inlet and outlet): worst intake path plus the '
              'whole discharge once, inlet and outlet fan system effect.',
    'roof': 'Exhaust fan type: upblast or downblast roof fan (inlet only): worst intake '
            'path, inlet fan system effect, no discharge walked.',
}
# VAV System is the full system from the unit, VAV Box one box's low pressure
# supply (Colin, 2026-11-03).
_SCOPE_LISTS = {'rtu': ('unit',), 'vav': ('unit',), 'vavbox': ('VAV box',),
                'exhaust': ('exhaust fan',)}
_KIND_SCOPE = {'unit': 'RTU or VAV System', 'VAV box': 'VAV Box',
               'exhaust fan': 'Exhaust System'}


def _equip_kind(eq):
    """'exhaust fan', 'VAV box' or 'unit': how the Scope lists equipment."""
    if hvac_graph.is_exhaust_fan(eq):
        return 'exhaust fan'
    if hvac_graph.is_fed_from_upstream_supply(eq):
        return 'VAV box'
    return 'unit'


class _UnitPickFilter(ISelectionFilter):
    """Only mechanical equipment the Scope lists can be clicked."""

    def __init__(self, scope):
        self.kinds = _SCOPE_LISTS.get(scope, ('unit',))

    def AllowElement(self, elem):
        try:
            return (hvac_graph.is_equipment(elem)
                    and _equip_kind(elem) in self.kinds)
        except Exception:
            return False

    def AllowReference(self, ref, point):
        return False


def _pick_unit(scope, scope_label):
    """Click one unit in the view (Colin, 2026-11-03). Returns the element,
    or None when Esc is pressed. The filter keeps it to the Scope's
    equipment; unavailable, the pick is unfiltered and checked after."""
    try:
        filt = _UnitPickFilter(scope)
    except Exception:
        filt = None
    prompt = 'Select the {} unit to analyze'.format(scope_label)
    try:
        if filt is not None:
            ref = uidoc.Selection.PickObject(ObjectType.Element, filt, prompt)
        else:
            ref = uidoc.Selection.PickObject(ObjectType.Element, prompt)
    except Exception:
        return None
    elem = doc.GetElement(ref.ElementId)
    if elem is None or not hvac_graph.is_equipment(elem):
        output.print_md('The picked element is not mechanical equipment.')
        return None
    if _equip_kind(elem) not in _SCOPE_LISTS.get(scope, ('unit',)):
        output.print_md('{} is a {}, not listed under the {} Scope.'.format(
            _elem_name(elem), _equip_kind(elem), scope_label))
        return None
    output.print_md('Unit picked: **{}** (id {})'.format(_elem_name(elem),
                                                        eid_int(elem.Id)))
    return elem

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


def _emit_corrections(corrections, headers=(), rows=(), status_col=None,
                      align=None, widths=None, subtitle=None, rows_text=None):
    """Print the System Corrections section: the duct table (rows may be
    empty) and every gathered correction line, Action required then Notes.
    rows_text: the padded monospace duct table lines for the plain text
    fallback, or None."""
    rows = list(rows)

    def _text():
        output.print_md('### System Corrections')
        if rows_text:
            output.print_code('\n'.join(rows_text))
        for line in duct_perf_console.corrections_text(corrections, bool(rows)):
            output.print_md(line)

    _emit_html(duct_perf_console.system_corrections(
        list(headers), rows, corrections, status_col=status_col, align=align,
        widths=widths, subtitle=subtitle), _text)


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
                        vav_box_ids=None, vav_box_drop_inwc=None,
                        fan_effect=None, discharge_roots=None, exhaust_cap=False,
                        backdraft_roots=None, inline_fan_roots=None):
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
    Duct Performance no longer offers a duct pick (Colin, 2026-11-03), so main()
    passes an empty set; the case is kept for the walker and its tests.
    Fire / smoke and backdraft dampers are priced at their OWN known drops, never
    the balancing damper figure; any other accessory on the run is reported as
    uncounted. exhaust_cap: every Exhaust Air index run not from a ceiling
    exhaust fan also pays the 'cap' drop once (an inline fan's roof cap).
    inline_fan_roots: inline exhaust fan roots (Colin, 2026-11-03). The
    connection reaching the grilles is the inlet, the other the discharge; the
    discharge's friction, fittings, accessories and outlet fan system effect,
    at the intake cfm, are added to every Exhaust Air run of the fan.
    backdraft_roots: exhaust fan roots whose Exhaust Air run pays one
    'backdraft' drop when no backdraft damper is modeled on it (Include
    backdraft damper; a ceiling fan's root is left out, its damper is built in).

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

    Fan system effect (Colin, 2026-10-29; lib/fan_system_effect.py):
    fan_effect is the dialog's {'on', 'fan_type', 'blast_ratio'}, None or
    off for none. For each unit root, each unit connection is walked to its
    first disruption inside its window (outlet: Le; inlet: the table's floor
    distance, 5 D or 5 H). An elbow is priced (supply outlet: SR7, worst of
    positions A to D, elbow also priced as a fitting; return or exhaust
    inlet: ED7-2 / ER7-1, which include the elbow, so the elbow is not
    priced again) and added to the subtotal BEFORE the safety factor, like
    any fitting; a transition, tap or other fitting inside the window is
    flagged, not priced.

    The walk leaves a unit at the unit connection's velocity (first duct cfm
    over the connector area) and the first duct's class, so a fitting right
    on the unit is priced; it started at 0 fpm before 2026-10-29. Each leaf carries 'fse_inwc' (0.0 when none) and
    'fse_notes' [(kind, text)], kind 'note' or 'warn'. A duct-rooted pick
    takes the picked duct's unit side end as the unit outlet or inlet (Colin,
    2026-11-02), with a note naming the duct.
    """
    dr_by_int_id = {}
    for dr in all_duct_results.values():
        dr_by_int_id[dr.element_id] = dr

    # Ceiling exhaust fans (Colin, 2026-11-02): {fan root id: cap node id}.
    # The run ends at the wall or roof cap, priced as the 'cap' component in
    # place of a diffuser, and the fan system effect is at the fan OUTLET.
    discharge_roots = discharge_roots or {}
    inline_fan_roots = set(inline_fan_roots or ())
    inline_cache = {}       # root id -> _inline_sides(root id)
    fan_kinds = {}          # root id -> 'inline', 'roof' or None (exhaust fans)
    cap_ids = set(discharge_roots.values())

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

    cap_cache = {}
    def _capped(nid):
        if nid not in cap_cache:
            cap_cache[nid] = hvac_graph.duct_is_capped(nid, all_nodes, all_children)
        return cap_cache[nid]

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
    tapafter_cache = {}     # duct id -> {tap id: main cfm left past its station}
    _TAP_STATION_FT = 0.1   # taps this close share one station (F2, 2026-11-04)
    def _tap_flows(duct_id):
        """({tap id: (local common cfm, branch cfm, branch area, feet)}, ordered).

        On an unbroken Revit main the duct's cfm is the flow at its unit end;
        each tap sees that less the branches taken off ahead of it, so taps are
        put in order along the duct (hvac_graph.tap_positions_ft). When the
        order cannot be read every tap sees the duct's own cfm (NOTE).

        Taps within _TAP_STATION_FT of each other are one station (Colin,
        2026-11-04, F2): each sees the flow arriving there and is priced from
        its own geometry and flow, feet is the station's, and the main loses
        the whole station's branch flow past it. Ties sort by element id, so
        Revit's connector order can never change the answer."""
        if duct_id in tapflow_cache:
            return tapflow_cache[duct_id]
        dr = dr_by_int_id[duct_id]
        taps = _taps(duct_id)
        pos = {}
        elem = all_nodes.get(duct_id)
        par = parent_of.get(duct_id)
        if par is None and duct_rooted_ids and duct_id in duct_rooted_ids and taps:
            # A picked duct has no parent in the walk: its taps are measured
            # from its unit side neighbor, the end the fan system effect walk
            # also uses (F2, Colin 2026-11-02).
            side = _unit_side_firsts(duct_id)
            par = side[0][0] if side else None
            fan_system_effect.log.info('tap order: picked duct %s measured from %s',
                                       duct_id, par if par is not None else 'NOT FOUND')
        if elem is not None and par is not None and taps:
            try:
                pos = hvac_graph.tap_positions_ft(elem, par)
            except Exception:
                pos = {}
        ordered = bool(taps) and all(t in pos for t in taps)
        order = (sorted(taps, key=lambda t: (pos[t], t)) if ordered
                 else sorted(taps))
        stations = []
        for t in order:
            if (ordered and stations and
                    pos[t] - pos[stations[-1][0]] <= _TAP_STATION_FT + 1e-9):
                stations[-1].append(t)
            else:
                stations.append([t])
        q = dr.cfm
        out, after = {}, {}
        for group in stations:
            q_in, taken = q, 0.0
            for t in group:
                qb, ab = _branch_flow(t)
                out[t] = (q_in, qb, ab, pos.get(group[0]))
                taken += qb or 0.0
            if ordered:
                q = max(q_in - taken, 0.0)
            for t in group:
                after[t] = q
            if len(group) > 1:
                fan_system_effect.log.info('tap station: duct %s taps %r at %s ft share '
                                           '%.0f cfm', duct_id, group,
                                           pos.get(group[0]), q_in)
        tapflow_cache[duct_id] = (out, ordered)
        tapafter_cache[duct_id] = after
        return tapflow_cache[duct_id]

    def _floor_return(loss, sys_class, what):
        """(loss, notes). ASHRAE prints negative converging junction C for
        the slower stream, a credit. On return and exhaust a tap loss is
        floored at zero for external static, a conservative estimate (Colin,
        2026-10-23); the report names each one floored."""
        if loss < 0.0 and not vav_checks.is_supply_class(sys_class):
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
                # A main that ends in a tap is a turn into the branch, not a
                # junction (Colin, 2026-10-26): all of its air turns (Qb/Qc at
                # 1.0), so it is priced as a 90 degree elbow at the main's own
                # Pv. Round main: die stamped r/D 1.5, the company standard.
                # Rect main: smooth radius elbow, r/W 1.5, H/W as drawn.
                # Unreadable size falls through to the tap table below.
                # A CAPPED main runs on past its last tap to the cap: that
                # tap is a real side tap, priced from the tap table (Colin,
                # 2026-11-05, TB-M10C).
                if (min(qb / qc, 1.0) >= 0.99 and not _continues(main)
                        and not _capped(main)):
                    ec, elabel, enotes = None, None, []
                    if _is_round_duct(dr):
                        _d = _round_d_in(dr)
                        if _d:
                            ec, elabel, enotes = ashrae_fittings.round_elbow_c(
                                _d, 1.5, 90.0)
                    else:
                        _wh = re.match(r'\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)',
                                       getattr(dr, 'size', '') or '')
                        if _wh and float(_wh.group(1)) > 0:
                            ec, elabel, enotes = ashrae_fittings.rect_smooth_elbow_c(
                                1.5, float(_wh.group(2)) / float(_wh.group(1)), 90.0)
                    if ec is not None:
                        vc = qc / dr.area_ft2
                        res = (ec * hvac_graph.velocity_pressure_inwg(vc),
                               '%s, end of main turn, C %.2f at main %.0f fpm'
                               % (elabel, ec, vc),
                               _located('tap %s, end of main %s (%s %s)' % (
                                   tap_id, main, dr.sys_class, getattr(dr, 'size', '')), enotes))
                        tapprice_cache[tap_id] = res
                        return res
                cb, label, notes = ashrae_fittings.tap_branch_c(
                    vav_checks.is_supply_class(dr.sys_class),
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
                           _located('tap %s on main %s (%s %s)' % (
                               tap_id, main, dr.sys_class, getattr(dr, 'size', '')), notes) + fl)
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
        # Why the table could not be read, for the FALLBACK line (Colin,
        # 2026-11-02: say why, not just "unreadable").
        if not qb:
            why = 'the branch carries no airflow, see the Flow = 0 warning'
        elif not qc:
            why = 'main airflow unreadable'
        elif not dr.area_ft2 or elem is None:
            why = 'main size or tap unreadable'
        else:
            why = 'outside the table'
        if qc and qb and dr.area_ft2 and elem is not None:
            fam = hvac_graph.fitting_family_name(elem)
            qs = max(qc - qb, 0.0)
            cs, lbl, _n = ashrae_fittings.tap_main_c(
                vav_checks.is_supply_class(dr.sys_class),
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
                   '(%s)' % (t, duct_id, c_ws, why))
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

    def _remaining_cfm(duct_id):
        """Flow left on an unbroken main past its last tap: the duct's unit
        end cfm less every tap's branch cfm, the same flow _seg_friction ends
        on. None when the duct has no taps or their order cannot be read."""
        dr = dr_by_int_id[duct_id]
        taps = _taps(duct_id)
        if not taps:
            return None
        flows, ordered = _tap_flows(duct_id)
        if not ordered:
            return None
        return max(dr.cfm - sum([(flows[t][1] or 0.0) for t in taps], 0.0), 0.0)

    def _tap_order_note(duct_id):
        """The NOTE for a main whose tap positions cannot be read, or None.
        Without the order, _seg_friction keeps the whole duct at its unit end
        cfm (Colin, 2026-11-02: say so in the run's notes)."""
        dr = dr_by_int_id[duct_id]
        if not _taps(duct_id) or not dr.area_ft2 or _tap_flows(duct_id)[1]:
            return None
        return ('NOTE: tap positions on duct %s unreadable, friction charged over '
                'its whole length at its unit end cfm' % duct_id)

    def _seg_friction(duct_id, child_id):
        """(friction in. wc, feet charged) on an unbroken main with taps, or
        None to keep the duct's own whole length figure.

        The main keeps one size but its flow drops at every tap, so the
        friction is the sum of its pieces, each at its own flow (Colin,
        2026-11-02): unit end to tap 1 at the duct cfm, tap 1 to tap 2 at
        that less tap 1's branch, and so on. A run leaving through a tap is
        charged only up to that tap. Needs the tap order (positions read);
        without it the whole duct is charged at its unit end cfm, as before."""
        dr = dr_by_int_id[duct_id]
        taps = _taps(duct_id)
        if not taps or not dr.area_ft2:
            return None
        flows, ordered = _tap_flows(duct_id)
        if not ordered:
            return None
        length = float(dr.length_ft or 0.0)
        stop = (min(float(flows[child_id][3]), length) if child_id in taps
                else length)
        d_h = getattr(dr, 'd_h_in', None)
        eps = getattr(dr, 'roughness_ft', None)
        q, x0, total = dr.cfm, 0.0, 0.0
        after = tapafter_cache.get(duct_id, {})
        for t in sorted(taps, key=lambda t: (flows[t][3], t)):
            x1 = min(max(float(flows[t][3]), x0), stop)
            if x1 > x0 and q > 0:
                total += (x1 - x0) / 100.0 * hvac_graph.duct_friction_loss_per_100ft(
                    q / dr.area_ft2, d_h, eps)
            x0 = x1
            if x0 >= stop:
                break
            q = after.get(t, max(flows[t][0] - (flows[t][1] or 0.0), 0.0))
        if stop > x0 and q > 0:
            total += (stop - x0) / 100.0 * hvac_graph.duct_friction_loss_per_100ft(
                q / dr.area_ft2, d_h, eps)
        hvac_graph.log.info('_seg_friction: duct %s exit %s, %.2f of %.2f ft, %.5f in. wc '
                            '(whole duct at unit end cfm %.5f)', duct_id, child_id, stop,
                            length, total, dr.friction_loss_inwc)
        return total, stop

    def _shown_class(nid, sys_class):
        """The system name printed for a fitting. A discharge duct carries
        flow_class 'Supply Air' so its fittings read the supply (away from the
        fan) tables (hvac_graph.build_network); the report names the duct's own
        classification, Exhaust Air on an exhaust fan discharge (Colin,
        2026-11-09, TB-M10C transition 3696775)."""
        if nid is None:
            return sys_class
        for d in vav_checks.nearest_ducts_down(nid, all_children, _kind):
            dr = dr_by_int_id.get(d)
            if (dr is not None and getattr(dr, 'flow_class', None) == sys_class
                    and dr.sys_class):
                return dr.sys_class
        return sys_class

    def _ashrae_fitting(elem, role, fam, sys_class, up_fpm, up_area, loc_id=None):
        """(loss in. wc, label, notes, downstream area) from the 2001 tables,
        or None to fall back to the worksheet. sys_class picks the table; the
        printed name is _shown_class."""
        eid = loc_id if loc_id is not None else '?'
        shown = _shown_class(loc_id, sys_class)
        if role == 'elbow':
            g = hvac_graph.elbow_geometry(elem)
            if g is None:
                return None
            where = ('elbow %s (%s, %s, %.0f deg)'
                     % (eid, shown, _size_text(g['round'], g['d_in'], g['w_in'],
                                                   g['h_in']), g['angle'] or 90.0))
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
                    '%s, C %.2f' % (label, c), _located(where, notes), None)
        if role == 'transition':
            ends, theta = hvac_graph.transition_geometry(elem)
            if not ends or theta is None or not up_area or not up_fpm:
                return None
            # Section 1 is the walk-upstream (unit side) end and o the far end,
            # for supply (air 1 -> o) and return (air o -> 1, 1 toward the fan)
            # alike, so the walk order IS the table order.
            e1, eo = sorted(ends, key=lambda e: abs(e[0] - up_area))
            c, label, notes = ashrae_fittings.transition_c(
                vav_checks.is_supply_class(sys_class), e1[1], eo[1], e1[0], eo[0], theta)
            if c is None:
                return None
            vo = up_fpm * up_area / eo[0]
            where = ('transition %s (%s, %s to %s, %.1f deg)'
                     % (eid, shown,
                        _size_text(e1[1] == 'round', e1[2], e1[2], e1[3]),
                        _size_text(eo[1] == 'round', eo[2], eo[2], eo[3]), theta))
            return (c * hvac_graph.velocity_pressure_inwg(vo),
                    '%s, Co %.2f at %.0f fpm' % (label, c, vo),
                    _located(where, notes), eo[0])
        return None

    def _terminal_label(nid):
        """'D-6 (id 3639741)': the terminal's Mark and id, so the end of the
        index run can be found on the plan. Id alone when there is no Mark."""
        mark = None
        try:
            p = all_nodes.get(nid).LookupParameter('Mark')
            mark = p.AsString() if p is not None else None
        except Exception:
            pass
        return '%s (id %s)' % (mark, nid) if mark else 'id %s' % nid

    def _size_text(is_round, d_in, w_in, h_in):
        """'16x12"' or '10"dia' for a report note."""
        if is_round:
            return '%.0f"dia' % float(d_in or 0)
        return '%.0fx%.0f"' % (float(w_in or 0), float(h_in or 0))

    def _located(where, notes):
        """Lookup notes led by the fitting they belong to, so the reader can
        find it: 'transition 3639637 (Supply Air, 18x12" to 16x12", 7.6 deg):
        SR4-1: angle 7.6 deg is below ...'."""
        return ['%s: %s' % (where, n) for n in (notes or []) if n]

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
                       nid=None, loc_id=None):
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
            a = _ashrae_fitting(elem, role, fam, sys_class, up_fpm, up_area,
                                loc_id=nid if loc_id is None else loc_id)
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

    def _unit_side_firsts(root_id):
        """[(first node, nodes in its subtree)] for the children of a picked
        duct root that lead back toward the unit: subtrees with no terminal
        that do not start at a take-off. Shared by _unit_side_chain and the
        duct pick fan system effect, so the two agree on the unit side."""
        out = []
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
            out.append((first, order))
        return out

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
        for _first, order in _unit_side_firsts(root_id):
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
                            root_dr.sys_class), loc_id=n)
                    fit   += loss
                    n_fit += counted
                    if note:
                        unpriced = unpriced + (note,)
                    if down_a:
                        up_area = down_a
        return (fric, fit, n_fit, n_duct, dlen, unpriced)

    # ── unit connections (Colin, 2026-10-29) ────────────────────────────────
    # The connector on the unit joined to each of its first nodes: its size
    # sets the velocity the walk starts with and the fan system effect's Ao.
    unit_conn_cache = {}
    def _unit_conn(root_id, first):
        """hvac_graph.equipment_connection_geometry for the unit root_id
        toward node first, or None (not equipment, or unreadable: logged)."""
        key = (root_id, first)
        if key not in unit_conn_cache:
            root_elem = all_nodes.get(root_id)
            geo = None
            if root_elem is not None and hvac_graph.is_equipment(root_elem):
                geo = hvac_graph.equipment_connection_geometry(root_elem, first)
            unit_conn_cache[key] = geo
        return unit_conn_cache[key]

    def _connection_cfm(first):
        """Air through the unit connection at node first: a duct's own cfm,
        else the sum over the nearest ducts below it. nearest_ducts_down
        reaches each duct once (it keeps a seen set and stops at the first
        duct on every path), so a tee right on the unit counts both legs once
        each (review, 2026-10-29). 0.0 when no duct below carries flow."""
        dr = dr_by_int_id.get(first)
        if dr is not None:
            return float(dr.cfm or 0.0)
        below = vav_checks.nearest_ducts_down(first, all_children, _kind)
        return sum([float(dr_by_int_id[d].cfm or 0.0)
                    for d in below if d in dr_by_int_id], 0.0)

    def _unit_start(root_id, first):
        """(fpm, area ft2, sys_class) the walk carries from the unit into node
        first, or None.

        A fitting right on the unit connection used to be priced at 0 fpm,
        because the walk starts with no upstream velocity (Colin, 2026-10-29:
        fix it). The flow is the whole connection's (_connection_cfm, both
        legs of a tee on the unit), the area the unit connector's (the
        section the air passes at the unit, and the end a transition there is
        oriented against), else the first duct's own.
        The class is the first duct's, so a fitting before any duct is priced
        as supply or return correctly."""
        root_elem = all_nodes.get(root_id)
        if root_elem is None or not hvac_graph.is_equipment(root_elem):
            return None
        flow = _first_duct(first)
        cfm = _connection_cfm(first)
        if flow is None or cfm <= 0:
            return None
        conn = _unit_conn(root_id, first)
        area = conn['area_ft2'] if conn is not None and conn.get('area_ft2') else flow.area_ft2
        if not area:
            return None
        return cfm / area, area, getattr(flow, 'flow_class', None) or flow.sys_class

    # ── fan system effect (Colin, 2026-10-29) ───────────────────────────────
    # The 2001 Fundamentals Eq. 38 (p. 34.14) adds FSE2 at the fan outlet and
    # FSE1 at the inlet. Priced from the first disruption inside each unit
    # connection's window; the math is fan_system_effect's.
    # The straight duct and swirl checks run every time; the fan system
    # effect is priced only when ticked (Colin, 2026-11-03, Q4). Unticked,
    # each connection keeps its RED lines and drops the rest (_checks_only).
    fse_price = bool(fan_effect and fan_effect.get('on'))
    fse_set = dict(fan_effect) if fan_effect else fan_system_effect.default_settings()
    _FSE_KIND_TEXT = {'takeoff': 'tap', 'split': 'branch', 'unknown': 'unrecognised fitting'}

    def _transition_length_ft(elem, nid):
        """A transition's own length in feet, the distance between its two
        connector origins (hvac_graph.transition_geometry). 0.0 when it
        cannot be read, logged."""
        ends, _theta = hvac_graph.transition_geometry(elem)
        try:
            length = float(ends[0][4].DistanceTo(ends[1][4]))
        except Exception as ex:
            fan_system_effect.log.info('fan system effect: transition %s length '
                                       'NOT FOUND (%s), 0.0 ft counted', nid, ex)
            return 0.0
        fan_system_effect.log.info('fan system effect: transition %s length %.2f ft',
                                   nid, length)
        return length

    def _fse_first_disruption(root_id, first, win, where='unit connection',
                              prev=None, seen=None, collect=None, passed=None,
                              through_taps=False, ducts=None, accessories=None):
        """(role or None, L ft, node id, notes): the first disruption within
        win feet of the unit connection at node first.

        L is the straight duct from the unit connection to it: duct lengths
        and transition lengths summed (Colin, 2026-11-02: a curb adapter is
        part of the straight length), unions and accessories passed at zero
        length. Every transition is passed; the first elbow is the one the
        system effect prices. A tap on an unbroken duct sits at its own
        distance along that duct (hvac_graph.tap_positions_ft from the unit
        end). role is a fitting_tables.classify_fitting role, 'takeoff' for
        a tap, 'split' for a duct that divides; None when nothing is within
        the window or the run ends inside it (a terminal or equipment).

        prev, seen: a duct pick starts at the picked duct itself (first ==
        root_id), prev its unit side neighbor (None when there is none) and
        seen the unit side nodes, which are not walked. A unit root leaves
        both None: prev and seen are the unit.

        collect: a list for the swirl check. Each elbow inside win is
        appended as (id, L ft) and the walk goes on past it, to the end of
        the window or the first disruption that is not an elbow.

        passed: a dict filled with {transition id: L ft} for every
        transition passed, for the fan velocity basis (Colin, 2026-11-03).

        through_taps: walk past taps on the main instead of stopping at
        them, for the straight duct check.

        ducts: a list filled with (duct id, L ft at its unit end) for every
        duct passed, so the straight duct before a RED elbow is RED too.

        accessories: a list filled with (id, L ft, family) for every damper
        passed (fan_system_effect.DAMPER_KINDS), for the damper in the fan
        clearance WARNING (Colin, 2026-11-04, F1). Still passed at zero
        length."""
        notes = []
        if seen is None:
            prev, seen = root_id, set([root_id])
        nid, l_ft = first, 0.0
        while nid is not None and nid not in seen and l_ft < win:
            seen.add(nid)
            dr = dr_by_int_id.get(nid)
            elem = all_nodes.get(nid)
            kids = [c for c in all_children.get(nid, []) if c not in seen]
            if dr is not None:
                taps = _taps(nid)
                if taps:
                    pos = {}
                    if elem is not None and prev is not None:
                        try:
                            pos = hvac_graph.tap_positions_ft(elem, prev)
                        except Exception:
                            pos = {}
                    if all(t in pos for t in taps):
                        tid = min(taps, key=lambda t: pos[t])
                        at = pos[tid]
                    elif prev is None:
                        tid, at = taps[0], 0.0
                        notes.append(('note', 'NOTE: duct pick, nothing found on the '
                                      'unit side of duct %s, so its taps cannot be put '
                                      'in order from the unit end; tap %s taken at that '
                                      'end for the fan system effect' % (nid, tid)))
                    else:
                        tid, at = taps[0], 0.0
                        notes.append(('note', 'NOTE: tap position on duct %s unreadable, '
                                      'taken at its unit end for the fan system effect'
                                      % nid))
                    fan_system_effect.log.info('fan system effect: tap %s on duct %s '
                                               'at %.2f ft', tid, nid, l_ft + at)
                    if through_taps:
                        pass                # the main goes on past the tap
                    elif l_ft + at < win:
                        return 'takeoff', l_ft + at, tid, notes
                    else:
                        return None, l_ft + at, None, notes
                if ducts is not None:
                    ducts.append((nid, l_ft))
                l_ft += float(dr.length_ft or 0.0)
                nxt = [c for c in kids if c not in taps]
            elif elem is not None and hvac_graph.is_fitting(elem):
                role = fitting_tables.classify_fitting(hvac_graph.fitting_family_name(elem))
                if role == 'transition':
                    # Passed, its own length counted in L; the main walk
                    # prices it as an ordinary fitting.
                    fan_system_effect.log.info('fan system effect: transition %s at '
                                               '%.2f ft passed', nid, l_ft)
                    if passed is not None:
                        passed[nid] = l_ft
                    l_ft += _transition_length_ft(elem, nid)
                elif collect is not None and role == 'elbow':
                    collect.append((nid, l_ft))
                elif role not in fitting_tables.ZERO_LOSS_ROLES:
                    fan_system_effect.log.info('fan system effect: %s %s at %.2f ft',
                                               role, nid, l_ft)
                    return role, l_ft, nid, notes
                nxt = kids
            elif elem is not None and hvac_graph.is_accessory(elem):
                if accessories is not None:
                    fam = hvac_graph.fitting_family_name(elem)
                    akey = fitting_tables.classify_accessory(fam)
                    fan_system_effect.log.info('fan system effect: accessory %s (%s, %s) '
                                               'at %.2f ft', nid, fam, akey, l_ft)
                    if akey in fan_system_effect.DAMPER_KINDS:
                        accessories.append((nid, l_ft, fam))
                nxt = kids
            else:
                return None, l_ft, None, notes      # a terminal or equipment
            if len(nxt) > 1:
                return 'split', l_ft, nid, notes
            if not nxt:
                return None, l_ft, None, notes
            prev, nid = nid, nxt[0]
        fan_system_effect.log.info('fan system effect: nothing within %.1f ft of '
                                   'connection %s (%.1f ft walked)', win, first, l_ft)
        return None, l_ft, None, notes

    def _fse_conn_size(conn, flow, notes, what):
        """(is_round, D or H in.) of an inlet from the unit connector, its
        LARGER side if rectangular (the higher C, and the longer window),
        else the first duct; each fallback is named. (None, None) if nothing
        is readable."""
        if conn is not None:
            if conn.get('round') and conn.get('d_in'):
                return True, conn['d_in']
            if conn.get('w_in') and conn.get('h_in'):
                return False, max(conn['w_in'], conn['h_in'])
        d = _round_d_in(flow)
        if d:
            notes.append(('note', 'NOTE: unit inlet unreadable; %s D taken from the '
                          'first duct (%.0f in.)' % (what, d)))
            return True, d
        if flow.area_ft2:
            side = math.sqrt(flow.area_ft2) * 12.0
            notes.append(('note', 'NOTE: unit inlet unreadable; %s H taken as the side '
                          'of a square of the first duct\'s area (%.0f in.)' % (what, side)))
            return False, side
        return None, None

    def _fse_window_size(conn, flow, notes):
        """(is_round, D in.) for the 5 D check and the inlet window: the
        LARGER of the unit opening (_fse_conn_size) and the first duct, each
        a round D or a rect larger side (Colin, 2026-11-03). is_round is the
        opening's. (None, None) if nothing is readable."""
        is_round, size = _fse_conn_size(conn, flow, notes, 'window')
        if size is None:
            return None, None
        dc = _duct_conn(flow)
        d_duct = None
        if dc is not None:
            if dc.get('d_in'):
                d_duct = dc['d_in']
            elif dc.get('w_in') and dc.get('h_in'):
                d_duct = max(dc['w_in'], dc['h_in'])
        fan_system_effect.log.info('fan system effect: 5 D size, opening %.1f in., '
                                   'first duct %s in.', size,
                                   '%.1f' % d_duct if d_duct else 'NOT FOUND')
        if d_duct and d_duct > size:
            return bool(dc.get('d_in')), d_duct
        return is_round, size

    def _fse_turn_plane_h(eid, conn):
        """H in. for the 2 H straight duct check at a rect unit inlet, elbow
        eid: its side in the plane of its turn (hvac_graph.elbow_geometry),
        or its D when it is round (past a transition), and not less than the
        unit opening's smaller side, since the opening's own side in that
        plane is not read (Colin, 2026-11-04). None when the elbow is
        unreadable (the caller keeps the larger side)."""
        g = hvac_graph.elbow_geometry(all_nodes.get(eid))
        h = None
        if g is not None:
            h = g.get('d_in') if g.get('round') else g.get('w_in')
        if not h:
            fan_system_effect.log.info('turn plane H: elbow %s NOT FOUND, larger side kept', eid)
            return None
        if conn is not None and conn.get('w_in') and conn.get('h_in'):
            h = max(h, min(conn['w_in'], conn['h_in']))
        fan_system_effect.log.info('turn plane H: elbow %s, %.1f in.', eid, h)
        return h

    def _fse_inlet_size(elem, conn, flow, notes, did=None, l_ft=0.0, where='inlet'):
        """(is_round, D or H in., r in. or None) for the inlet elbow. H is
        the side in the plane of the turn (hvac_graph.elbow_geometry swaps
        Width and Height to put it there). Unreadable elbow: the inlet
        connection's size (larger side, conservative), radius unknown.
        A readable rect mitered elbow (family name, as _ashrae_fitting reads
        it) has no radius and is read on the ER7-1 r/H 1.0 row (Colin,
        2026-11-02), not the conservative unreadable radius row."""
        g = hvac_graph.elbow_geometry(elem) if elem is not None else None
        if g is not None and g.get('round') and g.get('d_in'):
            return True, g['d_in'], g.get('r_in')
        if g is not None and not g.get('round') and g.get('w_in'):
            # ER7-1 is printed for square elbows only (figure, p. 34.60): a
            # non-square one is read on H and stated as an ASSUMPTION line
            # in the report (Colin, 2026-10-29).
            if g.get('h_in') and abs(g['w_in'] - g['h_in']) > 0.01:
                notes.append(('assume', fan_system_effect.ER7_1_SQUARE_ASSUMPTION))
            r_in = g.get('r_in')
            fam = (hvac_graph.fitting_family_name(elem) or '').lower()
            if not r_in and 'miter' in fam and 'radius' not in fam:
                r_in = fan_system_effect.MITERED_R_OVER_H * g['w_in']
                notes.append(('note', 'NOTE: elbow %s, %.1f ft from the %s, is mitered: '
                              'ER7-1 read on the r/H %.1f row (Colin, 2026-11-02).'
                              % (did, float(l_ft), where,
                                 fan_system_effect.MITERED_R_OVER_H)))
            return False, g['w_in'], r_in
        is_round, size = _fse_conn_size(conn, flow, notes, 'elbow')
        if size is not None:
            notes.append(('note', 'NOTE: inlet elbow geometry unreadable; %s %.0f in. '
                          'taken from the unit inlet (%s), radius unknown' % (
                              'D' if is_round else 'H', size,
                              'diameter' if is_round else 'larger side, conservative')))
        return is_round, size, None

    def _duct_conn(dr):
        """A picked duct's own section in the shape of
        hvac_graph.equipment_connection_geometry, for a duct pick's fan
        system effect (Colin, 2026-11-02), or None when its size is
        unreadable (logged)."""
        if dr is None or not dr.area_ft2:
            fan_system_effect.log.info('fan system effect: picked duct size NOT FOUND')
            return None
        d = _round_d_in(dr)
        if d:
            return {'round': True, 'area_ft2': dr.area_ft2, 'd_in': d,
                    'w_in': None, 'h_in': None}
        wh = re.match(r'\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)',
                      getattr(dr, 'size', '') or '')
        if wh is None:
            fan_system_effect.log.info('fan system effect: duct %s size %r unreadable, '
                                       'NOT FOUND', dr.element_id, getattr(dr, 'size', None))
            return {'round': False, 'area_ft2': dr.area_ft2, 'd_in': None,
                    'w_in': None, 'h_in': None}
        return {'round': False, 'area_ft2': dr.area_ft2, 'd_in': None,
                'w_in': float(wh.group(1)), 'h_in': float(wh.group(2))}

    def _fse_connection(root_id, first, tag=False, pick=None, direction=None,
                        cfm_in=None):
        """The fan system effect at one unit connection, or None when it is
        not a fan connection this prices (outside air, transfer).

        tag: the unit has more than one connection of this class, so the
        report names this one by its first element.
        pick: (prev, seen) for a duct pick, where first is the picked duct
        itself and its unit side end is taken as the unit connection: its
        own size and cfm are Ao and the flow, and the window walk starts at
        that end (_fse_first_disruption).
        direction: 'in' or 'out' for an inline exhaust fan connection, from
        its terminals (_inline_sides); cfm_in: the airflow for a discharge,
        which carries no terminal of its own.
        Returns {'loss', 'item' (index run item or None), 'where', 'notes',
        'skip': node ids whose own fitting price the system effect replaces
        (an inlet elbow: ED7-2 / ER7-1 include it), 'gray': elbow ids whose
        loss is HIGH_LOSS_INWC or more, 'red': the two elbow ids that turn
        in different planes within 6 Dh (swirl, not priced) and an elbow
        and the straight duct before an elbow under the straight duct line,
        'red_ducts': {duct id: (table reason, required)} for those ducts,
        'boost': {} (the fan velocity transition boost, retired 2026-11-04),
        'group': the straight duct, transition and first elbow ids charged
        once with the fan system effect (_fse_group_max), 'assumptions':
        report ASSUMPTION lines}."""
        rec = {'loss': 0.0, 'item': None, 'where': 'unit connection', 'notes': [],
               'skip': set(), 'gray': set(), 'red': set(), 'boost': {},
               'red_ducts': {}, 'assumptions': [], 'group': set(),
               # Every fan connection prints its own take-off row, outlet and
               # inlet apart (Colin, 2026-11-05): 'none' is the reason a
               # connection carries no fan system effect, shown at 0.000.
               'where_item': None, 'none': None}
        if _kind(first) == 'equipment':
            rec['notes'].append(('note', 'NOTE: unit connection to element %s not '
                                 'checked for fan system effect: it feeds equipment '
                                 'directly.' % first))
            return rec
        side = _side_below(first)
        if side is None:
            rec['notes'].append(('note', 'NOTE: unit connection to element %s not '
                                 'checked for fan system effect: no duct found past '
                                 'it.' % first))
            return rec
        # The role hook (Colin, 2026-10-29): a known connector direction would
        # go in here for exhaust fans, where both sides are Exhaust Air. A
        # unit's role is its classification today, so None is passed.
        role_here = fan_system_effect.connection_role(
            side, 'out' if root_id in discharge_roots else direction)
        if role_here is None:
            fan_system_effect.log.info('fan system effect: connection %s of %s is %r, '
                                       'not checked', first, root_id, side)
            return None
        where_item = fan_system_effect.where_text(side, role_here)
        where = where_item + (' at element %s' % first if tag else '')
        rec['where'] = where
        rec['where_item'] = where_item
        notes = rec['notes']
        flow = _first_duct(first)
        cfm = cfm_in if cfm_in else _connection_cfm(first)
        if flow is None or cfm <= 0 or not flow.area_ft2:
            notes.append(('note', 'NOTE: fan system effect at the %s not checked, no '
                          'airflow on its first duct' % where))
            rec['none'] = 'not checked, no airflow on its first duct'
            return rec
        conn = _duct_conn(flow) if pick is not None else _unit_conn(root_id, first)
        if conn is not None and conn.get('area_ft2'):
            area = conn['area_ft2']
        else:
            area = flow.area_ft2
            notes.append(('note', 'NOTE: unit %s connection size unreadable; the first '
                          'duct (%s) is used for Ao and the velocity pressure'
                          % (where, getattr(flow, 'size', '?'))))
        # Velocity basis (Colin, 2026-11-03): the larger of the default fan
        # velocity and the connection's, or the connection's alone. It sets
        # Pv, the outlet Le and the transitions passed below. Outlets only
        # (Colin, 2026-11-04): an inlet is priced at its own connection
        # velocity, cfm over the connection area.
        v_conn = cfm / area
        if role_here == fan_system_effect.OUTLET:
            v, fan_gov = fan_system_effect.design_velocity(fse_set, v_conn)
        else:
            v, fan_gov = v_conn, False
        if fan_gov:
            notes.append(('note', fan_system_effect.velocity_text(where, v, v_conn)))
        # D for the 5 D check and the inlet window: the larger of the unit
        # opening and the first duct. An outlet already has its fallback NOTE
        # above, so its size notes are not repeated.
        is_round_c, size_c = _fse_window_size(
            conn, flow, notes if role_here == fan_system_effect.INLET else [])
        if role_here == fan_system_effect.OUTLET:
            win, _eq = fan_system_effect.effective_duct_length_ft(area, v)
            win_name = 'Le'
            win_short = 'Le'
        else:
            if size_c is None:
                notes.append(('warn', 'WARNING: %s size unreadable; fan system effect '
                              'not checked' % where))
                rec['none'] = 'not checked, connection size unreadable'
                return rec
            win = fan_system_effect.inlet_window_ft(is_round_c, size_c)
            win_short = '5 %s' % ('D' if is_round_c else 'H')
            win_name = 'the inlet window, ' + win_short
        # The straight duct check (Colin, 2026-11-03, FAN_SYSTEM_EFFECT_PLAN.md
        # section C): an elbow under 0.5 Le of an outlet, or under 2 D of an
        # inlet, is RED with the straight duct before it, still priced below.
        # Elbows only, every one inside the line, walked past taps on the
        # main and past the other elbows, up to a split or another fitting.
        if role_here == fan_system_effect.OUTLET or size_c is not None:
            if role_here == fan_system_effect.OUTLET:
                red_ft = fan_system_effect.outlet_red_ft(win)
            else:
                red_ft = fan_system_effect.inlet_red_ft(size_c)
            elbows_r, ducts_r = [], []
            if pick is not None:
                _fse_first_disruption(root_id, first, red_ft, where, prev=pick[0],
                                      seen=set(pick[1]), collect=elbows_r,
                                      through_taps=True, ducts=ducts_r)
            else:
                _fse_first_disruption(root_id, first, red_ft, where, collect=elbows_r,
                                      through_taps=True, ducts=ducts_r)
            fan_system_effect.log.info('straight duct check: %s, RED line %.2f ft, '
                                       'elbows %r, ducts %r', where, red_ft, elbows_r,
                                       ducts_r)
            for e_r, l_r in elbows_r:
                if role_here == fan_system_effect.OUTLET:
                    notes.append(('red', fan_system_effect.le_text(where, e_r, l_r, win,
                                                                   _eq)))
                    need = fan_system_effect.outlet_red_ft(win)
                else:
                    # The walk used the larger side (conservative); a rect
                    # elbow is then judged on H, its side in the plane of its
                    # turn (Colin, 2026-11-04).
                    h_e = _fse_turn_plane_h(e_r, conn) if not is_round_c else None
                    size_e = h_e if h_e else size_c
                    if l_r >= fan_system_effect.inlet_red_ft(size_e) - 1e-9:
                        fan_system_effect.log.info(
                            'straight duct check: elbow %s at %.2f ft is past 2 H '
                            '(H %.1f in., turn plane), not RED', e_r, l_r, size_e)
                        continue
                    notes.append(('red', fan_system_effect.inlet_red_text(
                        where, e_r, l_r, size_e, is_round_c)))
                    need = fan_system_effect.inlet_red_ft(size_e)
                rec['red'].add(e_r)
                for d_r, at in ducts_r:
                    if at < l_r - 1e-9 and d_r not in rec['red_ducts']:
                        rec['red'].add(d_r)
                        rec['red_ducts'][d_r] = (
                            fan_system_effect.straight_duct_reason(e_r, l_r, role_here),
                            'Straight duct %.1f ft' % fan_system_effect._ceil_step(need))
        passed = {}     # transitions inside the pricing window (win)
        # Two elbows turning in different planes within 6 Dh of the
        # connection: RED, never priced (Colin, 2026-11-02; supply, return
        # and exhaust alike). Walked on its own window, past the first elbow.
        d_h = getattr(flow, 'd_h_in', None)
        if d_h:
            swin = fan_system_effect.swirl_window_ft(d_h)
            elbows = []
            if pick is not None:
                _fse_first_disruption(root_id, first, swin, where, prev=pick[0],
                                      seen=set(pick[1]), collect=elbows)
            else:
                _fse_first_disruption(root_id, first, swin, where, collect=elbows)
            pair = fan_system_effect.swirl_pair(
                [(e, l, hvac_graph.elbow_turn_normal(all_nodes.get(e))) for e, l in elbows])
            fan_system_effect.log.info('swirl check: %s, %d elbows within %.2f ft, pair %r',
                                       where, len(elbows), swin, pair)
            if pair is not None:
                notes.append(('red', fan_system_effect.swirl_text(
                    role_here, where, pair, swin, d_h)))
                rec['red'].update((pair[0][0], pair[1][0]))
            elif len(elbows) >= 2:
                # A double bend (Colin, 2026-11-04): two elbows inside 6 Dh
                # that are not a different-plane pair. Only the first is in
                # the fan system effect, so the real loss is higher. A
                # WARNING, not RED, and nothing extra is priced.
                notes.append(('warn', fan_system_effect.double_elbow_text(
                    where, [(e, l, None) for e, l in elbows], swin, d_h)))
        else:
            fan_system_effect.log.info('swirl check: %s first duct Dh NOT FOUND, '
                                       'not checked', where)
        # A damper inside the pricing window (1.0 Le at an outlet, 5 D / 5 H
        # at an inlet): a WARNING only (Colin, 2026-11-04, F1). Walked past
        # taps and elbows to the end of the window; L and the price below
        # are unchanged.
        acc = []
        if pick is not None:
            _fse_first_disruption(root_id, first, win, where, prev=pick[0],
                                  seen=set(pick[1]), collect=[], through_taps=True,
                                  accessories=acc)
        else:
            _fse_first_disruption(root_id, first, win, where, collect=[],
                                  through_taps=True, accessories=acc)
        fan_system_effect.log.info('damper check: %s, window %.2f ft, dampers %r',
                                   where, win, acc)
        for a_id, a_l, a_fam in acc:
            notes.append(('warn', fan_system_effect.damper_window_text(
                where, a_id, a_fam, a_l, win, win_name)))
        grp_ducts = []
        if pick is not None:
            role, l_ft, did, wnotes = _fse_first_disruption(
                root_id, first, win, where, prev=pick[0], seen=set(pick[1]),
                passed=passed, ducts=grp_ducts)
        else:
            role, l_ft, did, wnotes = _fse_first_disruption(root_id, first, win, where,
                                                            passed=passed, ducts=grp_ducts)
        notes.extend(wnotes)
        if role is None:
            rec['none'] = 'no elbow within %s (%.1f ft)' % (win_short, win)
            return rec
        if role != 'elbow':
            notes.append(('warn', fan_system_effect.disruption_text(
                _FSE_KIND_TEXT.get(role, role), did, where, l_ft, win, win_name)))
            rec['none'] = '%s %s at %.1f ft inside %s, flagged, not priced' % (
                _FSE_KIND_TEXT.get(role, role), did, l_ft, win_short)
            return rec
        if role_here == fan_system_effect.OUTLET:
            r = fan_system_effect.outlet_effect(
                fse_set.get('fan_type'), fse_set.get('blast_ratio'), area, cfm, l_ft,
                pv_of=hvac_graph.velocity_pressure_inwg, v_fpm=v)
        else:
            enotes = []
            is_round, size, r_in = _fse_inlet_size(all_nodes.get(did), conn, flow, enotes,
                                                   did, l_ft, where)
            if size is None:
                notes.append(('warn', 'WARNING: elbow %s, %.1f ft from the %s, but no size '
                              'could be read; fan system effect not priced'
                              % (did, l_ft, where)))
                rec['none'] = 'elbow %s size unreadable, not priced' % did
                return rec
            r = fan_system_effect.inlet_effect(is_round, size, r_in, area, cfm, l_ft,
                                               pv_of=hvac_graph.velocity_pressure_inwg,
                                               v_fpm=v)
            if not r['within']:
                # Past the table's floor distance for its own size: an
                # ordinary elbow, priced by the walk like any other.
                fan_system_effect.log.info('fan system effect: inlet elbow %s at %.2f ft '
                                           'is past %.2f ft, priced as a fitting',
                                           did, l_ft, r['floor_ft'])
                rec['none'] = ('elbow %s at %.1f ft, past %.1f ft, priced as a fitting'
                               % (did, l_ft, r['floor_ft']))
                return rec
            for _k, _t in enotes:
                if _k == 'assume':
                    rec['assumptions'].append(_t)
                else:
                    notes.append((_k, _t))
        loss = r['c'] * hvac_graph.velocity_pressure_inwg(r['v_fpm'])
        # The straight duct, any transition (a curb adapter) and the first
        # elbow are one group, charged once as the larger of the fan system
        # effect and their own losses (Colin, 2026-11-04), outlet and inlet.
        # The walk prices them as usual; each path's end picks the larger.
        rec['group'] = set(d for d, _at in grp_ducts) | set(passed) | set([did])
        # Which one governs is said once the path is known (_fse_group_max,
        # Colin, 2026-11-09), not here.
        rec['gov'] = {'eid': did, 'l_ft': l_ft, 'win': win, 'win_short': win_short}
        notes.extend(('note', 'NOTE: elbow %s, %.1f ft from the %s: %s' % (did, l_ft, where, n))
                     for n in r['notes'])
        if r['high']:
            # GRAY, not RED (Colin, 2026-10-29): still priced in the total,
            # reported with the straight length that brings the loss under
            # HIGH_LOSS_INWC (0.5 in. wc, Colin 2026-11-02, was C 1.0), and
            # the elbow is drawn gray on the plan.
            notes.append(('gray', fan_system_effect.high_text(r, where, did)))
            rec['gray'].add(did)
        name = 'Fan system effect, ' + where_item
        basis = fan_system_effect.basis_text(r)
        rec['loss'] = loss
        rec['item'] = ('Fan system effect {}'.format(did), basis, r['v_fpm'], loss,
                       (name, basis, r['c'], None))
        return rec

    def _checks_only(rec):
        """A connection record with the fan system effect pricing taken out,
        the straight duct and swirl RED kept (fan system effect unticked)."""
        rec['loss'] = 0.0
        rec['item'] = None
        rec['group'] = set()
        rec['skip'] = set()
        rec['gray'] = set()
        rec['boost'] = {}
        rec['assumptions'] = []
        rec['none'] = None
        rec['notes'] = [n for n in rec['notes'] if n[0] == 'red']
        return rec

    def _fse_root(root_id):
        """({unit connection node: record}, notes about the root itself).
        Unticked fan system effect: the RED checks only (_checks_only)."""
        recs, rnotes = _fse_root_all(root_id)
        if not fse_price:
            for _r in recs.values():
                _checks_only(_r)
            rnotes = []
        return recs, rnotes

    def _fse_root_all(root_id):
        recs, rnotes = {}, []
        root_elem = all_nodes.get(root_id)
        if (duct_rooted_ids and root_id in duct_rooted_ids
                and root_id not in dr_by_int_id):
            fan_system_effect.log.info('fan system effect: picked root %s is not a '
                                       'duct, NOT FOUND', root_id)
        elif duct_rooted_ids and root_id in duct_rooted_ids:
            # No unit on the run: the picked duct's unit side end is taken as
            # the unit outlet or inlet (Colin, 2026-11-02). The unit side
            # stretch behind it (_unit_side_firsts) is not walked; its first
            # node is the end the taps on the picked duct are measured from.
            side_firsts = _unit_side_firsts(root_id)
            prev = side_firsts[0][0] if side_firsts else None
            seen = set()
            for _f, _order in side_firsts:
                seen.update(_order)
            fan_system_effect.log.info('fan system effect: duct pick %s, unit side '
                                       'neighbor %s', root_id,
                                       prev if prev is not None else 'NOT FOUND')
            rec = _fse_connection(root_id, root_id, pick=(prev, seen))
            if rec is not None:
                role_here = fan_system_effect.connection_role(_side_below(root_id), None)
                rnotes.append(('note', 'NOTE: duct pick, fan system effect measured from '
                               'the unit end of duct %s, taken as the unit %s.'
                               % (root_id, role_here or 'connection')))
                recs[root_id] = rec
        elif root_id in box_ids:
            rnotes.append(('note', 'NOTE: fan system effect not priced at a VAV box.'))
        elif root_elem is None or not hvac_graph.is_equipment(root_elem):
            fan_system_effect.log.info('fan system effect: root %s is not equipment, '
                                       'NOT FOUND', root_id)
        else:
            # A unit connector also references its logical Duct System
            # element; that is no connection to price (OLMC, 2026-10-30).
            firsts = []
            for f in all_children.get(root_id, []):
                if _kind(f) is None:
                    fan_system_effect.log.info('fan system effect: node %s under %s '
                                               'is not ductwork, skipped', f, root_id)
                else:
                    firsts.append(f)
            classes = [_side_below(f) if _kind(f) != 'equipment' else None
                       for f in firsts]
            dirs, dis_cfm, _n, _kind_fan = inline_cache.get(root_id, ({}, 0.0, [], None))
            # Only the discharge that is priced gets a fan system effect
            # record, so a second one never shows a loss in no total (review,
            # 2026-11-03).
            outs = [f for f in all_children.get(root_id, []) if dirs.get(f) == 'out']
            for first, cls in zip(firsts, classes):
                tag = cls is not None and classes.count(cls) > 1
                d = dirs.get(first)
                if d == 'out' and first != outs[0]:
                    fan_system_effect.log.info('fan system effect: discharge %s of %s not '
                                               'priced (second discharge)', first, root_id)
                    continue
                rec = _fse_connection(root_id, first, tag, direction=d,
                                      cfm_in=dis_cfm if d == 'out' else None)
                if rec is not None:
                    recs[first] = rec
        return recs, rnotes

    # ── inline exhaust fans (Colin, 2026-11-03) ──────────────────────────────
    # Both connections are Exhaust Air. The one whose ductwork reaches the
    # grilles is the INLET, the other, with no terminals, the OUTLET
    # (discharge). No FlowDirectionType: the terminals decide.
    def _reaches_terminal(first):
        """[terminal ids] below node first (its subtree), walked once."""
        found, todo, seen = [], [first], set()
        while todo:
            n = todo.pop()
            if n in seen:
                continue
            seen.add(n)
            if n in all_terminals:
                found.append(n)
            todo.extend(all_children.get(n, []))
        return found

    def _inline_sides(root_id):
        """({first: 'in' or 'out'}, discharge cfm, notes, kind) for an
        exhaust fan that is not a ceiling fan, from its connections, read
        every run (Colin, 2026-11-03):
          'inline': an inlet (reaches the grilles) and an outlet (no
            grilles). The discharge is priced at the total grille cfm.
          'roof': inlet only, an upblast or downblast roof fan that
            discharges itself: the intake alone, no discharge walked. The
            fan type sets only the fan system effect and the discharge walk;
            the cap and backdraft boxes are charged as ticked (Colin,
            2026-11-03).
          None: no connection reaches the grilles; read on the
            classification, with a NOTE."""
        firsts = [f for f in all_children.get(root_id, []) if _kind(f) is not None]
        terms = dict((f, _reaches_terminal(f)) for f in firsts)
        intake = [f for f in firsts if terms[f]]
        # A connection straight into other equipment is no discharge (review,
        # 2026-11-03).
        equip = [f for f in firsts if not terms[f] and _kind(f) == 'equipment']
        dis = [f for f in firsts if not terms[f] and _kind(f) != 'equipment']
        for f in firsts:
            fan_system_effect.log.info('inline exhaust fan %s: connection %s (%s) reaches '
                                       '%d terminal(s)', root_id, f, _kind(f), len(terms[f]))
        if intake and not dis and equip:
            return {}, 0.0, [('note', 'NOTE: exhaust fan %s, connection %s feeds other '
                              'equipment; fan type not read, no discharge priced.'
                              % (root_id, ', '.join(str(f) for f in equip)))], None
        if not intake:
            # Colin, 2026-11-03: an error, not computed. build_network stops a
            # fan with no terminal at all first (main() reports it as an
            # error for an exhaust fan); this covers the rest.
            return {}, 0.0, [('error', 'ERROR: exhaust fan %s, no connection reaches a '
                              'grille or diffuser; the fan is not computed. Connect '
                              'its intake ductwork to the grilles.' % root_id)], 'error'
        if not dis:
            return (dict((f, 'in') for f in intake), 0.0,
                    [('note', 'NOTE: exhaust fan read as an upblast or downblast roof '
                      'fan (inlet only, at element %s): the worst intake path is priced, '
                      'no discharge walked.' % ', '.join(str(f) for f in intake))],
                    'roof')
        # The air picked up walking back from every grille to the fan, the
        # same post-order sum as every duct (Colin, 2026-11-03).
        cfm = sum([_connection_cfm(f) for f in intake], 0.0)
        dirs = dict((f, 'in') for f in intake)
        dirs.update((f, 'out') for f in dis)
        return dirs, cfm, [('note', 'NOTE: exhaust fan read as an inline fan: intake at '
                            'element %s, discharge at element %s (no grilles), carrying '
                            '%.0f cfm picked up from the grilles, added to every exhaust '
                            'run.' % (
                                ', '.join(str(f) for f in intake),
                                ', '.join(str(f) for f in dis), cfm))], 'inline'

    def _discharge_chain(first, cfm, boost=None):
        """The discharge duct and fittings from the fan out, priced at the
        intake cfm: no path ends down there, so they are added to every
        exhaust run of the fan, like the roof cap (Colin, 2026-11-03). The air
        moves with the walk, away from the fan, so fittings are priced on the
        supply tables in walk order. boost: the discharge connection's
        {transition id: velocity factor} (fan velocity basis), as the main
        walk applies it. A flex discharge duct is flagged: exhaust shall be
        rigid (Colin, 2026-11-03).
        Returns {'fric', 'fit', 'n_duct', 'n_fit', 'feet', 'comp_keys',
        'items', 'unpriced'}."""
        out = {'fric': 0.0, 'fit': 0.0, 'n_duct': 0, 'n_fit': 0, 'feet': 0.0,
               'comp_keys': (), 'items': (), 'unpriced': ()}
        boost = boost or {}
        nid, seen = first, set()
        up_fpm, up_area = 0.0, None
        while nid is not None and nid not in seen:
            seen.add(nid)
            elem = all_nodes.get(nid)
            dr = dr_by_int_id.get(nid)
            if dr is not None:
                fpm = cfm / dr.area_ft2 if dr.area_ft2 else 0.0
                d_h = getattr(dr, 'd_h_in', 0.0)
                fr = 0.0
                if d_h > 0 and fpm > 0:
                    fr = (hvac_graph.duct_friction_loss_per_100ft(
                        fpm, d_h, getattr(dr, 'roughness_ft', None))
                          * float(dr.length_ft) / 100.0)
                out['fric'] += fr
                out['n_duct'] += 1
                out['feet'] += float(dr.length_ft)
                up_fpm, up_area = fpm, dr.area_ft2
                out['items'] += (('Discharge duct {}'.format(nid), '{}, {:.1f} ft'.format(
                    getattr(dr, 'size', '?'), float(dr.length_ft)), fpm, fr,
                    ('Discharge duct', 'Darcy-Weisbach', None, float(dr.length_ft))),)
                # Exhaust is never flex: the low pressure collapses it
                # (Colin, 2026-11-03). Flagged, not priced as a flex run.
                if getattr(dr, 'is_flex', False):
                    out['unpriced'] += ('WARNING: flex duct %s on the exhaust discharge; '
                                        'exhaust ductwork shall be rigid' % nid,)
            elif elem is not None and hvac_graph.is_fitting(elem):
                if up_area is None:
                    flow = _first_duct(nid)
                    if flow is not None and flow.area_ft2:
                        up_area = flow.area_ft2
                        up_fpm = cfm / up_area
                _k = boost.get(nid)
                price_fpm = up_fpm * _k if (_k and up_fpm) else up_fpm
                loss, counted, note, down_a, detail = _price_fitting(
                    elem, 'Supply Air', price_fpm, up_area, against_flow=False, loc_id=nid)
                if _k and up_fpm:
                    detail = '{}, at fan velocity'.format(detail)
                fam = hvac_graph.fitting_family_name(elem) or '?'
                k = _takeoff_key(elem, fam, detail, counted, note)
                out['fit'] += loss
                out['n_fit'] += counted
                if note:
                    out['unpriced'] += (note,)
                out['items'] += (('Discharge fitting {}'.format(nid),
                                  '{}, {}'.format(fam, detail), up_fpm, loss,
                                  ('Discharge ' + k[0][:1].lower() + k[0][1:],) + k[1:]),)
                if down_a and up_area and up_fpm:
                    up_fpm, up_area = up_fpm * up_area / down_a, down_a
            elif elem is not None and hvac_graph.is_accessory(elem):
                fam = hvac_graph.fitting_family_name(elem)
                akey = fitting_tables.classify_accessory(fam)
                if akey is not None:
                    out['comp_keys'] += (akey,)
                    out['items'] += (('Discharge accessory {}'.format(nid), fam or '?', None,
                                      fitting_tables.component_of(akey, comp_values),
                                      (_COMP_LABEL.get(akey, akey), 'known drop', None,
                                       None)),)
                else:
                    out['unpriced'] += ((fam or '?') + ': accessory on the discharge, '
                                        'not priced',)
            else:
                break
            kids = [c for c in all_children.get(nid, []) if c not in seen]
            if len(kids) > 1:
                out['unpriced'] += ('discharge splits at element %s: priced along '
                                    'its first leg only' % nid,)
            nid = kids[0] if kids else None
        fan_system_effect.log.info('inline exhaust fan discharge from %s: %d ducts, %d '
                                   'fittings, %.1f ft, friction %.4f, fittings %.4f at '
                                   '%.0f cfm', first, out['n_duct'], out['n_fit'],
                                   out['feet'], out['fric'], out['fit'], cfm)
        return out

    _GROUP_ITEM = re.compile(r'^(?:Discharge )?(duct|fitting) (\d+)$', re.I)

    def _fse_group_max(rec, its):
        """One unit connection's group (rec['group']: the straight duct, any
        transition and the first elbow) against its fan system effect line,
        charged once as the larger (Colin, 2026-11-04).

        its: the take-off items of the path, or of a discharge chain.
        Returns (items, fse charged, friction removed, fittings removed,
        fse item or None, governs NOTE or None). The fan system effect wins a
        tie: its group items stay listed at 0. The group wins: the fan system
        effect line stays listed at 0, not charged, and the individual losses
        are charged (Colin, 2026-11-09). The NOTE and the fan system effect
        row both say which governed."""
        if rec is None or rec.get('item') is None or not rec.get('group'):
            item = (rec or {}).get('item')
            if item is None and rec and rec.get('none') and rec.get('where_item'):
                # No fan system effect here: its own row at 0.000 with the
                # reason, so outlet and inlet always read apart (2026-11-05).
                item = ('Fan system effect (none)', rec['none'], None, 0.0,
                        ('Fan system effect, ' + rec['where_item'],
                         'none: ' + rec['none'], None, None))
            return its, (rec or {}).get('loss', 0.0) or 0.0, 0.0, 0.0, item, None
        g_fric, g_fit, hits = 0.0, 0.0, []
        for i, it in enumerate(its):
            m = _GROUP_ITEM.match(it[0])
            if m is None or int(m.group(2)) not in rec['group']:
                continue
            hits.append(i)
            if m.group(1).lower() == 'duct':
                g_fric += it[3] or 0.0
            else:
                g_fit += it[3] or 0.0
        group = g_fric + g_fit
        fse = rec['loss']
        fan_system_effect.log.info('fan system effect group: %s, fse %.4f vs duct and '
                                   'fittings %.4f (%d items)', rec.get('where'), fse,
                                   group, len(hits))
        item = rec['item']
        gov = rec.get('gov') or {}
        note = None
        if gov:
            note = fan_system_effect.governs_text(
                fse >= group, gov['eid'], gov['l_ft'], rec.get('where'), fse, group,
                gov.get('win'), gov.get('win_short', 'Le'))
        if fse >= group:
            its = tuple(
                (it[0], '{}, in the fan system effect'.format(it[1]), it[2], 0.0, it[4])
                if i in hits else it for i, it in enumerate(its))
            item = (item[0], '{}; charged, at or above the straight duct and elbow '
                    '({:.4f} in. wc)'.format(item[1], group), item[2], item[3], item[4])
            return its, fse, g_fric, g_fit, item, note
        item = (item[0], '{}; not charged, the straight duct and elbow price higher '
                '({:.4f} in. wc)'.format(item[1], group), item[2], 0.0, item[4])
        return its, 0.0, 0.0, 0.0, item, note

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
        # Fan system effect per unit connection, charged to the paths that
        # leave the unit through that connection.
        # An inline exhaust fan: its sides from the terminals, and the
        # discharge priced once, added to each of its exhaust runs below.
        dis_rec = None
        inline_notes = []
        if root_id in inline_fan_roots:
            inline_cache[root_id] = _inline_sides(root_id)
            dirs_here, dis_cfm_here, inline_notes, fan_kind = inline_cache[root_id]
            fan_kinds[root_id] = fan_kind
        fse_recs, fse_root_notes = _fse_root(root_id)
        if root_id in inline_fan_roots:
            # After the fan system effect, so the discharge transitions take
            # the fan velocity its record sets.
            outs = [f for f in all_children.get(root_id, []) if dirs_here.get(f) == 'out']
            if outs and dis_cfm_here > 0:
                dis_rec = _discharge_chain(outs[0], dis_cfm_here,
                                           (fse_recs.get(outs[0]) or {}).get('boost'))
                dis_rec['first'] = outs[0]
                if len(outs) > 1:
                    inline_notes = inline_notes + [(
                        'note', 'NOTE: inline exhaust fan with %d discharge connections; '
                        'the one at element %s is priced.' % (len(outs), outs[0]))]
        fse_root_notes = list(inline_notes) + list(fse_root_notes)
        # A duct pick's record is keyed by the picked duct itself, on every path.
        fse_firsts = [c for c in list(all_children.get(root_id, [])) + [root_id]
                      if c in fse_recs]
        fse_skip = set()
        fse_gray = set()
        fse_red = set()
        fse_boost = {}     # transition id -> velocity factor (fan velocity basis)
        fse_red_ducts = {}  # straight duct id -> (reason, required), RED with its elbow
        fse_unit_notes = list(fse_root_notes)
        fse_unit_assume = []
        # Every connection's notes, not only the one the worst path leaves
        # through (review, 2026-10-29): a second outlet's GRAY is still shown.
        for _f in fse_firsts:
            _r = fse_recs[_f]
            fse_skip |= _r['skip']
            fse_gray |= _r['gray']
            fse_red |= _r['red']
            fse_boost.update(_r.get('boost') or {})
            for _d, _v in (_r.get('red_ducts') or {}).items():
                fse_red_ducts.setdefault(_d, _v)
            fse_unit_notes.extend(_r['notes'])
            for _a in _r['assumptions']:
                if _a not in fse_unit_assume:
                    fse_unit_assume.append(_a)
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
        # Last field: a balancing damper drawn on THIS branch, below the last
        # split or take-off the path passed (Colin, 2026-11-02: dampers on
        # mains are fire/smoke dampers, so only a branch damper replaces the
        # assumed one).
        pre_items = ()
        if pre[1] or pre[0]:
            pre_items = (('Unit side stretch', '{} fitting(s), {} duct(s), {:.0f} ft'.format(
                pre[2], pre[3], pre[4]), None, pre[0] + pre[1],
                ('Unit side stretch (picked duct to unit)', 'mixed', None, pre[4])),)
        stack = [(root_id, pre[0], pre[1], (), pre[3], pre[2], 0, pre[4], 0.0,
                  0.0, None, pre[5], frozenset([root_id]), False, pre_items, False)]
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
             items, br_damper) = stack.pop()

            elem = all_nodes.get(nid)
            dr   = dr_by_int_id.get(nid)

            if dr is not None:
                fric      = fric + dr.friction_loss_inwc
                dlen      = dlen + dr.length_ft
                dcount    = dcount + 1
                up_fpm    = dr.fpm
                up_area   = dr.area_ft2
                sys_class = getattr(dr, 'flow_class', None) or dr.sys_class
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

            elif elem is not None and nid in fse_skip:
                # An inlet elbow the fan system effect priced: ED7-2 / ER7-1
                # include the elbow's own loss (Colin, 2026-10-29), so it is
                # one line in the take-off, the fan system effect one.
                pass

            elif elem is not None and hvac_graph.is_fitting(elem):
                # A transition between the unit and the first disruption,
                # under the fan velocity basis: priced with the air at the fan
                # velocity (Colin, 2026-11-03). The walk carries on at the
                # model's own velocity past it.
                _k = fse_boost.get(nid)
                price_fpm = up_fpm * _k if (_k and up_fpm) else up_fpm
                loss, counted, note, _down_a, detail = _price_fitting(
                    elem, sys_class, price_fpm, up_area,
                    against_flow=fitting_tables.flows_toward_unit(sys_class),
                    nid=nid)
                if _k and up_fpm:
                    detail = '{}, at fan velocity'.format(detail)
                fit    = fit + loss
                fcount = fcount + counted
                if note:
                    unpriced = unpriced + (note,)
                _fam = hvac_graph.fitting_family_name(elem) or '?'
                if fitting_tables.classify_fitting(_fam) == 'takeoff':
                    br_damper = False      # a new branch starts at the take-off
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
                # matching none of the kinds is reported as uncounted rather
                # than priced wrong.
                fam  = hvac_graph.fitting_family_name(elem)
                akey = fitting_tables.classify_accessory(fam)
                if akey is not None:
                    comp_keys = comp_keys + (akey,)
                    if akey == 'damper':
                        br_damper = True
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
                    # Fan system effect is deliberately not part of the box
                    # check (Colin, 2026-10-29: no system effect on VAV boxes).
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
                # The diffuser the run ends at is itself a component drop; a
                # ceiling exhaust fan run ends at its wall or roof cap.
                is_cap = nid in cap_ids
                keys_here = comp_keys + ('cap' if is_cap else 'diffuser',)
                # Include exhaust roof cap: an inline fan's discharge cap,
                # charged once on each exhaust run (Colin, 2026-11-02).
                roof_cap = (exhaust_cap and not is_cap
                            and (term[1] or sys_class) == 'Exhaust Air')
                if roof_cap:
                    keys_here = keys_here + ('cap',)
                # Every supply diffuser is assumed to have a balancing damper
                # whether or not one is drawn (Colin, 2026-10-26), so a branch
                # with no damper accessory (a w_Damper take-off, or none shown)
                # still pays one. A damper drawn on the branch itself is never
                # charged twice; one on a main upstream of the branch is a
                # fire/smoke damper and does not cancel it (Colin, 2026-11-02).
                _assumed_damper = (
                        vav_checks.is_supply_class(term[1] or sys_class)
                        and not br_damper)
                if _assumed_damper:
                    keys_here = keys_here + ('damper',)
                # Include backdraft damper: an exhaust fan's run with none
                # modeled pays one (Colin, 2026-11-02); one drawn on the run
                # is already in comp_keys and is never charged twice.
                # The discharge of an inline exhaust fan: on every exhaust
                # run of the fan (Colin, 2026-11-03), like the roof cap.
                dis_here = (dis_rec if dis_rec is not None
                            and (term[1] or sys_class) == 'Exhaust Air' else None)
                if dis_here is not None:
                    keys_here = keys_here + dis_here['comp_keys']
                _assumed_backdraft = (
                        backdraft_roots is not None and root_id in backdraft_roots
                        and (term[1] or sys_class) == 'Exhaust Air'
                        and 'backdraft' not in comp_keys
                        and not (dis_here and 'backdraft' in dis_here['comp_keys']))
                if _assumed_backdraft:
                    keys_here = keys_here + ('backdraft',)
                comp_here = 0.0
                for _k in keys_here:
                    comp_here += _comp(_k)
                _cfm, term_class, _family = term
                key     = term_class or sys_class or 'Unknown'
                # Known drop added once to the RETURN path only (the "Include
                # Return Filter" option). Same on every return path, so it never
                # changes which run is the worst, only the total.
                filt    = return_filter_inwc if key == 'Return Air' else 0.0
                # The unit connection this path leaves through, and its fan
                # system effect (0 when none, or not checked).
                fse_rec = None
                for _first in fse_firsts:
                    if _first in path and (fse_rec is None or
                                           fse_recs[_first]['loss'] > fse_rec['loss']):
                        fse_rec = fse_recs[_first]
                # The group at the connection (straight duct, transition, first
                # elbow) is charged once with its fan system effect, the larger.
                items_l, fse_loss, _gfr, _gft, _fitem, _gnote = _fse_group_max(
                    fse_rec, items)
                fse_items = (_fitem,) if _fitem is not None else ()
                gov_notes = (('note', _gnote),) if _gnote else ()
                fric_l, fit_l, dis_items = fric - _gfr, fit - _gft, ()
                n_d, n_f, len_l, unpriced_l = dcount, fcount, dlen, unpriced
                if dis_here is not None:
                    _drec = fse_recs.get(dis_here['first'])
                    dis_items, _dl, _dfr, _dft, _ditem, _dnote = _fse_group_max(
                        _drec, dis_here['items'])
                    if _dnote:
                        gov_notes = gov_notes + (('note', _dnote),)
                    fric_l += dis_here['fric'] - _dfr
                    fit_l += dis_here['fit'] - _dft
                    n_d += dis_here['n_duct']
                    n_f += dis_here['n_fit']
                    len_l += dis_here['feet']
                    unpriced_l = unpriced_l + dis_here['unpriced']
                    if _drec is not None and _drec['loss']:
                        fse_loss += _dl
                    if _drec is not None and _ditem is not None:
                        fse_items = fse_items + (_ditem,)
                current = best_here.get(key)
                # Equal runs (two taps at one station, F2 2026-11-04): the
                # lower terminal id is the index run, so Revit's connector
                # order cannot change which terminal is named.
                _tot_here = fric_l + fit_l + comp_here + filt + fse_loss
                _tot_cur = None if current is None else (
                    current['friction_inwc'] + current['fitting_inwc'] +
                    current['component_inwc'] + current['filter_inwc'] +
                    current['fse_inwc'])
                if (current is None or _tot_here > _tot_cur + 1e-12 or
                        (abs(_tot_here - _tot_cur) <= 1e-12 and
                         nid < current['terminal_id'])):
                    best_here[key] = {
                        'friction_inwc':    fric_l,
                        'fitting_inwc':     fit_l,
                        'component_inwc':   comp_here,
                        'filter_inwc':      filt,
                        'fse_inwc':         fse_loss,
                        # The UNIT's fan system effect notes, the same on
                        # every class of this root: print them once per unit.
                        'fse_notes':        fse_unit_notes,
                        # Which governed at this run's connection(s): fan
                        # system effect or the individual losses (2026-11-09).
                        'fse_gov_notes':    gov_notes,
                        'fse_gray_ids':     fse_gray,
                        'fse_red_ids':      fse_red,
                        'fse_red_ducts':    fse_red_ducts,
                        'fse_assumptions':  fse_unit_assume,
                        'unit_side_fittings': pre[2],
                        'exhaust_fan_kind': ('ceiling' if root_id in discharge_roots
                                             else fan_kinds.get(root_id)),
                        'component_keys':   keys_here,
                        'duct_count':       n_d,
                        'fitting_count':    n_f,
                        'tap_bypass_count': tcount,
                        'length_ft':        len_l,
                        'terminal_id':      nid,
                        'terminal_name':    _family or '-',
                        'terminal_label':   _terminal_label(nid),
                        'unpriced':         unpriced_l,
                        'items':            fse_items + items_l + dis_items + (((
                            'Accessory (assumed)', 'Balancing damper at ' + (_family or '-'),
                            None, fitting_tables.component_of('damper', comp_values),
                            ('Balancing damper', 'known drop', None, None)),)
                            if _assumed_damper else ()) + (((
                            'Accessory (assumed)', 'Backdraft damper at the fan',
                            None, fitting_tables.component_of('backdraft', comp_values),
                            ('Backdraft damper', 'known drop', None, None)),)
                            if _assumed_backdraft else ()) + ((
                            'Terminal {}'.format(nid), _family or '-', None,
                            fitting_tables.component_of(
                                'cap' if is_cap else 'diffuser', comp_values),
                            (hvac_graph.CAP_LABEL if is_cap else 'Diffuser / grille',
                             'known drop', None, None)),) + (((
                            'Exhaust roof cap', 'Include exhaust roof cap', None,
                            fitting_tables.component_of('cap', comp_values),
                            (hvac_graph.CAP_LABEL, 'known drop', None, None)),)
                            if roof_cap else ()),
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
                c_fric, c_dlen = fric, dlen
                if dr is not None:
                    # Friction along a main with taps: each piece at its own
                    # flow, only as far as this run goes (_seg_friction).
                    seg = _seg_friction(nid, child_id)
                    _on = _tap_order_note(nid)
                    if _on and _on not in extra_unpriced:
                        extra_unpriced = extra_unpriced + (_on,)
                    if seg is not None:
                        _sf, _ft = seg
                        c_fric = fric - dr.friction_loss_inwc + _sf
                        c_dlen = dlen - float(dr.length_ft) + _ft
                        _name = 'Duct {}'.format(nid)
                        for _i in range(len(extra_items) - 1, -1, -1):
                            if extra_items[_i][0] == _name:
                                _old = extra_items[_i]
                                extra_items = extra_items[:_i] + ((
                                    _name,
                                    '{}, {:.1f} of {:.1f} ft, flow stepped at taps'.format(
                                        getattr(dr, 'size', '?'), _ft, float(dr.length_ft)),
                                    _old[2], _sf,
                                    (_old[4][0], _old[4][1], _old[4][2], _ft)),) + \
                                    extra_items[_i + 1:]
                                break
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
                c_class = sys_class
                if dr is not None and dr.area_ft2 and child_id not in _taps(nid):
                    # Past the last tap on an unbroken main only the flow the
                    # taps left is still moving, so a fitting there is priced
                    # at that velocity, not the unit end cfm (Colin,
                    # 2026-11-02, _remaining_cfm).
                    _rem = _remaining_cfm(nid)
                    if _rem is not None:
                        c_fpm = _rem / dr.area_ft2
                if nid == root_id and dr is None:
                    # Leaving the unit: start at the unit connection's
                    # velocity, not 0, so a fitting right on the unit is
                    # priced (Colin, 2026-10-29).
                    _st = _unit_start(root_id, child_id)
                    if _st is not None:
                        c_fpm, c_area, c_class = _st
                if box_here:
                    # Past a box the air moves at the box OUTLET duct's
                    # velocity, not the inlet runout's, so a fitting right on
                    # the box discharge is priced at the right Pv.
                    _od = _first_duct(child_id)
                    if _od is not None:
                        c_fpm, c_area = _od.fpm, _od.area_ft2
                stack.append((child_id, c_fric, extra_fit, comp_keys, dcount,
                              fcount, extra_t, c_dlen, c_fpm, c_area,
                              c_class, extra_unpriced,
                              path | frozenset([child_id]), in_flex,
                              extra_items,
                              # past a split the path is on a new branch
                              br_damper and len(all_children.get(nid, [])) < 2))

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
            # Fan system effect is in the subtotal, so the safety factor
            # applies to it like every other fitting (Colin, 2026-10-29).
            sub = (b['friction_inwc'] + b['fitting_inwc'] + b['component_inwc'] +
                   b['filter_inwc'] + b['fse_inwc'])
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
    'lines' of each record into the console's plain text fallback; the
    console and the drafting view render the other fields as a table, and
    duct_perf_console.vav_corrections() turns LOW boxes and diffusers on an
    MP main into System Corrections lines.

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
    # The MP diffuser line is a System Corrections line
    # (duct_perf_console.vav_corrections), not repeated here.
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
    return hvac_graph.fails_height_clearance(
        hvac_graph.effective_height_in(dr.elem), downstream_height_in,
        _CLEARANCE_MARGIN_IN)


def _clearance_reason(dr, down_h, down_elem, is_branch, branch_duct_ids, notes):
    """Specific Reason text for a clearance failure, and its NOTE appended to
    notes. What sits downstream decides the wording: the next duct of the same
    run, a branch duct leaving a main, or a diffuser neck."""
    duct_h = hvac_graph.effective_height_in(dr.elem)
    if duct_h is None or down_h is None or down_elem is None:
        return 'Diffuser/Duct Clearance'
    role = 'Branch' if is_branch else 'Main'
    down_id = eid_int(down_elem.Id)
    if hvac_graph.is_terminal(down_elem):
        kind = 'neck'
        down_label = '{} ({}" neck)'.format(down_id, hvac_graph._fmt_in(down_h))
    else:
        kind = ('branch' if (not is_branch and down_id in branch_duct_ids)
                else 'step')
        down_label = '{} ({})'.format(down_id, _duct_size_label(down_elem))
    up_label = '{} ({})'.format(dr.element_id, _duct_size_label(dr.elem))
    reason, note = hvac_graph.clearance_texts(kind, role, duct_h, down_h,
                                              up_label, down_label,
                                              _CLEARANCE_MARGIN_IN)
    notes.append(note)
    return reason


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

    A VAV medium pressure duct (dr.is_mp) is judged against the MP limits,
    PURPLE included: the shrink candidate is tested on the MP limits too
    (Colin, 2026-11-02; it was never PURPLE before).
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
    'round_sidewall': 'Branch (Round Sidewall)',
    'rect_ceiling':   'Branch (Rect Ceiling)',
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

    if dr.cfm <= 0:
        # A branch to a diffuser with Flow 0 cannot be judged: GRAY like a
        # main with no airflow, never a table pass (Colin, 2026-11-02, TB-M01
        # D-5 drew green). Only the clearance rule still overrides it.
        label, reason = 'GRAY', 'No airflow data'
        if _fails_height_clearance(dr, downstream_height_in):
            label, reason = 'RED', 'Diffuser/Duct Clearance'
        return label, 0.0, reason, None

    if (terminal_elem is not None and kind not in ('ceiling', 'sidewall')
            and not hvac_graph.is_auto_sizing_diffuser(terminal_elem)):
        # Not an RJA auto-sizing diffuser and no neck or face table to judge
        # it by (a slot from another family, whose length and slot count are
        # not read): the branch gets the main duct check. Ceiling and sidewall
        # diffusers from any family are judged on the RJA diffuser tables
        # (Colin, 2026-11-02; until then every non-RJA one got the main check).
        label, cap, reason = _duct_label(dr, custom_limits, tol_pct, downstream_height_in)
        return label, cap, reason, None

    if kind in ('round_sidewall', 'rect_ceiling'):
        # Round sidewall diffusers have no table yet; Colin plans to add round
        # diffuser data later (2026-10-16). A rectangular neck pointing up or
        # down is a ceiling diffuser, never judged on the sidewall tables, and
        # has no table of its own yet (Colin, 2026-11-02). Both fall back to
        # the main check.
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


def _duct_is_oval(elem, eid):
    """True when the duct's end connector is ConnectorProfileType.Oval. An
    oval duct is read as W x H for area, velocity and friction (hvac_graph),
    and the report says so (Colin, 2026-11-02). False when unreadable,
    logged."""
    try:
        for c in elem.ConnectorManager.Connectors:
            if c.ConnectorType == ConnectorType.End:
                if c.Shape == ConnectorProfileType.Oval:
                    hvac_graph.log.info('_duct_is_oval: duct %s Oval, read as W x H', eid)
                    return True
                return False
    except Exception as ex:
        hvac_graph.log.info('_duct_is_oval: duct %s connector shape unreadable: %s',
                            eid, ex)
        return False
    hvac_graph.log.info('_duct_is_oval: duct %s end connector NOT FOUND', eid)
    return False


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
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, float(std_d), dr.roughness_ft)
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
                    fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h, dr.roughness_ft)
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
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, float(std_d), dr.roughness_ft)
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
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h, dr.roughness_ft)
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
                    fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h, dr.roughness_ft)
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
def _legend_rows(words=None):
    """[(color key, meaning)] for the sheet legend, the branch wording
    from the Scope's words (Colin, 2026-11-09). The clearance wording is the
    Reason column's name and stays."""
    one = (words or duct_perf_console.scope_words('rtu'))['one']
    return [
        ('GREEN',  'Main: within limit. Branch: {} and duct both correctly sized.'.format(one)),
        ('PURPLE', 'Main only: oversized, a smaller standard size is available.'),
        ('YELLOW', 'Main only: approaching the limit.'),
        ('RED',    'Main: exceeds the limit. Branch: {} or duct undersized. '
                   'Either: fails the diffuser/duct height clearance (see Reason).'.format(one)),
        ('GRAY',   'No airflow data: no CFM reaches this element (broken or '
                   'disconnected system). Or an elbow at a unit connection whose fan '
                   'system effect is too high (see External Static Pressure).'),
    ]

# Columns read best left aligned; the rest are centred.
_LEFT_COLUMNS = ('reason', 'role')


# Sheet column proportions of the table width (the flagged duct table's).
_SHEET_EQUIP_COLS = (0.34, 0.16, 0.12, 0.38)
_SHEET_EQUIP_COLS_3 = (0.56, 0.24, 0.20)    # Exhaust System, no other air column
_SHEET_CLASS_W = 0.40       # by system class table, left of the types table
_SHEET_TYPES_W = 0.56       # diffuser types table, right of it
_SHEET_CLASS_COLS = (0.50, 0.30, 0.20)
_SHEET_TYPES_COLS = (0.64, 0.14, 0.22)
_SHEET_TAKEOFF_COLS = (0.34, 0.24, 0.14, 0.12, 0.16)
_SHEET_VAV_COLS = (0.22, 0.14, 0.16, 0.16, 0.18, 0.14)


def _sheet_columns(headers, fractions, width):
    return [(h, width * f) for h, f in zip(headers, fractions)]


def _draw_system_summary(st, view, y, total_w, rep):
    """SYSTEM SUMMARY: equipment table, by system class and diffuser types
    side by side. Its warnings are in SYSTEM CORRECTIONS. Returns the y under
    it."""
    draft = rja_drafting
    y = draft._band(st, view, 0.0, y, total_w, 'SYSTEM SUMMARY',
                    duct_perf_console.summary_band_text(rep), draft.SUB_BAND_H)
    words = rep.get('words') or duct_perf_console.scope_words('rtu')
    if rep.get('equipment'):
        heads = duct_perf_console.equipment_headers(words)
        four = len(heads) == 4
        y = draft._styled_table(
            st, view, 0.0, y, _sheet_columns(
                heads, _SHEET_EQUIP_COLS if four else _SHEET_EQUIP_COLS_3, total_w),
            [list(r[:len(heads)]) for r in rep['equipment']],
            left_cols=(0, 3) if four else (0,), clip=True)
        y -= draft.GAP
    if rep.get('caution'):
        # VAV Box scope with external static (Colin, 2026-11-09).
        y = draft._note_panel(st, view, 0.0, y, total_w, [rep['caution']])
        y -= draft.GAP
    bottoms = []
    if rep.get('by_class'):
        w = total_w * _SHEET_CLASS_W
        bottoms.append(draft._styled_table(
            st, view, 0.0, y, _sheet_columns(['System', 'CFM', words['count']],
                                             _SHEET_CLASS_COLS, w),
            rep['by_class'], clip=True))
    if rep.get('types'):
        w = total_w * _SHEET_TYPES_W
        ox = total_w - w if rep.get('by_class') else 0.0
        bottoms.append(draft._styled_table(
            st, view, ox, y, _sheet_columns([words['types'], 'Count', 'CFM'],
                                            _SHEET_TYPES_COLS, w),
            rep['types'], clip=True))
    if bottoms:
        y = min(bottoms) - draft.GAP
    return y


def _draw_external_static(st, view, y, total_w, esp_rec):
    """EXTERNAL STATIC PRESSURE: per unit one take-off table (supply then
    return, ending in the unit total) with where each index run ends and the
    unit max check, then picked duct pairings. Its warnings and notes are in
    SYSTEM CORRECTIONS. The rows come from duct_perf_console.esp_unit_rows,
    the same rows the console table shows. Returns the y under it."""
    draft = rja_drafting
    cols = _sheet_columns(duct_perf_console.TAKEOFF_HEADERS, _SHEET_TAKEOFF_COLS,
                          total_w)
    for unit in esp_rec.get('units') or []:
        y = draft._band(st, view, 0.0, y, total_w,
                        'EXTERNAL STATIC PRESSURE - ' + unit['label'], None,
                        draft.SUB_BAND_H)
        if unit.get('classes'):
            rows, total_row = duct_perf_console.esp_unit_rows(unit, sheet=True)
            if total_row is not None:
                rows.append(total_row)
            y = draft._styled_table(st, view, 0.0, y, cols, rows,
                                    totals_row=total_row is not None,
                                    left_cols=(0, 1), clip=True)
        notes = duct_perf_console.sheet_lines(
            duct_perf_console.esp_unit_notes(unit, sheet=True))
        if notes:
            y = draft._note_panel(st, view, 0.0, y, total_w, notes)
        y -= draft.GAP
    for comb in esp_rec.get('combined') or []:
        y = draft._band(st, view, 0.0, y, total_w,
                        'EXTERNAL STATIC PRESSURE - ' + comb['title'], None,
                        draft.SUB_BAND_H)
        y = draft._note_panel(st, view, 0.0, y, total_w, [comb['text']])
        y -= draft.GAP
    # No METHOD / ASSUMPTION / NOT INCLUDED block on the sheet (Colin,
    # 2026-11-02): the engineer reads the basis in the dialog's Calculation
    # basis. The console still prints it.
    return y


def _draw_vav_boxes(st, view, y, total_w, vav_rep):
    """VAV BOXES table, one row per box, with its table key (values in in.
    wc, the short cell texts). LOW boxes are in SYSTEM CORRECTIONS. Returns
    the y under it."""
    draft = rja_drafting
    rows = duct_perf_console.vav_rows(vav_rep, sheet=True)
    if not rows:
        return y
    y = draft._band(st, view, 0.0, y, total_w, 'VAV BOXES', vav_rep.get('basis'),
                    draft.SUB_BAND_H)
    y = draft._styled_table(
        st, view, 0.0, y, _sheet_columns(duct_perf_console.VAV_HEADERS,
                                         _SHEET_VAV_COLS, total_w),
        rows, left_cols=(0, 1), clip=True)
    notes = duct_perf_console.sheet_lines(duct_perf_console.vav_notes(vav_rep, sheet=True))
    if notes:
        y = draft._note_panel(st, view, 0.0, y, total_w, notes)
    return y - draft.GAP


def _draw_system_corrections(st, view, y, total_w, columns, rows, left_cols,
                             corrections):
    """SYSTEM CORRECTIONS: the flagged duct table (columns [(header,
    width)], rows already in keynote order), then every correction of the
    run in one block per group, ACTION REQUIRED then NOTES (Colin,
    2026-11-02). Nothing at all: one 'No corrections' line. Returns the y
    under it."""
    draft = rja_drafting
    y = draft._band(st, view, 0.0, y, total_w, 'SYSTEM CORRECTIONS',
                    '{} duct(s)'.format(len(rows)) if rows else None,
                    draft.SUB_BAND_H)
    groups = duct_perf_console.corrections_sheet(corrections)
    if rows:
        y = draft._styled_table(st, view, 0.0, y, columns, rows, left_cols=left_cols)
    else:
        y = draft._note_panel(st, view, 0.0, y, total_w, [
            duct_perf_console.NO_FLAGGED if groups
            else duct_perf_console.NO_CORRECTIONS])
    for heading, lines in groups:
        y -= draft.GAP / 2.0
        draft._t(st, view, 0.0, y, heading, 'bold')
        y -= draft.ROW_H
        y = draft._note_panel(st, view, 0.0, y, total_w, lines)
    return y - draft.GAP


def _build_summary_view(doc, rep, esp_rec, vav_rep, flagged_rows, selected_cols,
                        source_sheet_num, tn_type_id, ts, fill_id,
                        show_r_marker_note=False, corrections=()):
    """Create a Drafting View in the RJA table look: the System Summary,
    External Static Pressure and VAV Boxes sections drawn as tables from the
    same records the console renders (rep, esp_rec, vav_rep; the last two
    None when not computed), then SYSTEM CORRECTIONS (the flagged-duct table
    and every correction line of the run, corrections from main()) and the
    note legend (Colin, 2026-10-23: the summary looks as good as the rest of
    the sheet; 2026-11-02: one corrections block, no notes after a section).

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
        y -= rja_drafting.GAP

        # Each section fails soft on its own: a record shape this code did
        # not expect loses that section, never the sheet.
        sections = [(_draw_system_summary, rep)]
        if esp_rec is not None:
            sections.append((_draw_external_static, esp_rec))
        if vav_rep is not None:
            sections.append((_draw_vav_boxes, vav_rep))
        for draw, record in sections:
            try:
                y = draw(st, sched_view, y, total_w, record)
            except Exception as ex:
                hvac_graph.log.warning('Duct Performance sheet: %s not drawn: %s',
                                       draw.__name__, ex)
                y -= rja_drafting.GAP

        # '#' is this table's own row index: it matches the numbered
        # keynote circle placed on the same duct in the plan view.
        rows = [[str(ri + 1) if c[0] == 'num' else cells.get(c[0], '')
                 for c in cols] for ri, cells in enumerate(flagged_rows)]
        y = _draw_system_corrections(st, sched_view, y, total_w, columns, rows,
                                     left, corrections)

        # Legend: the plan colour as a swatch, then its meaning, boxed.
        # Named NOTE LEGEND and kept to the plan marks only (Colin, 2026-10-23).
        y = rja_drafting._band(st, sched_view, 0.0, y, total_w,
                               'NOTE LEGEND', None, rja_drafting.SUB_BAND_H)
        legend_top = y
        row_h = rja_drafting.ROW_H
        pad = rja_drafting.PAD
        swatch = 0.014
        for color_key, meaning in _legend_rows(rep.get('words')):
            if st.region_type_id is not None:
                rja_drafting._legend_swatch(
                    doc, sched_view, st.region_type_id, pad * 2.0,
                    y - swatch - (row_h - swatch) / 2.0, swatch,
                    _COLOR_MAP[color_key], fill_id)
            rja_drafting._t(st, sched_view, swatch + pad * 3.0, y,
                            '{}: {}'.format(color_key.title(), meaning))
            y -= row_h
        if show_r_marker_note:
            # Starts in the swatch column so '(R)' lines up with the colour
            # squares above it.
            rja_drafting._t(st, sched_view, pad * 2.0, y,
                            '(R): Most restrictive {}. It ends the index '
                            'run that sets the required external static '
                            'pressure for its equipment or system.'.format(
                                (rep.get('words') or
                                 duct_perf_console.scope_words('rtu'))['one']))
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
    # top_z is a bounding box Z (internal origin), so measure from the level's
    # ProjectElevation, not Elevation: Elevation follows the level type's base
    # (survey point) and put the cut plane 100 ft low on The Hub Boulder.
    vr.SetOffset(PlanViewPlane.CutPlane, top_z - base_lvl.ProjectElevation + 1.0)
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
    # VAV Box: external static from the box outlet to the worst diffuser
    # (Colin, 2026-11-09), low pressure ductwork after the box only.
    ext_static    = dialog_result['ext_static']
    safety_pct    = dialog_result['safety_pct']
    c_values      = dialog_result['c_values']
    comp_values   = dialog_result['comp_values']
    calc_basis    = dialog_result['calc_basis']
    filter_inwc   = dialog_result['filter_inwc']
    if dialog_result.get('scope') == 'vavbox':
        # No return path after a VAV box (Colin, 2026-11-09).
        filter_inwc = 0.0
    exhaust_cap   = dialog_result.get('exhaust_cap', False)
    exhaust_on    = dialog_result.get('exhaust', False)
    exh_set       = dialog_result.get('exhaust_settings') or {}
    # Include backdraft damper, default off (Colin, 2026-11-03); the legacy
    # dialog has no box for it.
    exh_backdraft = bool(exh_set.get('backdraft', False))
    vav_on        = dialog_result['vav']
    vav_set       = dialog_result['vav_settings']
    fan_effect    = dialog_result.get('fan_effect') or fan_system_effect.default_settings()

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


    # 3. List the Scope's equipment in the active view, and take ONE pick
    # (Colin, 2026-11-03: one unit per run, the Scope choice sets the list,
    # no duct pick).
    equip_in_view = list(FilteredElementCollector(doc, active_view.Id)
                         .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                         .WhereElementIsNotElementType())

    scope = dialog_result.get('scope', 'rtu')
    scope_label = _SCOPE_LABEL.get(scope, 'RTU')
    hidden = {'VAV box': 0, 'exhaust fan': 0, 'unit': 0}
    listed = []
    for eq in equip_in_view:
        kind = _equip_kind(eq)
        if kind in _SCOPE_LISTS[scope]:
            listed.append(eq)
        else:
            hidden[kind] += 1
    _hid = ['{} {}{} (choose {} in Scope)'.format(
        n, k, '' if n == 1 else 'es' if k.endswith('x') else 's', _KIND_SCOPE[k])
        for k, n in sorted(hidden.items()) if n]
    if _hid:
        output.print_md('Hidden from the equipment list: {}.'.format(', '.join(_hid)))
    equip_in_view = listed

    if not equip_in_view:
        output.print_md('No {} equipment found in the active view. Open a plan that '
                        'shows the unit, or choose another Scope.'.format(scope_label))
        return

    # Build display name → element map (deduplicate names)
    name_to_elem = {}
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

    choice = show_system_picker(sorted(name_to_elem.keys()), scope_label)
    if choice is None:
        output.print_md('**Cancelled.**')
        return

    # Every correction, warning and note of the run, gathered here and printed
    # once in System Corrections under the duct table, console and sheet
    # (Colin, 2026-11-02). Entries from duct_perf_console.correction().
    corrections = []

    def _correct(kind, text, unit=None, group=None):
        corrections.append(duct_perf_console.correction(kind, text, unit, group))

    mode, names = choice
    # A ceiling exhaust fan picked from the list: {fan id: {'cfm'}}. The fan
    # carries no airflow in the model, so it is entered here. No rated static
    # check (Colin, 2026-11-03): the engineer checks the fan selection.
    fan_inputs = {}
    if mode == 'pick':
        _picked = _pick_unit(scope, scope_label)
        if _picked is None:
            output.print_md('**Cancelled.**')
            return
        sel_elems = [_picked]
    else:
        sel_elems = [name_to_elem[n] for n in names]
    for eq in sel_elems:
        if not exhaust_on or not hvac_graph.is_ceiling_exhaust_fan(eq):
            continue
        vals = ui_helpers.show_exhaust_fan_inputs(
            'Duct Performance', '{} (id {})'.format(_elem_name(eq), eid_int(eq.Id)),
            hvac_graph.fan_rated_cfm(eq))
        if vals is None:
            output.print_md('**Cancelled.**')
            return
        fan_inputs[eid_int(eq.Id)] = vals

    # 4. Traverse each system and merge results
    all_duct_results = {}   # ElementId -> DuctResult  (worst-wins on overlap)
    all_nodes        = {}   # merged for fitting adjacency
    all_children     = {}   # merged
    all_root_ids     = []   # one per successfully traversed system
    root_labels      = {}   # root int_id -> equipment label, for per-unit output
    duct_rooted_ids  = set()  # roots that are a picked duct, not equipment (no unit connected)
    ahu_totals       = []   # (ahu_name, ahu_id, discharge_cfm, terminal_count, other_class_cfm)
    all_terminals    = {}   # int_id -> (cfm, sys_class, family_name)  (dedup across AHUs)
    all_zero_terms   = set()
    all_missing_flow = set()
    nodes_by_root    = {}   # root int_id -> set of node ids (parallel unit check)
    vav_by_root      = {}   # root int_id -> vav_checks.VavSystem (VAV on only)

    discharge_roots  = {}   # ceiling exhaust fan root int_id -> cap node int_id
    exhaust_fan_roots = set()   # every exhaust fan root (Exhaust System ticked)
    for sel_elem in sel_elems:
        _fan = fan_inputs.get(eid_int(sel_elem.Id))
        net = hvac_graph.build_network(
            sel_elem, doc, equipment_level=True,
            discharge_cfm=_fan['cfm'] if _fan else None)

        if net.errors:
            # The unit is skipped: an action for the engineer, in System
            # Corrections (printed even when no unit traced, see below). An
            # exhaust fan with no grille on either side is an ERROR, not
            # computed (Colin, 2026-11-03).
            _kind_err = ('error' if exhaust_on and hvac_graph.is_exhaust_fan(sel_elem)
                         else 'warn')
            for e in net.errors:
                _correct(_kind_err, e, '{} (id {})'.format(
                    _elem_name(sel_elem), eid_int(sel_elem.Id)))
            continue

        # Two picked ducts on the same system re-root at the same AHU. Without
        # this the unit would be listed, totalled and summed twice.
        if eid_int(net.root.Id) in all_root_ids:
            output.print_md('{} is already covered by **{}**, skipped.'.format(
                _elem_name(sel_elem),
                root_labels.get(eid_int(net.root.Id), _elem_name(net.root))))
            continue


        # build_network fell back to rooting at a duct. Only a duct pick did
        # that, and there is none now (Colin, 2026-11-03); kept so such a root
        # is still reported as duct rooted, not as a unit.
        if str(net.ahu_method).startswith('fallback'):
            duct_rooted_ids.add(eid_int(net.root.Id))

        unit_label_here = '{} (id {})'.format(_elem_name(net.root), eid_int(net.root.Id))

        # VAV: boxes and the MP supply. Run before the warnings are gathered
        # so an unreadable box direction is listed.
        if vav_on:
            vs = hvac_graph.analyze_vav(net)
            vav_by_root[eid_int(net.root.Id)] = vs
            if not vs.is_vav:
                _correct('warn', 'VAV System is ticked but no VAV box was found on '
                                 'this unit\'s supply. Judged as a constant volume '
                                 'system.', unit_label_here)
        nodes_by_root[eid_int(net.root.Id)] = set(net.nodes.keys())

        # Traversal warnings (Flow = 0, missing Flow, no dimensions, no unit
        # found). The pruned branch counts only say what was not walked: Notes.
        for w in net.warnings or []:
            if 'not traversed' in w:
                _correct('note', w, unit_label_here)
            else:
                _correct('warn', w, unit_label_here)

        root_id = eid_int(net.root.Id)
        all_root_ids.append(root_id)
        root_labels[root_id] = unit_label_here
        if exhaust_on and hvac_graph.is_exhaust_fan(net.root):
            exhaust_fan_roots.add(root_id)
        if net.discharge_end is not None:
            discharge_roots[root_id] = net.discharge_end
            _correct('note', 'NOTE: read as a ceiling exhaust fan (its only duct '
                     'connection is Exhaust Air, out): {:,.0f} CFM entered, the '
                     'discharge ends at element {}, priced as a {}.'.format(
                         float(_fan['cfm']), net.discharge_end,
                         hvac_graph.CAP_LABEL.lower()), unit_label_here)
        ahu_totals.append((
            _elem_name(net.root), root_id,
            net.equipment_discharge_cfm(root_id), len(net.terminal_cfms),
            {} if net.discharge_end is not None
            else net.equipment_other_class_cfm(root_id)))

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
            if nid == net.discharge_end:
                all_terminals[nid] = (cfm, 'Exhaust Air', hvac_graph.CAP_LABEL)
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
        _emit_corrections(corrections, subtitle='no duct results')
        output.print_md('**No duct results, check System Corrections above.**')
        return

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
        _correct('warn', 'Parallel units on one main. Its external static is unreliable.',
                 root_labels.get(rid, 'id {}'.format(rid)))

    mp_ids = set()
    vav_box_ids = set()
    inlet_runout_ids = set()   # MP with exactly one box downstream (review W5)
    for vs in vav_by_root.values():
        mp_ids |= vs.mp_ids
        vav_box_ids |= set(vs.boxes)
        inlet_runout_ids |= vs.inlet_runout_ids
    for dr in all_duct_results.values():
        dr.is_mp = dr.element_id in mp_ids
    # VAV Box scope: the picked boxes are walked as boxes, so no fan checks of
    # any kind at the box root (Colin, 2026-11-09): no fan system effect, no
    # straight duct or swirl RED, no damper window WARNING. A root is never
    # charged its own box drop (nid != root_id in the walker).
    box_root_ids = set(all_root_ids) if scope == 'vavbox' else set()
    walk_box_ids = vav_box_ids | box_root_ids

    # int duct id -> RED reason: a diffuser on an MP main. Applied after the
    # normal label, so it overrides it (an installation fault, like the
    # clearance rule). The flex at a box inlet rule was removed (Colin,
    # 2026-10-23).
    vav_red = {}
    if vav_on:
        vav_red = vav_checks.red_reasons(vav_by_root.values())

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
    # A unit connection serving only one terminal is a single diffuser main:
    # its whole run stays one branch, judged on the diffuser tables, and one
    # short note names it so the designer knows (Colin, 2026-11-02: very
    # rare, core and shell or a four way drop box diffuser;
    # hvac_graph.single_terminal_runs).
    term_ids = set(all_terminals.keys())
    equip_ids = set(nid for nid, e in all_nodes.items()
                    if e is not None and hvac_graph.is_equipment(e))
    downstream_terms = {}
    single_runs = {}      # (unit id, first node id) -> node ids on that run
    for rid in all_root_ids:
        partial = hvac_graph.compute_downstream_terminal_ids(
            rid, all_nodes, all_children, term_ids)
        for nid, tset in partial.items():
            if nid in downstream_terms:
                downstream_terms[nid] = downstream_terms[nid] | tset
            else:
                downstream_terms[nid] = tset
        if rid in discharge_roots:
            continue      # a fan discharge is a main, never a diffuser branch
        for first, run in hvac_graph.single_terminal_runs(
                rid, all_children, partial, equip_ids).items():
            single_runs[(rid, first)] = run
    for (rid, first), _run in sorted(single_runs.items()):
        _correct('note', hvac_graph.single_terminal_text(rid, first),
                 root_labels.get(rid, 'id {}'.format(rid)))
    hvac_graph.log.info('Branch rule: %d single diffuser main(s), judged as branches',
                        len(single_runs))

    branch_ctx    = {}   # ElementId -> (category, terminal_elem, terminal_cfm) or None
    duct_roles    = {}   # ElementId -> Duct Role display string
    kind_cache    = {}   # terminal int_id -> category  (many ducts share one terminal)

    for eid in all_duct_results.keys():
        if all_duct_results[eid].is_mp:
            # VAV medium pressure: always a main. With two or more boxes
            # downstream it is judged against the MP limits; with one it is
            # that box's inlet runout, sized to the box (Colin 2026-10-22).
            branch_ctx[eid] = None
            duct_roles[eid] = (vav_checks.ROLE_INLET
                               if eid_int(eid) in inlet_runout_ids
                               else vav_checks.ROLE_MP)
            continue
        tset = downstream_terms.get(eid_int(eid), set())
        _caps = set(discharge_roots.values())
        if (tset and tset <= _caps) or eid_int(eid) in _caps:
            # Ceiling exhaust fan discharge: judged against the Exhaust Air
            # main limits, 600 fpm / 0.05 (Colin, 2026-11-02).
            branch_ctx[eid] = None
            duct_roles[eid] = 'Main (fan discharge)'
            continue
        term_elem = None
        if len(tset) == 1 and hvac_graph.takeoff_child_ids(
                eid_int(eid), all_nodes, all_children):
            # A main stepped down to its last tap: the branch starts AT the
            # tap, so this piece is still a main, judged on the main limits
            # (Colin, 2026-11-05, TB-M10C 6x10 end piece read GRAY as a
            # rect branch on a round neck).
            branch_ctx[eid] = None
            duct_roles[eid] = _duct_role_label(False, None)
            continue
        if len(tset) == 1:
            term_elem = all_nodes.get(list(tset)[0])
        if term_elem is None:
            # 0 terminals downstream (a dead-end run), 2+ (a real main), or a
            # terminal that somehow isn't in the merged node map: all keep the
            # velocity/friction treatment.
            branch_ctx[eid] = None
            duct_roles[eid] = _duct_role_label(False, None)
            continue
        term_id = eid_int(term_elem.Id)
        if term_id not in kind_cache:
            kind_cache[term_id] = hvac_graph.classify_diffuser_branch(term_elem)
        kind = kind_cache[term_id]
        branch_ctx[eid] = (kind, term_elem, all_terminals.get(term_id, (0.0, '', ''))[0])
        duct_roles[eid] = _duct_role_label(True, kind)

    # 4b. System summary — total flow to each equipment, diffuser counts/types
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
    # summary_lines is the plain text form (the console fallback when
    # print_html fails); rep, esp_rec and vav_rep hold the same content as
    # records for the themed console sections (lib/duct_perf_console.py) and
    # the drafting view tables. Every number below is worked out once and
    # handed to both.
    # Labels follow the Scope that was picked (Colin, 2026-11-09): an
    # exhaust run reads Exhaust CFM and Grilles, never Supply or Diffusers.
    words = duct_perf_console.scope_words(scope)
    rep = {'grand': [], 'equipment': [], 'by_class': [], 'types': [], 'words': words}
    esp_rec = None       # External Static Pressure records, when computed
    vav_rep = None       # VAV Boxes records, when VAV is on with boxes
    summary_lines = ['SYSTEM SUMMARY']
    for name, aid, discharge_cfm, term_count, other_class_cfm in ahu_totals:
        summary_lines.append('Equipment: {} (id {})  -  {:.0f} CFM {} ({} {})'.format(
            name, aid, discharge_cfm, words['discharge'], term_count, words['many']))
        others = []
        # Exhaust System has no other air column: the exhaust cfm is the
        # airflow column already (it was listed twice).
        for cls in (sorted(other_class_cfm.keys()) if words['other'] else ()):
            if not other_class_cfm[cls] or other_class_cfm[cls] <= 0:
                continue   # '0 CFM Unknown' says nothing (Colin, 2026-11-02)
            summary_lines.append('    + {:.0f} CFM {} feeding this equipment (not counted in discharge total)'.format(
                other_class_cfm[cls], cls))
            others.append('{:,.0f} CFM {}'.format(float(other_class_cfm[cls]), cls))
        row = ['{} (id {})'.format(name, aid), '{:,.0f}'.format(float(discharge_cfm)),
               str(term_count)]
        if words['other']:
            row.append(', '.join(others) if others else '-')
        rep['equipment'].append(row)
    # One CFM line per system present, never summed, then the counts
    # (Colin, 2026-11-09).
    rep['grand'] = duct_perf_console.grand_rows(
        dict((k, v[0]) for k, v in sys_totals.items()),
        len(all_terminals), len(all_duct_results))
    summary_lines.append('  |  '.join('{}: {}'.format(k, v) for k, v in rep['grand']))
    for sys_class in sorted(sys_totals.keys()):
        s_cfm, s_cnt = sys_totals[sys_class]
        summary_lines.append('  {}: {:.0f} CFM  ({} {})'.format(sys_class, s_cfm, s_cnt,
                                                              words['many']))
        rep['by_class'].append([sys_class, '{:,.0f}'.format(float(s_cfm)), str(s_cnt)])
    summary_lines.append(words['types'] + 's:')
    for family in sorted(type_totals.keys()):
        t_cnt, t_cfm = type_totals[family]
        summary_lines.append('  {}  x{}  ({:.0f} CFM)'.format(family, t_cnt, t_cfm))
        rep['types'].append([family, str(t_cnt), '{:,.0f}'.format(float(t_cfm))])
    # Flow = 0 and missing Flow terminals are not repeated here: each unit's
    # traversal warning already says it, per unit, in System Corrections
    # (all_zero_terms and all_missing_flow are the union of those).
    # Oval ducts are read as rectangles, W x H (Colin, 2026-11-02): kept, and
    # said once for the run, naming the ducts, in System Corrections.
    oval_ids = sorted(dr.element_id for dr in all_duct_results.values()
                      if _duct_is_oval(dr.elem, dr.element_id))
    if oval_ids:
        _ids = ', '.join(str(i) for i in oval_ids[:5])
        if len(oval_ids) > 5:
            _ids += ' and {} more'.format(len(oval_ids) - 5)
        _msg = ('NOTE: oval duct {} read as rectangular W x H for area, velocity and '
                'friction; oval sizing is not supported.'.format(_ids))
        hvac_graph.log.info('Duct Performance: %d oval duct(s): %s', len(oval_ids), oval_ids)
        _correct('note', _msg)

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
    fse_gray_ids = set()      # elbows drawn GRAY: fan system effect too high
    fse_red_ids = set()       # elbows drawn RED: two turn planes within 6 Dh of a unit
    fse_red_ducts = {}        # straight duct id -> (reason, required): RED with its elbow
    fse_extra_assume = []     # ASSUMPTION lines from the priced fittings
    if not ext_static:
        # External static off: the straight duct and swirl checks still run
        # (Colin, 2026-11-03, Q4). The walk runs with the fan system effect
        # unticked and only its RED lines and red ids are used.
        _chk = _critical_path_loss(
            all_root_ids, all_children, all_duct_results, all_terminals,
            all_nodes, 0.0, c_values, comp_values,
            duct_rooted_ids=duct_rooted_ids, vav_box_ids=walk_box_ids,
            fan_effect=dict(fan_effect or fan_system_effect.default_settings(), on=False),
            discharge_roots=discharge_roots,
            inline_fan_roots=set(exhaust_fan_roots) - set(discharge_roots or ()))
        for root_id in all_root_ids:
            per_class = _chk.get(root_id)
            if not per_class:
                continue
            _leaf = per_class[sorted(per_class.keys())[0]]
            for _kind, _text in _leaf.get('fse_notes', ()):
                if _kind == 'red':
                    _correct('red', _text, root_labels.get(root_id, 'id {}'.format(root_id)))
            fse_red_ids.update(_leaf.get('fse_red_ids', ()))
            for _d, _v in (_leaf.get('fse_red_ducts') or {}).items():
                fse_red_ducts.setdefault(_d, _v)
    if ext_static:
        critical = _critical_path_loss(
            all_root_ids, all_children, all_duct_results, all_terminals,
            all_nodes, safety_pct, c_values, comp_values,
            duct_rooted_ids=duct_rooted_ids, return_filter_inwc=filter_inwc,
            vav_box_ids=walk_box_ids,
            vav_box_drop_inwc=vav_set['box_drop'] if vav_on else None,
            fan_effect=fan_effect, discharge_roots=discharge_roots,
            exhaust_cap=exhaust_cap,
            # Every exhaust fan but a ceiling one, whose damper is built in.
            backdraft_roots=(set(exhaust_fan_roots) - set(discharge_roots or ())
                             if exh_backdraft else None),
            # Inline exhaust fans: every exhaust fan but a ceiling one.
            inline_fan_roots=set(exhaust_fan_roots) - set(discharge_roots or ()))
        summary_lines.append(
            'TOTAL EXTERNAL STATIC PRESSURE ({})'.format(duct_perf_console.ESP_BASIS))
        if box_root_ids:
            summary_lines.append(duct_perf_console.VAVBOX_ESP_CAUTION)
            rep['caution'] = duct_perf_console.VAVBOX_ESP_CAUTION
        esp_rec = {'units': [], 'combined': [],
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
            unit_rec = {'label': unit_label, 'classes': [],
                        'esp': None, 'one_sided': None, 'checks': []}
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
                if c.get('fse_inwc', 0.0) > 0.0:
                    esp_rows.append(('Fan system effect', c['fse_inwc'],
                                     'unit connection, 2001 Fundamentals Eq. 38'))
                for _item, _val, _det in esp_rows:
                    summary_lines.append(_esp_text_line(_item, _val, _det))
                # Every line said about this run goes to System Corrections
                # (Colin, 2026-11-02), not under the take-off.
                if c.get('unit_side_fittings', 0) > 0:
                    _correct('note', '{}: (Includes {} fitting(s) between the picked '
                             'duct and the unit connection.)'.format(
                                 sys_class, c['unit_side_fittings']), unit_label)
                safety_inwc = c['total_inwc'] - c['subtotal_inwc']
                safety_det = '+{:.0f}%'.format(safety_pct)
                total_det = 'worst run ends at {}'.format(
                    c.get('terminal_label') or c['terminal_name'])
                summary_lines.append(_esp_text_line('Subtotal', c['subtotal_inwc']))
                summary_lines.append(_esp_text_line('Safety factor', safety_inwc, safety_det))
                summary_lines.append(_esp_text_line('TOTAL', c['total_inwc'], total_det))
                for u in c['unpriced']:
                    # A FALLBACK note is priced (RJA worksheet C), not missing.
                    # A NOTE (a lookup read at a table edge, tap order
                    # unreadable) lands in Notes, the rest in Action required.
                    _u = (u if (u.startswith('FALLBACK') or u.startswith('NOTE'))
                          else 'UNPRICED: ' + u)
                    _correct('warn', _u, unit_label)
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
                    _correct('warn', '{}: {}{}'.format(sys_class, _msg[0].upper(), _msg[1:]),
                             unit_label)
                _filt = c['filter_inwc'] if c.get('filter_inwc', 0.0) > 0.0 else None
                if c.get('truncated'):
                    _msg = ('path search hit its step limit, so this '
                            'may not be the true worst run.')
                    _correct('warn', '{}: {}{}'.format(sys_class, _msg[0].upper(), _msg[1:]),
                             unit_label)
                unit_rec['classes'].append({
                    'sys_class': sys_class,
                    'airflow': '{:,.0f} CFM, {} terminals'.format(float(a_cfm), a_cnt),
                    'rows': esp_rows, 'subtotal': c['subtotal_inwc'],
                    'takeoff': _to_rows, 'filter': _filt,
                    'safety': safety_inwc, 'safety_detail': safety_det,
                    'total': c['total_inwc'], 'total_detail': total_det,
                    'notes': [], 'warnings': []})

            # Fan system effect notes belong to the unit, not to a class:
            # every class leaf carries the same list, so say it once, in
            # System Corrections. An ASSUMPTION stays with the ESP basis.
            _fse_leaf = per_class[sorted(per_class.keys())[0]]
            for _kind, _text in _fse_leaf.get('fse_notes', ()):
                if _kind == 'assume':
                    if _text not in fse_extra_assume:
                        fse_extra_assume.append(_text)
                else:
                    _correct(_kind, _text, unit_label)
            # Which governed at each index run's connection: per class, since
            # supply and return leave through different connections.
            for _cls in sorted(per_class.keys()):
                for _kind, _text in per_class[_cls].get('fse_gov_notes', ()):
                    _correct(_kind, _text, unit_label)
            fse_gray_ids.update(_fse_leaf.get('fse_gray_ids', ()))
            fse_red_ids.update(_fse_leaf.get('fse_red_ids', ()))
            for _d, _v in (_fse_leaf.get('fse_red_ducts') or {}).items():
                fse_red_ducts.setdefault(_d, _v)
            for _a in _fse_leaf.get('fse_assumptions', ()):
                if _a not in fse_extra_assume:
                    fse_extra_assume.append(_a)

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
                _correct('warn', _msg.replace(':  ', ': ', 1), unit_label)

            unit_total = 0.0
            sides = []
            if supply is not None:
                unit_total += supply['total_inwc']
                sides.append('Supply Air')
            if returns:
                worst_k, worst_c = max(returns, key=lambda kv: kv[1]['total_inwc'])
                unit_total += worst_c['total_inwc']
                sides.append(worst_k)
            if (root_id in discharge_roots or root_id in exhaust_fan_roots
                    or root_id in box_root_ids) and sides:
                # An exhaust fan has one side and that is the whole system.
                # No rated static check (Colin, 2026-11-03). A VAV box root
                # is its low pressure supply alone, never the total
                # (Colin, 2026-11-09).
                _sides = ' + '.join(sides)
                _is_box = root_id in box_root_ids
                summary_lines.append(
                    '    {} = {:.3f} in. wc   ({})'.format(
                        duct_perf_console.VAVBOX_ESP_LABEL if _is_box
                        else 'EXTERNAL STATIC PRESSURE', unit_total, _sides))
                if _is_box:
                    summary_lines.append('    ' + duct_perf_console.VAVBOX_ESP_CAUTION)
                    unit_rec['caution'] = duct_perf_console.VAVBOX_ESP_CAUTION
                unit_rec['esp'] = _esp_console_text(unit_total, _sides)
                unit_rec['esp_total'] = unit_total
                unit_rec['esp_sides'] = list(sides)
                # The fan type read from its connections (Colin, 2026-11-03).
                _fk = None
                for _leaf in per_class.values():
                    _fk = _fk or _leaf.get('exhaust_fan_kind')
                unit_rec['fan_type'] = _EXHAUST_FAN_KIND_TEXT.get(_fk)
                if unit_rec['fan_type']:
                    summary_lines.append('    ' + unit_rec['fan_type'])
            elif supply is not None and returns:
                summary_lines.append(
                    '    EXTERNAL STATIC PRESSURE = {:.3f} in. wc   ({})'.format(
                        unit_total, ' + '.join(sides)))
                unit_rec['esp'] = _esp_console_text(unit_total, ' + '.join(sides))
                unit_rec['esp_total'] = unit_total
                unit_rec['esp_sides'] = list(sides)
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
                if _ok:
                    # A pass is a result of the take-off and stays with it.
                    summary_lines.append('    ' + _msg)
                    unit_rec['checks'].append((_ok, _msg))
                else:
                    _correct('error', _msg, unit_label)
            # Parallel units on one main: said once per unit in System
            # Corrections (step 4-VAV above), not repeated here.

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
            _correct('warn', _msg[0].upper() + _msg[1:])

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
        _fse_method, _fse_assume = fan_system_effect.method_lines(fan_effect)
        _fse_assume = list(_fse_assume) + [a for a in fse_extra_assume
                                           if a not in _fse_assume]
        _method += ' ' + _fse_method
        # The filter clause only when a return run was priced: an exhaust
        # run has no return path (Colin, 2026-11-04).
        _return_priced = any('Return Air' in (_pc or {})
                             for _pc in (critical or {}).values())
        # Worded per Scope (Colin, 2026-11-09): the roof cap clause on
        # Exhaust System only, an exhaust fan's own housing in place of the
        # coil and cabinet, a VAV box's own drop and what feeds it.
        if scope == 'exhaust':
            _not_inc = ("NOT INCLUDED: the fan's own housing losses, already in its "
                        'published fan curve.')
        elif scope == 'vavbox':
            _not_inc = ('NOT INCLUDED: the VAV box drop and everything upstream of the '
                        'box outlet (see the caution above).')
        elif filter_inwc > 0.0 and _return_priced:
            _not_inc = ('NOT INCLUDED: coil and cabinet, already in the published ESP. '
                        'Return filter IS included as a {:.3f} in. wc known drop.'.format(
                            filter_inwc))
        else:
            _not_inc = ('NOT INCLUDED: filter, coil and cabinet, already in the published '
                        'ESP.')
        if scope == 'exhaust':
            _not_inc += (' Exhaust roof / wall cap IS included on every exhaust run.'
                         if exhaust_cap
                         else ' Exhaust roof / wall cap not included (upblast roof fan, '
                              'or Include roof / wall cap unticked).')
        if exhaust_on:
            _not_inc += (' A backdraft damper IS added to an exhaust fan run with none '
                         'modeled (not a ceiling exhaust fan, built in).' if exh_backdraft
                         else ' No backdraft damper is added where none is modeled.')
        summary_lines.append('')
        summary_lines.append('  ' + _method)
        for _a in _fse_assume:
            summary_lines.append('  ' + _a)
        summary_lines.append('  ' + _not_inc)
        esp_rec['method'] = _method
        esp_rec['assumptions'] = list(_fse_assume)
        esp_rec['not_included'] = _not_inc

    vav_start = len(summary_lines)
    if vav_on and any(vs.boxes for vs in vav_by_root.values()):
        vav_rep = _vav_box_report(
            all_root_ids, vav_by_root, root_labels, critical, vav_set)
        summary_lines.extend(_vav_summary_lines(vav_rep))
        corrections.extend(duct_perf_console.vav_corrections(vav_rep))

    # The plain text is only the console fallback below, folded per section
    # (folding is per line, so a slice folds the same as the whole). The
    # drafting view draws tables from rep, esp_rec and vav_rep instead.
    raw_summary = summary_lines

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
        # One NOTE per clearance failure, naming both elements (Colin,
        # 2026-11-02: a main stepping down too little and a branch too tall for
        # its main shall read differently). (duct id, NOTE) pairs; each is an
        # Action required line in System Corrections, under its unit.
        clearance_notes = []
        branch_duct_ids = set(eid_int(k) for k, v in branch_ctx.items() if v is not None)

        # A short stub straight off a unit connection that runs into a
        # transition is the connection adapter at the unit's own size, not an
        # oversized duct: never PURPLE (Colin, 2026-11-05, TB-M05C / M12C /
        # M13C exhaust fans, 12x12 stub then transition to 10x10).
        _ADAPTER_MAX_FT = 2.0
        _parent_of = {}
        for _p, _cs in all_children.items():
            for _c in _cs:
                _parent_of.setdefault(_c, _p)
        def _unit_adapter_stub(nid, dr):
            par = all_nodes.get(_parent_of.get(nid))
            if par is None or not hvac_graph.is_equipment(par):
                return False
            if (getattr(dr, 'length_ft', 0.0) or 0.0) > _ADAPTER_MAX_FT:
                return False
            kids = [all_nodes.get(c) for c in all_children.get(nid, [])]
            return (len(kids) == 1 and kids[0] is not None
                    and hvac_graph.is_fitting(kids[0])
                    and fitting_tables.classify_fitting(
                        hvac_graph.fitting_family_name(kids[0])) == 'transition')

        for eid, dr in all_duct_results.items():
            downstream_h, downstream_elem = hvac_graph.downstream_height_elem(
                eid_int(eid), all_nodes, all_children)
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
                if label == 'PURPLE' and _unit_adapter_stub(eid_int(eid), dr):
                    hvac_graph.log.info('duct %s: unit connection adapter stub, '
                                        'not judged oversized', eid_int(eid))
                    label, reason = 'GREEN', ''
            if dr.element_id in vav_red:
                label, reason = 'RED', vav_red[dr.element_id]
            # The straight duct before an elbow under the straight duct line
            # at a unit connection (Colin, 2026-11-03).
            if dr.element_id in fse_red_ducts and reason != 'Diffuser/Duct Clearance':
                label, reason = 'RED', fse_red_ducts[dr.element_id][0]
            duct_labels[eid]    = (label, green_cap, reason)
            branch_results[eid] = branch_res
            if reason == 'Diffuser/Duct Clearance':
                clearance_count += 1
                _cn = []
                reason = _clearance_reason(dr, downstream_h, downstream_elem,
                                           eid_int(eid) in branch_duct_ids,
                                           branch_duct_ids, _cn)
                clearance_notes.extend((dr.element_id, _n) for _n in _cn)
                duct_labels[eid] = (label, green_cap, reason)
            counts[label] = counts.get(label, 0) + 1
            elem_labels[eid] = label

        # A clearance NOTE asks for a resize: Action required, named by the
        # first picked unit whose ductwork holds the duct.
        def _unit_of(nid):
            for _rid in all_root_ids:
                if nid in nodes_by_root.get(_rid, ()):
                    return root_labels.get(_rid)
            return None
        for _did, _n in sorted(clearance_notes, key=lambda dn: dn[1]):
            _correct('warn', _n, _unit_of(_did), duct_perf_console.ACTION)

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
            # An elbow at a unit connection whose fan system effect is too
            # high is GRAY (Colin, 2026-10-29), unless a duct beside it
            # fails outright (RED keeps priority, as everywhere else).
            if nid in fse_gray_ids and worst != 'RED':
                worst = 'GRAY'
            # An elbow of a pair turning in different planes within 6 Dh of
            # a unit is RED (Colin, 2026-11-02), whatever the ducts beside it.
            if nid in fse_red_ids:
                worst = 'RED'
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
            if dr.element_id in fse_red_ducts and reason == fse_red_ducts[dr.element_id][0]:
                cells['required'] = fse_red_ducts[dr.element_id][1]
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
                    _loc = term_elem.Location
                    if hasattr(_loc, 'Curve'):
                        # A ceiling exhaust fan's cap is the open end duct.
                        r_pt = _loc.Curve.Evaluate(0.5, True)
                    else:
                        r_pt = _loc.Point
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
                    doc, rep, esp_rec, vav_rep, [item[2] for item in flagged_items],
                    selected_cols, source_sheet_num, tn_type_id, ts, fill_id,
                    show_r_marker_note=bool(restrictive_terminal_ids),
                    corrections=corrections)
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

    # 9. Console report: System Corrections (flagged ducts and every
    # correction line), design limits and legend, sheets.
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

    flagged.sort(key=_flag_sort_key)

    # Same shared column definitions and same user selection as the
    # schedule view. Always printed (Colin, 2026-11-02): with nothing
    # flagged and nothing to say it reads 'No corrections'.
    cols = _selected_columns(selected_cols)
    duct_rows = []
    for idx, (label, fpm, cells) in enumerate(flagged, 1):
        duct_rows.append([
            str(idx) if col[0] == 'num' else cells.get(col[0], '')
            for col in cols])
    keys = [col[0] for col in cols]

    def _fmt_row(cells):
        # Padded monospace text, console widths from _COLUMN_DEFS.
        return '  '.join(str(c).ljust(col[3]) for c, col in zip(cells, cols))
    duct_text = None
    if duct_rows:
        duct_text = [_fmt_row([col[1] for col in cols]),
                     _fmt_row(['-' * col[3] for col in cols])]
        duct_text.extend(_fmt_row(r) for r in duct_rows)

    _emit_corrections(
        corrections, [col[1] for col in cols], duct_rows,
        status_col=keys.index('status') if 'status' in keys else None,
        align=['r' if k in ('num', 'cfm', 'fpm', 'fric', 'length') else 'l'
               for k in keys],
        widths=[col[3] for col in cols],
        subtitle=('{} duct(s), every duct listed'.format(len(duct_rows)) if full_diag
                  else '{} duct(s)'.format(len(duct_rows))),
        rows_text=duct_text)

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
         'Main: within limit. Branch: {} and duct both correctly sized'.format(words['one'])],
        ['Purple', str(counts.get('PURPLE', 0)), str(fitting_counts.get('PURPLE', 0)),
         'Main only: oversized (low velocity, review for cost)'],
        ['Yellow', str(counts.get('YELLOW', 0)), str(fitting_counts.get('YELLOW', 0)),
         'Main only: approaching limit'],
        ['Red', str(counts.get('RED', 0)), str(fitting_counts.get('RED', 0)),
         'Main: exceeds limit. Branch: {} or duct undersized'.format(words['one'])],
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
            ('Branch ducts are checked against the {} sizing tables '
             '(MEP Design Standards, section 1), not against velocity or '
             'friction.'.format(words['tables']),
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
        'words': words,
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
            ['Color', 'Ducts', duct_perf_console.legend_fittings_header(words), 'Meaning'],
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
