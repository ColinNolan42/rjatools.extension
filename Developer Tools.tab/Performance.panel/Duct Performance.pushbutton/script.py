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
          Sizing Warning (Colin, 2026-10-22), and its Flow against the slot
          table for its length and slot count (Colin, 2026-11-09).

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
    BoundingBoxXYZ,
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

# Engine moved to lib/duct_perf_run.py (2026-11-12); bound to this
# document and output window below.
import duct_perf_run
from duct_perf_run import (
    MEETS_STANDARD_REASON, MR_LINE_COLOR, MR_LINE_WEIGHT, _COLOR_MAP,
    _PRIORITY, _build_summary_view, _canon_level_pick, _clearance_reason,
    _create_level_plan, _critical_path_loss, _crop_box_xy, _crop_to_elements,
    _duct_is_oval, _duct_role_label, _elem_box_points, _elem_level_id,
    _elem_name, _emit_corrections, _emit_html, _esp_console_text,
    _esp_text_line, _fails_height_clearance, _file_like_source_sheet,
    _fit_level_view_scale, _label_duct, _layout_level_viewports, _level_name,
    _plan_levels, _root_class_airflow, _row_cells, _selected_columns,
    _sheets_by_view, _target_plan_views, _vav_box_report, _vav_summary_lines,
    _view_name, _wrap_summary_lines,
)
duct_perf_run.bind(doc, output)

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
                <RadioButton x:Name="rb_scope_rtu" GroupName="scope" Content="Rooftop Unit"
                             ToolTip="A constant volume rooftop or packaged unit with supply and return. The roof curb is counted. VAV boxes and exhaust fans are not listed."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,0"
                           Text="Lists units only, no VAV boxes or exhaust fans."/>
                <!-- Scope options (Colin, 2026-11-12): the accessories of the
                     selected Scope, moved under its radio button by code -->
                <Border x:Name="pnl_scope_opts" Style="{DynamicResource RjaGroupBorder}" Margin="20,4,0,4">
                    <StackPanel>
                        <StackPanel x:Name="pnl_curb" Orientation="Horizontal">
                            <CheckBox x:Name="cb_curb" Content="Include roof curb" Width="190" VerticalAlignment="Center"
                                      ToolTip="Counts the curb height as straight duct at the unit connections."/>
                            <TextBox x:Name="tb_curb" Width="56" Margin="0,0,6,0"
                                     ToolTip="Curb height, in., 0 to 48."/>
                            <TextBlock Text="in. high" VerticalAlignment="Center"/>
                        </StackPanel>
                        <StackPanel x:Name="pnl_vav_static" Margin="0,8,0,0" Visibility="Collapsed">
                            <TextBlock FontWeight="SemiBold" Text="VAV boxes"/>
                            <Grid x:Name="grid_vav_static" Margin="0,4,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                            </Grid>
                        </StackPanel>
                        <StackPanel x:Name="pnl_exh_static" Visibility="Collapsed">
                            <StackPanel Orientation="Horizontal">
                                <CheckBox x:Name="cb_exh_backdraft" Content="Include backdraft damper" Width="190" VerticalAlignment="Center"
                                          ToolTip="Adds the backdraft damper drop once to an exhaust fan run."/>
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
                        <TextBlock x:Name="tb_scope_opts_hint" Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="0,4,0,0"
                                   Text="Used when external static pressure is calculated (Outputs)."/>
                    </StackPanel>
                </Border>
                <RadioButton x:Name="rb_scope_fancoil" GroupName="scope" Content="Indoor System" Margin="0,8,0,0"
                             ToolTip="An indoor unit with supply and return: fan coil, furnace, indoor AHU. Checked like an RTU."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,0"
                           Text="Lists units only: fan coil, furnace, indoor AHU."/>
                <RadioButton x:Name="rb_scope_vav" GroupName="scope" Content="VAV System" Margin="0,8,0,0"
                             ToolTip="The full system from the unit: supply mains to the VAV box inlets checked as medium pressure (MP), the boxes, low pressure after them and the return, with external static."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,0"
                           Text="Lists units only (RTU, AHU). Boxes are found by position: supply equipment entered through its inlet. "/>
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
                <CheckBox x:Name="cb_full_diag" Content="Full Report"
                          ToolTip="Checked, every duct is listed and numbered in the plan, with its total CFM and length. Unchecked, only the ducts flagged red, yellow or purple are listed."/>
                <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,0,0,6"
                           Text="Every duct with its CFM and length. Unchecked lists only the red, yellow and purple ducts."/>

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
                                  ToolTip="Prices the first elbow past each unit connection from the 2001 ASHRAE Fundamentals fan system effect tables and adds it before the safety factor. A tap is priced as an elbow on a run that leaves through it (off a continuing main, off a capped drop, or the tap the unit duct tees into); a run that continues on the main passes it."/>
                        <StackPanel x:Name="pnl_fse" Margin="20,4,0,0">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="Fan type" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                <RadioButton x:Name="rb_fan_swsi" GroupName="fan_type" Content="SWSI" VerticalAlignment="Center" Margin="0,0,14,0"
                                             ToolTip="Outlet tables SR7-5 to SR7-8. Set by the Scope: Exhaust System."/>
                                <RadioButton x:Name="rb_fan_dwdi" GroupName="fan_type" Content="DWDI" VerticalAlignment="Center"
                                             ToolTip="Outlet tables SR7-9 to SR7-12. Set by the Scope: Rooftop Unit and VAV System."/>
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
                                         ToolTip="Fan velocity, fpm, greater than 0. Set by the Scope: 1000 fpm Exhaust System, 1500 fpm Rooftop Unit, VAV System, VAV Box and Indoor System."/>
                                <TextBlock Text="fpm" VerticalAlignment="Center"/>
                            </StackPanel>
                            <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,0"
                                       Text="Set by the Scope: 1000 fpm Exhaust System, 1500 fpm otherwise. The unit connection velocity (cfm over the connector area) is used when it is higher."/>
                        </StackPanel>
                        <!-- Ranked run (Colin, 2026-11-10) -->
                        <CheckBox x:Name="cb_rank" Content="Ranked run" Margin="0,10,0,0"
                                  ToolTip="Prints the take-off of a lesser run (2 = the second most restrictive) as a sanity check. The most restrictive total still prints and stays the design ESP."/>
                        <StackPanel x:Name="pnl_rank" Orientation="Horizontal" Margin="20,4,0,0">
                            <TextBlock Text="Supply rank" VerticalAlignment="Center"/>
                            <TextBox x:Name="tb_rank_sup" Width="56" Margin="8,0,16,0"
                                     ToolTip="Which supply run to print, 2 or higher (2 = the second most restrictive)."/>
                            <TextBlock x:Name="lbl_rank_ret" Text="Return rank" VerticalAlignment="Center"/>
                            <TextBox x:Name="tb_rank_ret" Width="56" Margin="8,0,6,0"
                                     ToolTip="Which return or exhaust run to print, 2 or higher (2 = the second most restrictive)."/>
                        </StackPanel>
                        <TextBlock Style="{DynamicResource RjaHint}" FontSize="11" TextWrapping="Wrap" Margin="20,2,0,0"
                                   Text="2 = the second most restrictive run. The most restrictive total still prints as the design ESP."/>
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
                                      ToolTip="Used where a fitting's geometry or flows cannot be read (flagged FALLBACK in the report), and for offsets and the flex bend, which have no 2001 table lookup. Cannot be negative.">
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


def _scope_of(vav, vavbox, exhaust, fancoil=False):
    """The Scope key from the dialog's choice."""
    if vavbox:
        return 'vavbox'
    if fancoil:
        return 'fancoil'
    return 'vav' if vav else 'exhaust' if exhaust else 'rtu'


def _full_report_columns(full_diag):
    """The optional table columns: Total CFM through Duct and Length with
    Full Report, none otherwise (Colin, 2026-11-12: one Full Report box
    replaces Report all ducts and the two column checkboxes)."""
    return set(['cfm', 'length']) if full_diag else set()


def _settings_result(limits, tol_pct, selected_cols, full_diag, ext_static,
                     safety_pct, c_values, comp_values, calc_basis, filter_inwc,
                     vav, vav_settings, fan_effect=None, exhaust_cap=False,
                     exhaust=False, exhaust_settings=None, scope=None,
                     ranked=None):
    """The settings dialog's answer, a dict, built in ONE place for both the
    themed and the legacy dialog so main() reads the same keys from either.

    With VAV ticked, the MP limits are also added to limits under
    vav_checks.MP_KEY, so every limit lookup finds them in one dict.

    fan_effect: {'on', 'fan_type', 'blast_ratio'}; None (the legacy dialog,
    which has no fields for it) takes fan_system_effect.default_settings().
    exhaust_settings: {'backdraft' (Include backdraft damper)}; a missing
    'backdraft' is taken as off in main() (Colin, 2026-11-03).
    'scope': 'rtu', 'fancoil', 'vav', 'vavbox' or 'exhaust' (Colin, 2026-11-03: one
    Scope choice). None derives it from vav / exhaust (no VAV Box).
    'ranked': None (Ranked run off, the legacy dialog too) or {'supply': n,
    'return': n}, n 2 or higher (Colin, 2026-11-10).
    """
    selected_cols = _full_report_columns(full_diag)
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
        'ranked':        dict(ranked) if ranked else None,
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
    rb_scope_fancoil = N('rb_scope_fancoil')
    cb_curb = N('cb_curb')
    tb_curb = N('tb_curb')
    pnl_exh_static = N('pnl_exh_static')
    cb_exh_backdraft = N('cb_exh_backdraft')
    tb_backdraft = N('tb_backdraft')
    tb_exh_cap = N('tb_exh_cap')
    pnl_vav_static = N('pnl_vav_static')
    grid_vav_static = N('grid_vav_static')
    grid_limits = N('grid_limits')
    tb_tol = N('tb_tol')
    pnl_scope_opts = N('pnl_scope_opts')
    pnl_curb = N('pnl_curb')
    tb_scope_opts_hint = N('tb_scope_opts_hint')
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
    cb_rank = N('cb_rank')
    pnl_rank = N('pnl_rank')
    tb_rank_sup = N('tb_rank_sup')
    tb_rank_ret = N('tb_rank_ret')
    lbl_rank_ret = N('lbl_rank_ret')
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
    tb_curb.Text = '%g' % fan_system_effect.DEFAULT_CURB_IN
    cb_curb.IsChecked = True     # RTU, the default Scope (Colin, 2026-11-09)
    tb_backdraft.Text = '%.2f' % fitting_tables.component_of('backdraft')
    tb_exh_cap.Text = '%.2f' % fitting_tables.component_of('cap')
    tb_rank_sup.Text = str(duct_perf_console.DEFAULT_RANK)
    tb_rank_ret.Text = str(duct_perf_console.DEFAULT_RANK)
    cb_rank.IsChecked = False        # default off on every Scope (Colin, 2026-11-10)
    for b in (tb_tol, tb_sf, tb_filter, tb_blast, tb_fan_fpm, tb_backdraft, tb_exh_cap,
              tb_curb, tb_rank_sup, tb_rank_ret):
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

    # No optional column checkboxes (Colin, 2026-11-12): Full Report adds the
    # CFM and Length columns, see _full_report_columns.
    col_boxes = {}

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
                   'diffuser capacity tables, slots on the slot tables by length and '
                   'number of slots.'),
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
                     'unreadable, and for offsets. Tees, crosses and wyes on the leg the run leaves '
                     'through: the tap tables by shape, SD5-24 for a round supply cross, SR5-14 / ER5-4 '
                     'for a symmetrical split. One flex bend (C 3.3, ASHRAE Duct Systems Design '
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
                            'cap is ticked (never a second cap on a ceiling exhaust fan).'),
        ('Fan system effect', '2001 ASHRAE Fundamentals Eq. 38 (p. 34.14). The first '
                              'disruption past each unit connection: within Le '
                              '(Eq. 36 or 37) of an outlet, within the table floor '
                              'distance (5 D or 5 H) of an inlet. Outlet elbow: SR7 '
                              'tables for the fan type, worst of positions A to D, at '
                              'Ab/Ao and L/Le. Inlet elbow: ED7-2 round, ER7-1 rect '
                              '(SWSI tables, an assumption for DWDI), at r/D and L/D. '
                              'The straight duct, any transition on it and the first '
                              'elbow are charged once, the larger of the fan system '
                              'effect and their own losses. Velocity at an outlet, '
                              'and at an inline or roof exhaust fan inlet: the larger '
                              'of the fan velocity (default by the Scope, 1000 fpm '
                              'Exhaust System, 1500 fpm otherwise) '
                              'and the connection '
                              'velocity; at a cabinet unit inlet (RTU, AHU, furnace, '
                              'fan coil) the connection velocity, the fan sits past '
                              'the filter and coil. Added before the '
                              'safety factor. Transitions are priced as fittings. A '
                              'tap inside the window is priced as an elbow on a run '
                              'that leaves through it (off a continuing main, off a '
                              'capped drop, or the tap the unit duct tees into); a run '
                              'that continues on the main passes it. An elbow under 0.5 Le of an '
                              'outlet, or under 2 H (its side in the plane of the turn; '
                              'D if round) of an inlet, is red with the '
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
    scope_list = rb_scope_rtu.Parent

    def _place_scope_opts():
        """Move the Scope options group to just under the selected Scope's
        radio button and its hint line (Colin, 2026-11-12)."""
        target = rb_scope_rtu
        for rb in (rb_scope_rtu, rb_scope_fancoil, rb_scope_vav, rb_scope_vavbox,
                   rb_scope_exhaust):
            if bool(rb.IsChecked):
                target = rb
        try:
            owner = pnl_scope_opts.Parent
            if owner is not None:
                owner.Children.Remove(pnl_scope_opts)
            idx = scope_list.Children.IndexOf(target)
            scope_list.Children.Insert(idx + 2, pnl_scope_opts)   # after its hint
        except Exception as ex:
            hvac_graph.log.warning('scope options not moved: %s', ex)

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
        tb_curb.IsEnabled = bool(cb_curb.IsEnabled) and bool(cb_curb.IsChecked)
        # Ranked run: Exhaust System has the exhaust side only, VAV Box the
        # supply after the box only (Colin, 2026-11-10).
        rank_on = master and bool(cb_rank.IsChecked)
        pnl_rank.IsEnabled = rank_on
        pnl_rank.Opacity = 1.0 if rank_on else 0.55
        tb_rank_sup.IsEnabled = rank_on and not bool(rb_scope_exhaust.IsChecked)
        tb_rank_ret.IsEnabled = rank_on and not box
        lbl_rank_ret.Text = ('Exhaust rank' if bool(rb_scope_exhaust.IsChecked)
                             else 'Return rank')
        # Scope options sit under the selected Scope (Colin, 2026-11-12):
        # roof curb on RTU and VAV System, VAV boxes on VAV System, backdraft
        # damper and roof / wall cap on Exhaust System, nothing on Indoor
        # System and VAV Box. Always live, external static or not (Colin,
        # 2026-11-12); the drops are only charged when it is ticked.
        _place_scope_opts()
        show_curb = bool(rb_scope_rtu.IsChecked) or bool(rb_scope_vav.IsChecked)
        show_vav = bool(rb_scope_vav.IsChecked)
        show_exh = bool(rb_scope_exhaust.IsChecked)
        pnl_curb.Visibility = Visibility.Visible if show_curb else Visibility.Collapsed
        pnl_vav_static.Visibility = Visibility.Visible if show_vav else Visibility.Collapsed
        pnl_exh_static.Visibility = Visibility.Visible if show_exh else Visibility.Collapsed
        tb_scope_opts_hint.Visibility = (Visibility.Visible if (show_vav or show_exh)
                                         else Visibility.Collapsed)
        pnl_scope_opts.Visibility = (Visibility.Visible if (show_curb or show_vav or show_exh)
                                     else Visibility.Collapsed)
        # An exhaust fan has no return path, so the return filter is not
        # asked for under Exhaust System (Colin, 2026-11-04).
        pnl_filter.Visibility = (Visibility.Collapsed if bool(rb_scope_exhaust.IsChecked)
                                 else Visibility.Visible)
        # The medium pressure supply row shows with VAV System alone.
        vis = Visibility.Visible if bool(rb_scope_vav.IsChecked) else Visibility.Collapsed
        for el in mp_row:
            el.Visibility = vis
    for cb in (cb_static, cb_filter, cb_fse, rb_scope_rtu, rb_scope_vav, rb_scope_vavbox,
               rb_scope_exhaust, rb_scope_fancoil, cb_curb, cb_rank):
        cb.Checked += sync_enable
        cb.Unchecked += sync_enable
    sync_enable()

    # Fan type, fan velocity and roof curb follow the Scope (Colin,
    # 2026-11-03, 2026-11-09 and 2026-11-11): RTU and VAV System DWDI at 1500
    # fpm with the curb on, Indoor System DWDI at 1500 fpm, Exhaust System SWSI
    # at 1000 fpm, the curb off and greyed out on Indoor System, VAV Box and Exhaust
    # System. Set only when the Scope changes, so a value typed after the
    # switch holds for the run.
    def _dialog_scope():
        return _scope_of(bool(rb_scope_vav.IsChecked), bool(rb_scope_vavbox.IsChecked),
                         bool(rb_scope_exhaust.IsChecked), bool(rb_scope_fancoil.IsChecked))

    def _scope_fan_type(s=None, e=None):
        key = _dialog_scope()
        swsi = key == 'exhaust'
        rb_fan_swsi.IsChecked = swsi
        rb_fan_dwdi.IsChecked = not swsi
        tb_fan_fpm.Text = '%g' % fan_system_effect.scope_fan_fpm(key)
        curb = key in fan_system_effect.CURB_SCOPES
        cb_curb.IsEnabled = curb
        cb_curb.IsChecked = curb
        tb_curb.Text = '%g' % fan_system_effect.DEFAULT_CURB_IN
        tb_curb.IsEnabled = curb
    for rb in (rb_scope_rtu, rb_scope_vav, rb_scope_vavbox, rb_scope_exhaust,
               rb_scope_fancoil):
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
        if exhaust:
            # The two drops price modeled dampers and a ceiling fan's own cap
            # too, so they are read whether or not their box is ticked. Always
            # live, external static or not (Colin, 2026-11-12).
            bd = num(tb_backdraft, 'Backdraft damper drop', 'nonneg')
            if bd is not None:
                comp_values['backdraft'] = bd
            cap = num(tb_exh_cap, 'Roof / wall cap drop', 'nonneg')
            if cap is not None:
                comp_values['cap'] = cap

        scope_key = _dialog_scope()
        # The Scope's fan velocity even with external static off: the
        # straight duct RED reads the outlet Le at it (Colin, 2026-11-09).
        fan_effect = fan_system_effect.default_settings(scope_key)
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
        # Roof curb (Colin, 2026-11-09): RTU and VAV System only, read
        # whether or not external static is ticked (the RED checks use it).
        fan_effect['curb_in'] = 0.0
        if bool(cb_curb.IsChecked) and scope_key in fan_system_effect.CURB_SCOPES:
            v, err = fan_system_effect.parse_curb_in(tb_curb.Text)
            if err:
                errors.append((tb_curb, 'Roof curb height ' + err + '.', ()))
            else:
                fan_effect['curb_in'] = v

        vav = bool(rb_scope_vav.IsChecked)
        vav_settings = vav_checks.default_settings()
        if vav:
            fields = [(f, ()) for f in vav_checks.VAV_LIMIT_FIELDS]
            fields += [(f, ()) for f in vav_checks.VAV_STATIC_FIELDS]   # always live (2026-11-12)
            for (key, label, _unit, _default, kind), exps in fields:
                v, err = vav_checks.parse_field(vav_boxes[key].Text, kind)
                if err:
                    errors.append((vav_boxes[key], label + ' ' + err + '.', exps))
                else:
                    vav_settings[key] = v

        # Ranked run (Colin, 2026-11-10): 2 or higher, no fixed top; a rank
        # above the terminal count is a WARNING at run time, not an error here.
        ranked = None
        if ext_static and bool(cb_rank.IsChecked):
            ranked = {'supply': duct_perf_console.DEFAULT_RANK,
                      'return': duct_perf_console.DEFAULT_RANK}
            for key, box, label in (('supply', tb_rank_sup, 'Supply rank'),
                                    ('return', tb_rank_ret, lbl_rank_ret.Text)):
                if not box.IsEnabled:
                    continue
                v, err = duct_perf_console.parse_rank(box.Text)
                if err:
                    errors.append((box, label + ' ' + err + '.', ()))
                else:
                    ranked[key] = v

        selected_cols = set(k for k, cb in col_boxes.items() if bool(cb.IsChecked))
        full_diag = bool(cb_full_diag.IsChecked)
        values = _settings_result(out, gpct, selected_cols, full_diag,
                                  ext_static, safety_pct, c_values, comp_values,
                                  calc_basis, filter_inwc, vav, vav_settings,
                                  fan_effect,
                                  ext_static and exhaust and bool(cb_exh_cap.IsChecked),
                                  exhaust, exhaust_settings,
                                  scope=scope_key, ranked=ranked)
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
        'Full Report.  Every duct with its CFM and length. Unchecked = only '
        'the red, yellow and purple ducts.')
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

    rb_rtu = _scope_radio('Rooftop Unit.  Lists units only, no VAV boxes or exhaust fans. Roof '
                          'curb %g in. counted.' % fan_system_effect.DEFAULT_CURB_IN, True)
    rb_fancoil = _scope_radio('Indoor System.  Lists units only: fan coil, furnace, indoor '
                              'AHU.')
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
    # 2026-10-08: "Include Return Filter", 0.10 known drop (MERV 8, Colin
    # 2026-11-12, was 0.14). On by default
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

    col_boxes = {}   # Full Report sets the optional columns (Colin, 2026-11-12)

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
            fb_scope = _scope_of(bool(rb_vav.IsChecked), bool(rb_vavbox.IsChecked),
                                 bool(rb_exhaust.IsChecked), bool(rb_fancoil.IsChecked))
            result[0] = _settings_result(out, gpct, selected_cols, full_diag,
                                         ext_static, safety_pct, c_values,
                                         comp_values, calc_basis, filter_inwc,
                                         vav, vav_settings,
                                         # no fan fields here: the fan type
                                         # follows the Scope (Colin, 2026-11-03)
                                         fan_effect=dict(
                                             fan_system_effect.default_settings(fb_scope),
                                             fan_type='SWSI' if bool(rb_exhaust.IsChecked)
                                             else 'DWDI',
                                             curb_in=fan_system_effect.scope_curb_in(fb_scope)),
                                         exhaust_cap=ext_static and bool(cb_exh_cap.IsChecked),
                                         exhaust=bool(rb_exhaust.IsChecked),
                                         scope=fb_scope)
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
_SCOPE_LABEL = {'rtu': 'Rooftop Unit', 'fancoil': 'Indoor System', 'vav': 'VAV System',
                'vavbox': 'VAV Box',
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
_SCOPE_LISTS = {'rtu': ('unit',), 'fancoil': ('unit',), 'vav': ('unit',),
                'vavbox': ('VAV box',),
                'exhaust': ('exhaust fan',)}
_KIND_SCOPE = {'unit': 'Rooftop Unit, Indoor System or VAV System', 'VAV box': 'VAV Box',
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
    # 'Select the Rooftop Unit to analyze', not '... Unit unit' (Colin, 2026-11-09)
    prompt = 'Select the {}{} to analyze'.format(
        scope_label, '' if scope_label.lower().endswith('unit') else ' unit')
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
    # Ranked run (Colin, 2026-11-10): None when off.
    ranked_set    = dialog_result.get('ranked') if ext_static else None
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
    scope_label = _SCOPE_LABEL.get(scope, 'Rooftop Unit')
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
        if len(tset) == 1 and hvac_graph.main_end_reaches_tap(
                eid_int(eid), all_nodes, all_children):
            # A main stepped down to its last tap: the branch starts AT the
            # tap, so this piece is still a main, judged on the main limits
            # (Colin, 2026-11-05, TB-M10C 6x10 end piece read GRAY as a
            # rect branch on a round neck). Also when the tap is further down,
            # past an elbow or transition (Colin, 2026-11-11, TB-M18).
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
    # terminal id -> plan mark: 'R', or 'R1' plus 'R2', 'R3' ... for the
    # ranked runs with Ranked run on (Colin, 2026-11-11). A run tied with a
    # marked one at the printed 0.001 in. wc gets the same mark.
    restrictive_terminal_ids = {}

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
                if _kind == 'red' or (_kind == 'warn'
                                      and fan_system_effect.is_check_warning(_text)):
                    # the cabinet inlet 3 D WARNING runs with them (Colin,
                    # 2026-11-09)
                    _correct(_kind, _text, root_labels.get(root_id, 'id {}'.format(root_id)))
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
            inline_fan_roots=set(exhaust_fan_roots) - set(discharge_roots or ()),
            # Every run kept: Ranked run, and a tie at the printed 0.001
            # in. wc names both runs (Colin, 2026-11-11).
            keep_runs=True)
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

            # Ranked run (Colin, 2026-11-10): the take-off shows the Nth run
            # on each side; per_class stays the most restrictive (MR) run,
            # the design ESP.
            shown = {}
            rank_of = {}       # sys_class -> (rank used, run count)
            for sys_class in per_class:
                c0 = per_class[sys_class]
                shown[sys_class] = c0
                if not ranked_set:
                    continue
                want = (ranked_set['supply'] if vav_checks.is_supply_class(sys_class)
                        else ranked_set['return'])
                rec, used, n_runs, over = duct_perf_console.pick_ranked(
                    c0.get('runs') or [c0], want)
                if rec is not None:
                    shown[sys_class] = rec
                    rank_of[sys_class] = (used, n_runs)
                if over:
                    _correct('warn', duct_perf_console.rank_over_text(
                        sys_class.split(' ')[0], want, n_runs), unit_label)

            for sys_class in sorted(per_class.keys()):
                c = shown[sys_class]
                _mr = per_class[sys_class]
                _runs = _mr.get('runs') or [_mr]
                _m1 = 'R1' if ranked_set else 'R'
                for _r in [_mr] + duct_perf_console.tied_runs(_runs, _mr):
                    restrictive_terminal_ids[_r['terminal_id']] = _m1
                _shown_ties = duct_perf_console.tied_runs(_runs, c)
                if sys_class in rank_of:
                    _mn = 'R{}'.format(rank_of[sys_class][0])
                    for _r in [c] + _shown_ties:
                        restrictive_terminal_ids.setdefault(_r['terminal_id'], _mn)
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
                                     'unit connection group, the fan system effect '
                                     '(2001 Fundamentals Eq. 38) or its own duct '
                                     'losses, whichever governs'))
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
                _lbl = lambda r: r.get('terminal_label') or r['terminal_name']
                if sys_class in rank_of:
                    total_det = 'run {} of {} ends at {}'.format(
                        rank_of[sys_class][0], rank_of[sys_class][1], _lbl(c))
                    if _shown_ties:
                        total_det += ', tied with {}'.format(
                            duct_perf_console.and_join(_lbl(r) for r in _shown_ties))
                elif _shown_ties:
                    total_det = 'worst run ends at {}, tied'.format(
                        duct_perf_console.and_join(_lbl(r) for r in [c] + _shown_ties))
                else:
                    total_det = 'worst run ends at {}'.format(_lbl(c))
                # One safety factor on the unit total, not per side (Colin,
                # 2026-11-12): the side ends at its subtotal.
                summary_lines.append(_esp_text_line('{} subtotal'.format(sys_class),
                                                    c['subtotal_inwc'], total_det))
                for u in c['unpriced']:
                    # A FALLBACK note is priced (RJA worksheet C), not missing.
                    # A NOTE (a lookup read at a table edge, tap order
                    # unreadable) lands in Notes, the rest in Action required.
                    _u = (u if (u.startswith('FALLBACK') or u.startswith('NOTE'))
                          else 'UNPRICED: ' + u)
                    _correct('warn', _u, unit_label)
                # The take-off: every item on the run in route order from the
                # unit, one row each, straight duct between merged (Colin,
                # 2026-11-11). Its rows add up to the subtotal less the filter.
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
                    'rank': rank_of[sys_class][0] if sys_class in rank_of else None,
                    'rank_text': ('run {} of {}'.format(*rank_of[sys_class])
                                  if sys_class in rank_of else None),
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
            # The run NOTE for each connection the shown run leaves through:
            # per class, since supply and return leave through different
            # connections. A ranked take-off says its own run, 'run N'
            # (Colin, 2026-11-11).
            for _cls in sorted(per_class.keys()):
                _shown = shown.get(_cls) or per_class[_cls]
                for _kind, _text in _shown.get('fse_gov_notes', ()):
                    if _cls in rank_of:
                        _text = _text.replace(fan_system_effect.MR_RUN,
                                              'run {}'.format(rank_of[_cls][0]), 1)
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

            # unit_total: the runs the take-off shows (the MR runs, or the
            # ranked runs with Ranked run on). mr_total: always the MR runs,
            # the design ESP (Colin, 2026-11-10).
            unit_total = 0.0
            mr_total = 0.0
            sides = []
            mr_parts = []
            rk_parts = []

            def _mr_part(side, leaf):
                _ties = duct_perf_console.tied_runs(leaf.get('runs') or [leaf], leaf)
                _names = duct_perf_console.and_join(
                    r.get('terminal_label') or r['terminal_name'] for r in [leaf] + _ties)
                # Subtotal: the safety factor is applied once at the end.
                return '{} {:.3f} at {}{}'.format(
                    side, float(leaf['subtotal_inwc']), _names, ', tied' if _ties else '')
            if supply is not None:
                unit_total += shown['Supply Air']['total_inwc']
                mr_total += supply['total_inwc']
                sides.append('Supply Air')
                mr_parts.append(_mr_part('Supply Air', supply))
                rk_parts.append('Supply Air #{}'.format(
                    rank_of.get('Supply Air', (1, 0))[0]))
            if returns:
                # The larger return side by its MR run; the same side is
                # shown ranked, so the two totals never mix sides.
                worst_k, worst_c = max(returns, key=lambda kv: kv[1]['total_inwc'])
                unit_total += shown[worst_k]['total_inwc']
                mr_total += worst_c['total_inwc']
                sides.append(worst_k)
                mr_parts.append(_mr_part(worst_k, worst_c))
                rk_parts.append('{} #{}'.format(worst_k, rank_of.get(worst_k, (1, 0))[0]))
            if sides:
                unit_rec['mr_text'] = '{}, plus {:.0f}%'.format(
                    ' + '.join(mr_parts), float(safety_pct))
                # Plain text: the combined subtotal and the one safety factor
                # (Colin, 2026-11-12), printed before the total line below.
                _sub_all = mr_total / (1.0 + float(safety_pct) / 100.0)
                summary_lines.append(_esp_text_line('Subtotal', _sub_all,
                                                    ' + '.join(sides)))
                summary_lines.append(_esp_text_line('Safety factor', mr_total - _sub_all,
                                                    '+{:.0f}%'.format(float(safety_pct))))
            if ranked_set and sides:
                unit_rec['ranked'] = True
                unit_rec['ranked_text'] = ' + '.join(rk_parts)
                unit_rec['mr_total'] = mr_total
                summary_lines.append('    {} = {:.3f} in. wc   ({}, {})'.format(
                    duct_perf_console.RANKED_TOTAL_LABEL.upper(), unit_total,
                    unit_rec['ranked_text'], duct_perf_console.RANKED_TOTAL_NOTE))
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
                        else 'EXTERNAL STATIC PRESSURE (MOST RESTRICTIVE)', mr_total, _sides))
                if _is_box:
                    summary_lines.append('    ' + duct_perf_console.VAVBOX_ESP_CAUTION)
                    unit_rec['caution'] = duct_perf_console.VAVBOX_ESP_CAUTION
                unit_rec['esp'] = _esp_console_text(mr_total, _sides)
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
                    '    EXTERNAL STATIC PRESSURE (MOST RESTRICTIVE) = {:.3f} in. wc   ({})'.format(
                        mr_total, ' + '.join(sides)))
                unit_rec['esp'] = _esp_console_text(mr_total, ' + '.join(sides))
                unit_rec['esp_total'] = unit_total
                unit_rec['esp_sides'] = list(sides)
            elif sides:
                # One side only. Held back, not warned about yet: whether it is
                # half of a system depends on what else was selected.
                one_sided.append((root_id, mr_total, sides, supply is not None))
                unit_rec['esp_sides'] = list(sides)
            if (vav_on and sides and vav_set.get('unit_max_esp') is not None
                    and root_id in vav_by_root and vav_by_root[root_id].is_vav):
                _umax = float(vav_set['unit_max_esp'])
                _ok = mr_total <= _umax + 1e-9
                _msg = ('Unit max external static {:.3f} in. wc: computed {:.3f} in. wc '
                        '({}), {}.'.format(_umax, mr_total, ' + '.join(sides),
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
                   'Eq. 34, ER5-2 and ER5-3 at their printed Ab/Ac 0.5, conservative for a larger '
                   'tap, a smaller one noted). Unbroken '
                   'main: As/Ac read on the 0.9 row. A tap C of 0 or less from the '
                   '2001 tables is replaced by the RJA worksheet C (FALLBACK). The '
                   '2001 book has no round shoe (boot) table on a round main; ED5-2 / '
                   'SD5-1 (45 deg wye) stand in for it. Every tap on a rect supply main is '
                   'a 45 deg entry (RJA standard), priced on SR5-13; a round tap is '
                   'read at its own area, SR5-13 standing in. RJA SP_LOSS_WORKSHEET C where '
                   'geometry is unreadable (FALLBACK).'.format(
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
                         else ' Exhaust roof / wall cap not included (Include roof / '
                              'wall cap unticked).')
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
    # A level with no floor plan (a leftover or a ceiling level) is drawn on
    # the plan level at its elevation, else the nearest below (Colin,
    # 2026-11-11, TB-M12 'LEVEL 0 - DELETE BEFORE COPYING').
    _plan_lvls, _plan_lvl_ids = _plan_levels(sheets_by_view)
    _plan_ids = set(l for l, _e in _plan_lvls)
    _canon_cache = {}

    def _draw_level_id(elem):
        lid = _elem_level_id(elem)
        if lid is None:
            return None
        key = eid_int(lid)
        if key in _plan_ids:
            return lid
        if key not in _canon_cache:
            pick = None
            try:
                pick = _canon_level_pick(float(doc.GetElement(lid).Elevation), _plan_lvls)
            except Exception as ex:
                hvac_graph.log.warning('level: %s elevation NOT READ: %s', key, ex)
            _canon_cache[key] = _plan_lvl_ids.get(pick, lid) if pick is not None else lid
            hvac_graph.log.info('level: %s drawn on level %s', key, pick)
        return _canon_cache[key]
    # The levels the system touches: every duct (its Reference Level) and
    # every terminal. A level holding only the unit (an RTU on a roof deck)
    # gets no plan; the unit is colored in whichever plan shows it (Colin,
    # 2026-11-09, TB-M12C "LEVEL 2 - DECK").
    level_ids = {}
    level_elems = {}          # int level id -> [elements drawn on it]: crop and plan pick
    for eid, dr in all_duct_results.items():
        lid = _draw_level_id(dr.elem)
        if lid is not None:
            level_ids[eid_int(lid)] = lid
    for _nid in all_terminals.keys():
        _el = all_nodes.get(_nid)
        lid = _draw_level_id(_el) if _el is not None else None
        if lid is not None:
            level_ids[eid_int(lid)] = lid
    for _nid, _el in all_nodes.items():
        lid = _draw_level_id(_el) if _el is not None else None
        if lid is not None:
            level_elems.setdefault(eid_int(lid), []).append(_el)
    # Each level's system as a model XY rectangle: a source plan's crop
    # shall overlap it (Colin, 2026-11-09, TB-M12C took TB-M01C's plan).
    level_rects = {}
    for _lid in level_ids:
        level_rects[_lid] = _crop_box_xy(_elem_box_points(level_elems.get(_lid, [])), 0.0)
    level_targets, missing_levels = _target_plan_views(active_view, level_ids, sheets_by_view,
                                                       level_rects)

    # One colored plan per level, all on one sheet (Colin, 2026-11-09,
    # replacing the 2026-10-16 single plan with its view range opened over
    # every level, where the floors drew on top of each other). Each copy
    # keeps its source plan's own view range and template and is cropped to
    # that level's system (5 ft margin). A level with no mechanical plan gets
    # one made from the active view's type and template, inside the
    # transaction. A single level system is unchanged: the active view only.
    # Multi level: the system is on more than one level, or on one level that
    # is not the active view's (that level's own plan is used).
    try:
        _act_lid = eid_int(active_view.GenLevel.Id)
    except Exception:
        _act_lid = None
    multi_level = len(level_ids) > 1 or (len(level_ids) == 1 and _act_lid not in level_ids)
    if multi_level:
        # Listed lowest level first, as the sheet stacks them (Colin,
        # 2026-11-11: printed LEVEL 3, LEVEL 1, LEVEL 2).
        def _t_elev(t):
            try:
                return float(doc.GetElement(t[0]).Elevation)
            except Exception:
                return 0.0
        targets = sorted(level_targets, key=_t_elev)
    else:
        targets = [(None, active_view, sheets_by_view.get(eid_int(active_view.Id)))]
    _plan_txt = ', '.join(
        '{} ({})'.format(_level_name(doc.GetElement(_l)) if _l is not None else _view_name(_v),
                         _s.SheetNumber if _s is not None else 'no sheet')
        for _l, _v, _s in targets)
    if multi_level and missing_levels:
        _plan_txt += ', ' + ', '.join('{} (new plan)'.format(_level_name(doc.GetElement(_l)))
                                      for _l in missing_levels)
    output.print_md('Plans: **{}**'.format(_plan_txt))

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
        gray_notes = []
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
                    label, reason = 'GREEN', MEETS_STANDARD_REASON
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
            if label == 'GRAY':
                # Say which duct and why it was not judged (Colin, 2026-11-11):
                # the legend alone read 'no airflow' for a rect duct on a
                # round neck.
                gray_notes.append((dr.element_id,
                                   'NOTE: duct {} ({}, {}) not judged: {}.'.format(
                                       dr.element_id, dr.sys_class or '-', dr.size or '-',
                                       reason or 'Cannot check')))

        # A clearance NOTE asks for a resize: Action required, named by the
        # first picked unit whose ductwork holds the duct.
        def _unit_of(nid):
            for _rid in all_root_ids:
                if nid in nodes_by_root.get(_rid, ()):
                    return root_labels.get(_rid)
            return None
        for _did, _n in sorted(clearance_notes, key=lambda dn: dn[1]):
            _correct('warn', _n, _unit_of(_did), duct_perf_console.ACTION)
        for _did, _n in sorted(gray_notes):
            _correct('note', _n, _unit_of(_did))

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
        ogs_bold_by_label = {}
        for lbl_key, color in _COLOR_MAP.items():
            for bold in (False, True):
                ogs = OverrideGraphicSettings()
                ogs.SetSurfaceForegroundPatternColor(color)
                if fill_id != ElementId.InvalidElementId:
                    ogs.SetSurfaceForegroundPatternId(fill_id)
                # The most restrictive run: an orange outline over its label
                # fill, so it reads even with Thin Lines on (Colin,
                # 2026-11-11: weight alone did not show).
                ogs.SetProjectionLineColor(MR_LINE_COLOR if bold else color)
                if bold:
                    try:
                        ogs.SetProjectionLineWeight(MR_LINE_WEIGHT)
                    except Exception as ex:
                        hvac_graph.log.warning('bold run: line weight %s NOT SET: %s',
                                               MR_LINE_WEIGHT, ex)
                (ogs_bold_by_label if bold else ogs_by_label)[lbl_key] = ogs

        # The most restrictive run, unit to its (R) / (R1) terminal, drawn with
        # a bold outline so its routing can be followed (Colin, 2026-11-11).
        _key_of = dict((eid_int(k), k) for k in elem_labels)
        _up = {}
        for _p, _cs in all_children.items():
            for _c in _cs:
                _up.setdefault(_c, _p)
        bold_ids = set()
        for _tid, _mk in restrictive_terminal_ids.items():
            if _mk not in ('R', 'R1'):
                continue
            _n, _guard = _tid, 0
            while _n is not None and _guard < 5000:
                if _n in _key_of:
                    bold_ids.add(_key_of[_n])
                if _n in all_root_ids:
                    break
                _n, _guard = _up.get(_n), _guard + 1
        hvac_graph.log.info('bold run: %d elements on the most restrictive run(s)',
                            len(bold_ids))

        # A marker is only placed in the plan for the level its element is
        # on, so a main level duct never gets a circle drawn on the roof plan.
        def _on_level(elem, lvl_eid):
            if lvl_eid is None:
                return True
            lid = _draw_level_id(elem)
            return lid is None or lid == lvl_eid

        new_sheets = []
        used_keys  = set()
        made_ids   = set()        # int ids of level plans made below
        # Multi level: a level with no mechanical plan gets one made here,
        # before anything is drawn, so its NOTE is on the sheet (Colin,
        # 2026-11-09). It is the copy itself (named like the others), not a
        # source left in the model. Then lowest level first.
        if multi_level:
            for _lid in missing_levels:
                _made = _create_level_plan(active_view, _lid)
                _lname = _level_name(doc.GetElement(_lid))
                if _made is None:
                    _correct('warn', 'WARNING: no mechanical floor plan on {} and one could '
                             'not be created; its ductwork is colored but not shown on '
                             'the sheet.'.format(_lname))
                    continue
                _correct('note', 'NOTE: no mechanical floor plan on {}; one was created '
                         'from the active view\'s type and template for this sheet.'
                         .format(_lname))
                made_ids.add(eid_int(_made.Id))
                targets.append((_lid, _made, None))

            def _lvl_elev(item):
                try:
                    return doc.GetElement(item[0]).Elevation
                except Exception:
                    return float('inf')
            targets.sort(key=_lvl_elev)
        act_sheet = sheets_by_view.get(eid_int(active_view.Id))
        if multi_level and not targets:
            # No level plan found or made: the active view alone, as for a
            # single level system, so the run still gives a sheet.
            multi_level = False
            targets = [(None, active_view, act_sheet)]
        multi_sheet = None        # the one sheet every level is placed on
        multi_vps = []            # (viewport, level name) placed on it, in order
        for lvl_eid, src_view, src_sheet in targets:
            made_here = eid_int(src_view.Id) in made_ids
            if multi_level:
                # One sheet named after the active view's sheet; each plan
                # named with its level.
                source_sheet_num = act_sheet.SheetNumber if act_sheet is not None else 'NoSheet'
                lvl_txt = (_level_name(doc.GetElement(lvl_eid)) if lvl_eid is not None
                           else _view_name(src_view))
                view_key = '{}-{}'.format(source_sheet_num, lvl_txt)
            else:
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
                view_key = source_sheet_num
            mark_off_ft = (5.0 / (64.0 * 12.0)) * float(getattr(src_view, 'Scale', 48)) * 2.0

            # Copy floor plan (a plan made above for a level is used as is)
            if made_here:
                new_view = src_view
                new_vid = new_view.Id
            else:
                new_vid  = src_view.Duplicate(ViewDuplicateOption.Duplicate)
                new_view = doc.GetElement(new_vid)
            base_name = 'Duct Performance - {} - {}'.format(view_key, ts)
            try:
                new_view.Name = base_name
            except Exception:
                new_view.Name = base_name + ' (2)'

            if multi_level and lvl_eid is not None:
                _crop_to_elements(new_view, level_elems.get(eid_int(lvl_eid), []), 5.0)
                _fit_level_view_scale(new_view)
                mark_off_ft = (5.0 / (64.0 * 12.0)) * float(getattr(new_view, 'Scale', 48)) * 2.0

            for oid, lbl_key in elem_labels.items():
                new_view.SetElementOverrides(
                    oid, (ogs_bold_by_label if oid in bold_ids else ogs_by_label)[lbl_key])

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
                r_mark = restrictive_terminal_ids[tid]
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
                            num_param.Set(r_mark)
                            placed = True
                        else:
                            doc.Delete(inst.Id)
                    if not placed and tn_type_id is not None:
                        opts = TextNoteOptions(tn_type_id)
                        TextNote.Create(doc, new_vid, mark_pt, '({})'.format(r_mark), opts)
                except Exception:
                    pass

            # Multi level: every level after the first goes on the one sheet,
            # to the right of the one before it; overlaps are fine, the
            # engineer moves them (Colin, 2026-11-09).
            if multi_level and multi_sheet is not None:
                _vp = Viewport.Create(doc, multi_sheet.Id, new_vid, XYZ(0, 0, 0))
                rja_drafting.set_viewport_type_none(doc, _vp)
                multi_vps.append((_vp, lvl_txt))
                continue

            # Output sheet
            new_sheet             = ViewSheet.Create(doc, tb_id)
            new_sheet.SheetNumber = 'DV-{}-{}'.format(source_sheet_num, ts)
            # Sheet/View names share one uniqueness pool in Revit — new_view above
            # already claimed base_name verbatim, so reusing it here always collides.
            new_sheet.Name = ('Duct Performance - {} - {}'.format(source_sheet_num, ts)
                              if multi_level else base_name) + ' (Sheet)'
            _file_like_source_sheet(new_sheet, act_sheet if multi_level else src_sheet)
            new_sheets.append(new_sheet)
            if multi_level:
                multi_sheet = new_sheet

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
            if multi_level:
                multi_vps.append((diagram_vp, lvl_txt))

            # System Summary + flagged-duct table + legend, placed as second
            # viewport on every sheet (the table covers all levels).
            if tn_type_id is not None:
                sched_view, content_h, total_w = _build_summary_view(
                    doc, rep, esp_rec, vav_rep, [item[2] for item in flagged_items],
                    selected_cols, source_sheet_num, tn_type_id, ts, fill_id,
                    show_r_marker_note=set(restrictive_terminal_ids.values()),
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

        # Multi level: lay the level plans out left to right from the first
        # one's fixed center, each labeled with its level name above it.
        if multi_level and multi_sheet is not None and multi_vps:
            _layout_level_viewports(multi_sheet, multi_vps, tn_type_id)

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
         'Main: meets the design standard. Branch: {} and duct both correctly sized'.format(words['one'])],
        ['Purple', str(counts.get('PURPLE', 0)), str(fitting_counts.get('PURPLE', 0)),
         'Main only: reduced duct size available (a smaller standard size still meets the design standard)'],
        ['Yellow', str(counts.get('YELLOW', 0)), str(fitting_counts.get('YELLOW', 0)),
         'Main only: above the max standard, within the tolerance'],
        ['Red', str(counts.get('RED', 0)), str(fitting_counts.get('RED', 0)),
         'Main: high velocity or pressure loss, over the max standard plus tolerance. Branch: {} or duct undersized'.format(words['one'])],
        ['Gray', str(counts.get('GRAY', 0)), str(fitting_counts.get('GRAY', 0)),
         'Not judged: no CFM reaches it (broken or disconnected system), or a check '
         'could not be read (a NOTE names the duct and why). Not in the flagged table'],
    ]
    legend_rec = {
        'roles_note': ('Mains and branches were checked by different methods.',
                       'A duct is a branch when exactly one terminal is reachable '
                       'downstream of it and no tap lies between it and that terminal, '
                       'a main otherwise (the end of a main runs to its last tap).'),
        'limits_caption': ('Main ducts, design limits used (green at or under max, '
                           'yellow within {}% above max, red over max + {}%, purple '
                           'when a smaller standard size still fits)'.format(
                               int(tol_pct), int(tol_pct))),
        'limits': limit_rows,
        'vav_basis': vav_basis,
        # Terminal wording (Colin, 2026-11-09), the auto-sizing diffuser and
        # slot notes left out on Exhaust System, where they cannot apply.
        'notes': duct_perf_console.branch_notes(words),
        'legend': legend_rows,
        'words': words,
        'clearance': (clearance_count > 0,
                      duct_perf_console.clearance_text(clearance_count)),
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
