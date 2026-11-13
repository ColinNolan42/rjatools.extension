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
# Dialog defaults, the dialog answer and the Scope tables live in
# lib/duct_perf_run.py (2026-11-13) so the headless entry builds the same
# answer the dialog does.
from duct_perf_run import (
    _DV_ROWS, _DV_DEFAULT_TOL_PCT, _DV_DEFAULT_SAFETY_PCT, _scope_of,
    _full_report_columns, _settings_result, _SCOPE_LABEL,
    _EXHAUST_FAN_KIND_TEXT, _SCOPE_LISTS, _KIND_SCOPE, _equip_kind,
)

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
    win = ui_helpers.rja_wpf_window(xaml)

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
    win = ui_helpers.rja_wpf_window(xaml)

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
    # Friction basis set and printed before the equipment list, as always;
    # the run itself is duct_perf_run.perf_core (headless step 4b).
    duct_perf_run.apply_calc_basis(dialog_result)
    exhaust_on    = dialog_result.get('exhaust', False)



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

    res = duct_perf_run.perf_core(dialog_result, active_view, sel_elems,
                                  fan_inputs, make_sheets=True,
                                  apply_basis=False)

    if res['new_sheets']:
        uidoc.ActiveView = res['new_sheets'][0]

main()
