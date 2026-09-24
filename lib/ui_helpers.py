# -*- coding: ascii -*-
# ui_helpers.py
# Shared WPF dialog helpers for Gas Sizing buttons.
# IronPython 2.7

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows.Markup import XamlReader
from System.Windows.Media import Brushes

import gas_tables
import shared_params
import water_tables


_PICKER_XAML = (
    '<Window'
    ' xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
    ' Height="310" Width="460"'
    ' ResizeMode="NoResize"'
    ' WindowStartupLocation="CenterScreen">'
    '<StackPanel Margin="15">'
    '<TextBlock Text="Pipe Material" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<ComboBox Name="cbMaterial" Margin="0,0,0,12"/>'
    '<TextBlock Text="IFGC Table" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<ComboBox Name="cbTable" Margin="0,0,0,12"/>'
    '<TextBlock Text="Heat Content of Gas (BTU/CF)" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<TextBox Name="tbHeatContent" Margin="0,0,0,4"/>'
    '<TextBlock Text="RJA standard: CFH = BTUH / Heat Content of Gas. Get this value from the'
    ' utility (default: Denver, 840 BTU/CF). Sea level = 1000 BTU/CF."'
    ' FontStyle="Italic" FontSize="10" Foreground="Gray" TextWrapping="Wrap" Margin="0,0,0,16"/>'
    '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">'
    '<Button Name="btnCancel" Content="Cancel" Width="70" Margin="0,0,8,0"/>'
    '<Button Name="btnOK" Content="OK" Width="70"/>'
    '</StackPanel>'
    '</StackPanel>'
    '</Window>'
)


def show_water_dialog(title, project_info, return_detection):
    """The one Size Water dialog, laid out like the Duct Velocity dialog.

    Same idiom as show_velocity_settings_dialog(): options first, bold grey
    section headers, separators between sections, and the basis of design as
    label / value rows at the BOTTOM, directly above OK / Cancel.

    Two check boxes say what the run does. Sizing is always on and shown
    disabled, so the dialog states the job rather than offering it. WSFU
    Calculations is optional and produces the take-off.

    The hot water RETURN is NOT asked about. It is detected from the pipes'
    System Types before this dialog opens, because System Type separates the
    recirculation system from the hot supply even though system
    CLASSIFICATION does not. The finding is printed to the pyRevit window and
    restated here as an assumption, so there is nothing to ask.

    Args:
        title: window title string.
        project_info: dict with "job", "job_number", "by" defaults, read from
            Revit Project Information by the caller.
        return_detection: the dict from
            water_graph.detect_return_system_types(). Reported, not asked.

    Returns:
        dict with "job", "job_number", "by", "wsfu_calcs" (bool) and
        "return_system_type_ids" (set of ints), or None if cancelled.
    """
    from System.Windows import (
        Thickness, TextWrapping, SizeToContent, WindowStartupLocation,
        HorizontalAlignment, FontWeights, Window)
    from System.Windows.Controls import (
        StackPanel, TextBlock, TextBox, CheckBox, Button, Label, Separator,
        Orientation)
    from System.Windows.Media import SolidColorBrush, Colors

    WIN_WIDTH = 520
    CONTENT_W = WIN_WIDTH - 28 - 20

    result = [None]

    win = Window()
    win.Title = title
    win.Width = WIN_WIDTH
    win.SizeToContent = SizeToContent.Height
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    outer = StackPanel()
    outer.Margin = Thickness(14)

    def _section(text):
        hdr = TextBlock()
        hdr.Text = text
        hdr.FontWeight = FontWeights.Bold
        hdr.Foreground = SolidColorBrush(Colors.DimGray)
        hdr.Margin = Thickness(0, 0, 0, 4)
        outer.Children.Add(hdr)

    def _separator():
        sep = Separator()
        sep.Margin = Thickness(0, 12, 0, 8)
        outer.Children.Add(sep)

    def _note(text, indent=20):
        tb = TextBlock()
        tb.Text = text
        tb.TextWrapping = TextWrapping.Wrap
        tb.Width = CONTENT_W - indent
        tb.Foreground = SolidColorBrush(Colors.DimGray)
        tb.Margin = Thickness(indent, 0, 0, 8)
        outer.Children.Add(tb)

    def _checkbox(text, checked, enabled=True, bold=False):
        box = CheckBox()
        caption = TextBlock()
        caption.Text = text
        caption.TextWrapping = TextWrapping.Wrap
        caption.Width = CONTENT_W - 20
        box.Content = caption
        box.IsChecked = checked
        box.IsEnabled = enabled
        if bold:
            box.FontWeight = FontWeights.Bold
        box.Margin = Thickness(2, 0, 0, 2)
        outer.Children.Add(box)
        return box

    def _field(label_text, value):
        lbl = Label()
        lbl.Content = label_text
        outer.Children.Add(lbl)
        tb = TextBox()
        tb.Text = value or ""
        tb.Margin = Thickness(2, 0, 2, 8)
        outer.Children.Add(tb)
        return tb

    # -- 1. what this run does ----------------------------------------------
    _section("This Run Will")
    _checkbox("Size the piping and write the sizes into the model",
              True, enabled=False, bold=True)
    _note("Always on. Cold water is sized on TOTAL fixture units and hot "
          "water on HOT fixture units. Drawn sizes are overwritten and the "
          "fittings are resized to match.")
    cb_wsfu = _checkbox("WSFU Calculations", True)
    _note("Puts the water supply fixture unit take-off on a drafting view "
          "and a new sheet, so the fixture unit count behind every pipe size "
          "can be checked against the model. The take-off prints in the "
          "pyRevit window either way.")

    _separator()

    # -- 2. project information ---------------------------------------------
    _section("Project Information")
    tb_job = _field("Job name", project_info.get("job", ""))
    tb_job_no = _field("Job number", project_info.get("job_number", ""))
    tb_by = _field("By", project_info.get("by", ""))

    _separator()

    # -- 3. basis of design, at the bottom, matching Calculation Basis ------
    _section("Basis of Design")

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

    # Straight from the data module, so the dialog can never show a basis
    # different from the one the sizing actually used.
    for label_text, value_text in water_tables.basis_of_design_rows():
        _info_row(label_text, value_text)

    # The detected return system is NOT shown here. The basis block's
    # "Not included" row already says recirculation sizing is still to do,
    # and the pyRevit window prints which System Type was identified and why.
    # Carried through to the result so those pipes are skipped by the engine.
    detected = set((return_detection or {}).get("detected") or set())

    # -- OK / Cancel ---------------------------------------------------------
    btn_panel = StackPanel()
    btn_panel.Orientation = Orientation.Horizontal
    btn_panel.HorizontalAlignment = HorizontalAlignment.Right
    btn_panel.Margin = Thickness(0, 14, 0, 0)

    ok_btn = Button()
    ok_btn.Content = "OK"
    ok_btn.Width = 72
    ok_btn.Margin = Thickness(0, 0, 8, 0)

    cancel_btn = Button()
    cancel_btn.Content = "Cancel"
    cancel_btn.Width = 72

    def on_ok(sender, e):
        result[0] = {
            "job": tb_job.Text.strip(),
            "job_number": tb_job_no.Text.strip(),
            "by": tb_by.Text.strip(),
            "wsfu_calcs": bool(cb_wsfu.IsChecked),
            # Detected, not chosen. The dialog never asked.
            "return_system_type_ids": set(detected),
        }
        win.Close()

    def on_cancel(sender, e):
        win.Close()

    ok_btn.Click += on_ok
    cancel_btn.Click += on_cancel
    btn_panel.Children.Add(ok_btn)
    btn_panel.Children.Add(cancel_btn)
    outer.Children.Add(btn_panel)

    win.Content = outer
    win.ShowDialog()

    return result[0]


def show_table_picker(title):
    """Single WPF window with linked material/table dropdowns plus a Heat
    Content of Gas field (used for the RJA MBH->CFH conversion - see
    sizing_engine.mbh_to_cfh()).

    Selecting a material instantly repopulates the table list.

    Args:
        title: Window title string.

    Returns:
        (pipe_material, short_table_label, heat_content_btu_per_cf) or
        (None, None, None) if cancelled.
    """
    window          = XamlReader.Parse(_PICKER_XAML)
    window.Title    = title
    cb_material     = window.FindName('cbMaterial')
    cb_table        = window.FindName('cbTable')
    tb_heat_content = window.FindName('tbHeatContent')
    btn_ok          = window.FindName('btnOK')
    btn_cancel      = window.FindName('btnCancel')

    materials = gas_tables.get_material_labels()
    for m in materials:
        cb_material.Items.Add(m)
    cb_material.SelectedIndex = 0

    def populate_table(mat):
        cb_table.Items.Clear()
        for lbl in gas_tables.get_table_option_labels_for_material(mat):
            cb_table.Items.Add(lbl)
        if cb_table.Items.Count > 0:
            cb_table.SelectedIndex = 0

    populate_table(materials[0])
    tb_heat_content.Text = str(int(shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF))

    def on_material_changed(sender, e):
        if cb_material.SelectedItem is not None:
            populate_table(cb_material.SelectedItem)

    cb_material.SelectionChanged += on_material_changed

    result = [None, None, None]

    def on_ok(sender, e):
        try:
            heat_content = float(tb_heat_content.Text.strip())
            if heat_content <= 0:
                raise ValueError
        except ValueError:
            tb_heat_content.Background = Brushes.LightPink
            return
        tb_heat_content.Background = Brushes.White
        result[0] = cb_material.SelectedItem
        result[1] = cb_table.SelectedItem
        result[2] = heat_content
        window.Close()

    def on_cancel(sender, e):
        window.Close()

    btn_ok.Click     += on_ok
    btn_cancel.Click += on_cancel
    window.ShowDialog()

    return result[0], result[1], result[2]
