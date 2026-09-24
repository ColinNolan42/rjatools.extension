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


_WATER_XAML = (
    '<Window'
    ' xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
    ' Height="600" Width="560"'
    ' ResizeMode="NoResize"'
    ' WindowStartupLocation="CenterScreen">'
    '<Grid Margin="15">'
    '<Grid.RowDefinitions>'
    '<RowDefinition Height="Auto"/>'
    '<RowDefinition Height="*"/>'
    '<RowDefinition Height="Auto"/>'
    '</Grid.RowDefinitions>'

    '<StackPanel Grid.Row="0">'
    '<TextBlock Text="Basis of Design" FontWeight="Bold" Margin="0,0,0,4"/>'
    '<Border BorderBrush="#CCCCCC" BorderThickness="1" Background="#F7F7F7"'
    ' Padding="8" Margin="0,0,0,12">'
    '<TextBlock Name="tbBasis" TextWrapping="Wrap" FontSize="11"'
    ' Foreground="#333333"/>'
    '</Border>'
    '</StackPanel>'

    '<ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">'
    '<StackPanel>'
    '<TextBlock Text="Project Information" FontWeight="Bold" Margin="0,0,0,6"/>'
    '<TextBlock Text="Job name" FontSize="11" Margin="0,0,0,2"/>'
    '<TextBox Name="tbJob" Margin="0,0,0,8"/>'
    '<TextBlock Text="Job number" FontSize="11" Margin="0,0,0,2"/>'
    '<TextBox Name="tbJobNo" Margin="0,0,0,8"/>'
    '<TextBlock Text="By" FontSize="11" Margin="0,0,0,2"/>'
    '<TextBox Name="tbBy" Margin="0,0,0,14"/>'

    '<TextBlock Text="This Run Will" FontWeight="Bold" Margin="0,0,0,6"/>'
    '<CheckBox Name="cbSizing" IsChecked="True" IsEnabled="False"'
    ' Margin="0,0,0,2" Content="Size the piping and write the sizes into the'
    ' model"/>'
    '<TextBlock FontSize="10" Foreground="Gray" TextWrapping="Wrap"'
    ' Margin="20,0,0,8"'
    ' Text="Always on. Sizes domestic cold water on TOTAL fixture units and'
    ' domestic hot water on HOT fixture units, overwrites the drawn sizes and'
    ' resizes the fittings to match."/>'
    '<CheckBox Name="cbWsfu" IsChecked="True" Margin="0,0,0,2"'
    ' Content="WSFU Calculations"/>'
    '<TextBlock FontSize="10" Foreground="Gray" TextWrapping="Wrap"'
    ' Margin="20,0,0,14"'
    ' Text="Puts the water supply fixture unit take-off on a drafting view and'
    ' a new sheet, so the fixture unit count behind every pipe size can be'
    ' checked against the model. The take-off also prints in the pyRevit'
    ' window either way."/>'

    '<TextBlock Text="Hot Water Return System" FontWeight="Bold"'
    ' Margin="0,0,0,4"/>'
    '<TextBlock Name="tbReturnNote" FontSize="10" Foreground="Gray"'
    ' TextWrapping="Wrap" Margin="0,0,0,6"/>'
    '<StackPanel Name="spHotSystems" Margin="8,0,0,14"/>'
    '</StackPanel>'
    '</ScrollViewer>'

    '<StackPanel Grid.Row="2" Orientation="Horizontal"'
    ' HorizontalAlignment="Right" Margin="0,12,0,0">'
    '<Button Name="btnCancel" Content="Cancel" Width="80" Margin="0,0,8,0"/>'
    '<Button Name="btnSize" Content="Size Water" Width="110"/>'
    '</StackPanel>'
    '</Grid>'
    '</Window>'
)


def show_water_dialog(title, project_info, return_detection):
    """The one Size Water dialog. One dialog, then everything is automatic.

    Two things happen on a run, and the check boxes say which: sizing, which
    is always on, and the WSFU Calculations take-off, which is optional.

    The hot water RETURN is DETECTED, not asked. It is read off the pipes'
    System Types, which do separate the recirculation system from the hot
    supply even though the system CLASSIFICATION does not. The detected
    system is shown already ticked, with the reason it was picked, so the
    dialog reports a finding the user can overrule instead of asking a
    question the model already answers.

    Args:
        title: window title string.
        project_info: dict with "job", "job_number", "by" defaults, read from
            Revit Project Information by the caller.
        return_detection: the dict from
            water_graph.detect_return_system_types().

    Returns:
        dict with "job", "job_number", "by", "wsfu_calcs" (bool) and
        "return_system_type_ids" (set of ints), or None if cancelled.
    """
    from System.Windows.Controls import CheckBox, TextBlock
    from System.Windows import Thickness, TextWrapping

    window = XamlReader.Parse(_WATER_XAML)
    window.Title = title

    tb_basis = window.FindName('tbBasis')
    tb_job = window.FindName('tbJob')
    tb_job_no = window.FindName('tbJobNo')
    tb_by = window.FindName('tbBy')
    cb_wsfu = window.FindName('cbWsfu')
    tb_return_note = window.FindName('tbReturnNote')
    sp_hot = window.FindName('spHotSystems')
    btn_size = window.FindName('btnSize')
    btn_cancel = window.FindName('btnCancel')

    # The basis block comes from the data module, so the dialog can never show
    # a basis different from the one the sizing actually used.
    tb_basis.Text = "\n".join(water_tables.basis_of_design_lines())

    tb_job.Text = project_info.get("job", "") or ""
    tb_job_no.Text = project_info.get("job_number", "") or ""
    tb_by.Text = project_info.get("by", "") or ""

    candidates = (return_detection or {}).get("candidates") or []
    certain = (return_detection or {}).get("certain", False)

    if not candidates:
        tb_return_note.Text = (
            "No hot water piping was found on this network, so there is "
            "nothing to mark as a return.")
    elif certain:
        tb_return_note.Text = (
            "Detected from the model. A ticked system is REPORTED BUT NOT "
            "SIZED, because return piping is sized on circulation flow "
            "rather than on fixture units. Change a tick if this is wrong.")
    else:
        tb_return_note.Text = (
            "No recirculation pump or return-to-heater connection was found, "
            "so the tick below is a best guess from the System Type names "
            "and pipe counts. CHECK IT. A ticked system is reported but NOT "
            "sized, because return piping is sized on circulation flow "
            "rather than on fixture units.")

    checkboxes = []
    for entry in candidates:
        box = CheckBox()
        box.Content = "{}  -  {} pipe(s)".format(
            entry["name"], entry["pipe_count"])
        box.IsChecked = bool(entry.get("detected"))
        box.Margin = Thickness(0, 0, 0, 2)
        sp_hot.Children.Add(box)
        checkboxes.append((box, entry["id"]))

        # Say WHY, every time, so a wrong tick is obvious rather than magic.
        reason = TextBlock()
        if entry.get("reasons"):
            reason.Text = "Return, because " + "; ".join(entry["reasons"]) + "."
        else:
            reason.Text = "Supply. Nothing marks this as a return."
        reason.FontSize = 10
        reason.Foreground = Brushes.Gray
        reason.TextWrapping = TextWrapping.Wrap
        reason.Margin = Thickness(20, 0, 0, 8)
        sp_hot.Children.Add(reason)

    if not candidates:
        empty = TextBlock()
        empty.Text = "No hot water System Type found on this network."
        empty.FontSize = 11
        empty.Foreground = Brushes.Gray
        sp_hot.Children.Add(empty)

    result = [None]

    def on_size(sender, e):
        returns = set()
        for box, type_id in checkboxes:
            if box.IsChecked:
                returns.add(type_id)
        result[0] = {
            "job": tb_job.Text.strip(),
            "job_number": tb_job_no.Text.strip(),
            "by": tb_by.Text.strip(),
            "wsfu_calcs": bool(cb_wsfu.IsChecked),
            "return_system_type_ids": returns,
        }
        window.Close()

    def on_cancel(sender, e):
        window.Close()

    btn_size.Click += on_size
    btn_cancel.Click += on_cancel
    window.ShowDialog()

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
