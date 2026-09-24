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

    '<TextBlock Text="Hot Water System Types In This Model"'
    ' FontWeight="Bold" Margin="0,0,0,4"/>'
    '<TextBlock Text="Revit classifies a recirculation system as Domestic Hot'
    ' Water, exactly like the hot supply, so the tool cannot tell them apart on'
    ' its own. Tick any system below that is a RECIRCULATION or RETURN system.'
    ' Ticked systems are reported but NOT sized, because return piping is sized'
    ' on circulation flow rather than fixture units."'
    ' FontSize="10" Foreground="Gray" TextWrapping="Wrap" Margin="0,0,0,6"/>'
    '<StackPanel Name="spHotSystems" Margin="8,0,0,14"/>'

    '<TextBlock Text="What This Run Will Do" FontWeight="Bold"'
    ' Margin="0,0,0,4"/>'
    '<TextBlock FontSize="10" Foreground="Gray" TextWrapping="Wrap"'
    ' Margin="0,0,0,2"'
    ' Text="CREATE REPORT - traverses the system, checks it is complete, works'
    ' out every pipe size, and puts the WSFU take-off on a drafting view and a'
    ' new sheet. Nothing in the model is changed."/>'
    '<TextBlock FontSize="10" Foreground="Gray" TextWrapping="Wrap"'
    ' Margin="0,0,0,6"'
    ' Text="SIZE WATER - does the same, then writes the new size into every'
    ' pipe it sized and resizes the fittings to match. The report prints in'
    ' the pyRevit window."/>'
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


def show_water_dialog(title, project_info, hot_system_types):
    """The one Size Water dialog. One dialog, then everything is automatic.

    There is no separate report action. Sizing always produces the WSFU
    take-off on a drafting view and a sheet, because the take-off is how the
    fixture unit count gets checked, and a size nobody can check against a
    take-off is not worth writing.

    It deliberately asks nothing that the firm standard already settles:
    minimum pipe sizes always apply, and the take-off always goes on a sheet.
    The only question left is the one no standard can answer, which piping
    system types are the hot water RETURN, because Revit classifies a
    recirculation system exactly like the hot supply.

    Args:
        title: window title string.
        project_info: dict with "job", "job_number", "by" defaults, read from
            Revit Project Information by the caller.
        hot_system_types: list of (element_id, name, pipe_count) for every
            PipingSystemType classified as Domestic Hot Water.

    Returns:
        dict with "job", "job_number", "by" and "return_system_type_ids"
        (set of ints), or None if cancelled.
    """
    from System.Windows.Controls import CheckBox
    from System.Windows import Thickness

    window = XamlReader.Parse(_WATER_XAML)
    window.Title = title

    tb_basis = window.FindName('tbBasis')
    tb_job = window.FindName('tbJob')
    tb_job_no = window.FindName('tbJobNo')
    tb_by = window.FindName('tbBy')
    sp_hot = window.FindName('spHotSystems')
    btn_size = window.FindName('btnSize')
    btn_cancel = window.FindName('btnCancel')

    # The basis block comes from the data module, so the dialog can never show
    # a basis different from the one the sizing actually used.
    tb_basis.Text = "\n".join(water_tables.basis_of_design_lines())

    tb_job.Text = project_info.get("job", "") or ""
    tb_job_no.Text = project_info.get("job_number", "") or ""
    tb_by.Text = project_info.get("by", "") or ""

    checkboxes = []
    if hot_system_types:
        for entry in hot_system_types:
            type_id, name = entry[0], entry[1]
            pipe_count = entry[2] if len(entry) > 2 else None
            box = CheckBox()
            # The pipe count is the tell: on a real job the return carries far
            # fewer pipes than the supply, which makes the right tick obvious
            # without the tool ever matching on a project's system names.
            if pipe_count is None:
                box.Content = "{}  (id {})".format(name, type_id)
            else:
                box.Content = "{}  -  {} pipe(s)".format(name, pipe_count)
            box.Margin = Thickness(0, 0, 0, 6)
            box.IsChecked = False
            sp_hot.Children.Add(box)
            checkboxes.append((box, type_id))
    else:
        from System.Windows.Controls import TextBlock
        empty = TextBlock()
        empty.Text = "No Domestic Hot Water system types found in this model."
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
