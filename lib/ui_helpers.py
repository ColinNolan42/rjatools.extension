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


_PICKER_XAML = (
    '<Window'
    ' xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
    ' Height="290" Width="460"'
    ' ResizeMode="NoResize"'
    ' WindowStartupLocation="CenterScreen">'
    '<StackPanel Margin="15">'
    '<TextBlock Text="Pipe Material" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<ComboBox Name="cbMaterial" Margin="0,0,0,12"/>'
    '<TextBlock Text="IFGC Table" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<ComboBox Name="cbTable" Margin="0,0,0,12"/>'
    '<TextBlock Text="Project Elevation (ft)" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<TextBox Name="tbElevation" Margin="0,0,0,4"/>'
    '<TextBlock Text="RJA standard: 3% MBH/CFH derate per 1,000 ft above sea level."'
    ' FontStyle="Italic" FontSize="10" Foreground="Gray" Margin="0,0,0,16"/>'
    '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">'
    '<Button Name="btnCancel" Content="Cancel" Width="70" Margin="0,0,8,0"/>'
    '<Button Name="btnOK" Content="OK" Width="70"/>'
    '</StackPanel>'
    '</StackPanel>'
    '</Window>'
)


def show_table_picker(title):
    """Single WPF window with linked material/table dropdowns plus a project
    elevation field (used for the RJA altitude derate - see
    sizing_engine.altitude_derate_factor()).

    Selecting a material instantly repopulates the table list.

    Args:
        title: Window title string.

    Returns:
        (pipe_material, short_table_label, elevation_ft) or (None, None, None)
        if cancelled.
    """
    window       = XamlReader.Parse(_PICKER_XAML)
    window.Title = title
    cb_material  = window.FindName('cbMaterial')
    cb_table     = window.FindName('cbTable')
    tb_elevation = window.FindName('tbElevation')
    btn_ok       = window.FindName('btnOK')
    btn_cancel   = window.FindName('btnCancel')

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
    tb_elevation.Text = str(int(shared_params.DEFAULT_PROJECT_ELEVATION_FT))

    def on_material_changed(sender, e):
        if cb_material.SelectedItem is not None:
            populate_table(cb_material.SelectedItem)

    cb_material.SelectionChanged += on_material_changed

    result = [None, None, None]

    def on_ok(sender, e):
        try:
            elevation = float(tb_elevation.Text.strip())
        except ValueError:
            tb_elevation.Background = Brushes.LightPink
            return
        tb_elevation.Background = Brushes.White
        result[0] = cb_material.SelectedItem
        result[1] = cb_table.SelectedItem
        result[2] = elevation
        window.Close()

    def on_cancel(sender, e):
        window.Close()

    btn_ok.Click     += on_ok
    btn_cancel.Click += on_cancel
    window.ShowDialog()

    return result[0], result[1], result[2]
