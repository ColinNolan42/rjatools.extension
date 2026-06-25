# -*- coding: ascii -*-
# ui_helpers.py
# Shared WPF dialog helpers for Gas Sizing buttons.
# IronPython 2.7

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows.Markup import XamlReader

import gas_tables


_PICKER_XAML = (
    '<Window'
    ' xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
    ' Height="230" Width="460"'
    ' ResizeMode="NoResize"'
    ' WindowStartupLocation="CenterScreen">'
    '<StackPanel Margin="15">'
    '<TextBlock Text="Pipe Material" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<ComboBox Name="cbMaterial" Margin="0,0,0,12"/>'
    '<TextBlock Text="IFGC Table" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<ComboBox Name="cbTable" Margin="0,0,0,16"/>'
    '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">'
    '<Button Name="btnCancel" Content="Cancel" Width="70" Margin="0,0,8,0"/>'
    '<Button Name="btnOK" Content="OK" Width="70"/>'
    '</StackPanel>'
    '</StackPanel>'
    '</Window>'
)


def show_table_picker(title):
    """Single WPF window with linked material and table dropdowns.

    Selecting a material instantly repopulates the table list.

    Args:
        title: Window title string.

    Returns:
        (pipe_material, short_table_label) or (None, None) if cancelled.
    """
    window      = XamlReader.Parse(_PICKER_XAML)
    window.Title = title
    cb_material = window.FindName('cbMaterial')
    cb_table    = window.FindName('cbTable')
    btn_ok      = window.FindName('btnOK')
    btn_cancel  = window.FindName('btnCancel')

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

    def on_material_changed(sender, e):
        if cb_material.SelectedItem is not None:
            populate_table(cb_material.SelectedItem)

    cb_material.SelectionChanged += on_material_changed

    result = [None, None]

    def on_ok(sender, e):
        result[0] = cb_material.SelectedItem
        result[1] = cb_table.SelectedItem
        window.Close()

    def on_cancel(sender, e):
        window.Close()

    btn_ok.Click     += on_ok
    btn_cancel.Click += on_cancel
    window.ShowDialog()

    return result[0], result[1]
