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
    ' Width="460" SizeToContent="Height"'
    ' ResizeMode="NoResize"'
    ' WindowStartupLocation="CenterScreen">'
    '<StackPanel Margin="15">'
    '<TextBlock Text="Pipe Material" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<ComboBox Name="cbMaterial" Margin="0,0,0,12"/>'
    '<TextBlock Text="IFGC Table" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<ComboBox Name="cbTable" Margin="0,0,0,12"/>'
    '<CheckBox Name="cbPrv" Content="Mid Stream PRV" FontWeight="SemiBold"'
    ' Visibility="Collapsed" Margin="0,0,0,8"/>'
    '<StackPanel Name="pnlPrv" Visibility="Collapsed" Margin="18,0,0,8">'
    '<TextBlock Text="IFGC Table Downstream of the PRV" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<ComboBox Name="cbDownstream" Margin="0,0,0,4"/>'
    '<TextBlock Text="Same material and gas as above. Pipe after each mid-stream regulator is sized on'
    ' this table and on its own longest run; pipe before it uses the table above. A regulator with'
    ' 10 ft of pipe or less after it is treated as an equipment regulator and ignored."'
    ' FontStyle="Italic" FontSize="10" Foreground="Gray" TextWrapping="Wrap" Margin="0,0,0,4"/>'
    '<TextBlock Name="tbPrvError" Foreground="Firebrick" FontSize="11"'
    ' TextWrapping="Wrap" Visibility="Collapsed" Margin="0,0,0,4"/>'
    '</StackPanel>'
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


_EXISTING_GAS_XAML = (
    '<Window'
    ' xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
    ' Width="460" SizeToContent="Height"'
    ' ResizeMode="NoResize"'
    ' WindowStartupLocation="CenterScreen">'
    '<StackPanel Margin="15">'
    '<TextBlock Text="A pipe was picked, so the run starts at the connection to existing gas'
    ' piping. The existing piping is assumed not modeled. Enter it below."'
    ' TextWrapping="Wrap" Margin="0,0,0,12"/>'
    '<TextBlock Text="Existing Pipe Size at the Connection" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<ComboBox Name="cbSize" Margin="0,0,0,12"/>'
    '<TextBlock Text="Length, Meter to New Connection (ft)" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<TextBox Name="tbLength" Margin="0,0,0,12"/>'
    '<TextBlock Text="Total Existing Load (MBH)" FontWeight="SemiBold" Margin="0,0,0,4"/>'
    '<TextBox Name="tbLoad" Margin="0,0,0,4"/>'
    '<TextBlock Text="The existing pipe is checked for the existing plus new load and is never'
    ' resized. The length is added to the new longest run, which is assumed to be the longest'
    ' run in the system. The engineer shall verify the existing appliances."'
    ' FontStyle="Italic" FontSize="10" Foreground="Gray" TextWrapping="Wrap" Margin="0,0,0,16"/>'
    '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">'
    '<Button Name="btnCancel" Content="Cancel" Width="70" Margin="0,0,8,0"/>'
    '<Button Name="btnOK" Content="OK" Width="70"/>'
    '</StackPanel>'
    '</StackPanel>'
    '</Window>'
)


def show_existing_gas_dialog(title, pipe_sizes, defaults=None):
    """Existing Conditions box, shown when a pipe was picked instead of a meter.

    Opens the RJA themed dialog (sizer_dialogs), and falls back to the plain
    one below if that cannot be built. See _show_existing_gas_dialog_legacy
    for the arguments and the return value.
    """
    import sizer_dialogs
    return sizer_dialogs.themed_or(
        lambda use_theme: sizer_dialogs.show_existing_gas_dialog(
            title, pipe_sizes, defaults, use_theme),
        lambda: _show_existing_gas_dialog_legacy(title, pipe_sizes, defaults))


def _show_existing_gas_dialog_legacy(title, pipe_sizes, defaults=None):
    """FALLBACK ONLY. Plain Existing Conditions box.

    Colin, 2026-10-13: the existing gas piping is assumed not modeled, so the
    designer enters its size, its length from the meter to the new connection
    and its total load. Shown AFTER the pick, and only for a pipe.

    Args:
        title: window title string.
        pipe_sizes: nominal size strings of the selected IFGC table, smallest
            first (gas_tables.list_pipe_sizes).
        defaults: optional dict with "size_nominal", "length_ft", "load_mbh"
            to prefill, for example what was entered in Size Gas.

    Returns:
        dict with "size_nominal" (str), "length_ft" and "load_mbh" (floats),
        or None if cancelled.
    """
    window     = XamlReader.Parse(_EXISTING_GAS_XAML)
    window.Title = title
    cb_size    = window.FindName('cbSize')
    tb_length  = window.FindName('tbLength')
    tb_load    = window.FindName('tbLoad')
    btn_ok     = window.FindName('btnOK')
    btn_cancel = window.FindName('btnCancel')

    defaults = defaults or {}
    pick = 0
    for i, size in enumerate(pipe_sizes):
        cb_size.Items.Add(size)
        if size == defaults.get("size_nominal"):
            pick = i
    if cb_size.Items.Count > 0:
        cb_size.SelectedIndex = pick
    if defaults.get("length_ft") is not None:
        tb_length.Text = "{:g}".format(float(defaults["length_ft"]))
    if defaults.get("load_mbh") is not None:
        tb_load.Text = "{:g}".format(float(defaults["load_mbh"]))

    result = [None]

    def _number(box):
        """A number of zero or more, or None. Marks the box when it is bad."""
        try:
            value = float(box.Text.strip())
            if value < 0 or value != value or value == float("inf"):
                raise ValueError
        except ValueError:
            box.Background = Brushes.LightPink
            return None
        box.Background = Brushes.White
        return value

    def on_ok(sender, e):
        length = _number(tb_length)
        load = _number(tb_load)
        if length is None or load is None or cb_size.SelectedItem is None:
            return
        result[0] = {
            "size_nominal": cb_size.SelectedItem,
            "length_ft":    length,
            "load_mbh":     load,
        }
        window.Close()

    def on_cancel(sender, e):
        window.Close()

    btn_ok.Click     += on_ok
    btn_cancel.Click += on_cancel
    window.ShowDialog()
    return result[0]


def show_water_dialog(title, project_info, return_detection):
    """The one Size Water dialog.

    Opens the RJA themed dialog (sizer_dialogs), and falls back to the plain
    one below if that cannot be built. See _show_water_dialog_legacy for the
    arguments and the return value.
    """
    import sizer_dialogs
    return sizer_dialogs.themed_or(
        lambda use_theme: sizer_dialogs.show_water_dialog(
            title, project_info, return_detection, use_theme),
        lambda: _show_water_dialog_legacy(
            title, project_info, return_detection))


def _show_water_dialog_legacy(title, project_info, return_detection):
    """FALLBACK ONLY. The code-built Size Water dialog.

    Same idiom as show_velocity_settings_dialog(): options first, bold grey
    section headers, separators between sections, and the basis of design as
    label / value rows at the BOTTOM, directly above OK / Cancel.

    Two check boxes say what the run does. Sizing is always on and shown
    disabled, so the dialog states the job rather than offering it. WSFU
    Calculations is optional and produces the take-off.

    The hot water RETURN is NOT asked about. It is detected from the pipes'
    System Types before this dialog opens, because System Type separates the
    recirculation system from the hot supply even though system
    CLASSIFICATION does not. The finding is printed to the pyRevit window.
    What IS asked is the pipe heat loss and temperature drop the circuit
    setter flows are worked out from (Colin, 2026-10-12).

    Args:
        title: window title string.
        project_info: dict with "job", "job_number", "by" defaults, read from
            Revit Project Information by the caller.
        return_detection: the dict from
            water_graph.detect_return_system_types(). Reported, not asked.

    Returns:
        dict with "job", "job_number", "by", "wsfu_calcs" (bool),
        "return_system_type_ids" (set of ints),
        "recirc_heat_loss_btu_per_100ft" and "recirc_delta_t_f" (floats),
        or None if cancelled.
    """
    from System.Windows import (
        Thickness, TextWrapping, SizeToContent, WindowStartupLocation,
        HorizontalAlignment, FontWeights, Window, Visibility)
    from System.Windows.Controls import (
        StackPanel, TextBlock, TextBox, CheckBox, Button, Label, Separator,
        Orientation, Grid, ColumnDefinition, RowDefinition)
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

    # -- 3. hot water recirculation inputs ----------------------------------
    # The two numbers the firm circ size calc asks for. Greyed out when the
    # traversal found no return piping, so the dialog does not ask about a
    # system that is not there.
    detected = set((return_detection or {}).get("detected") or set())
    use_recirc = return_detection is None or bool(detected)

    _section("Hot Water Recirculation")
    tb_heat_loss = _field(
        "Pipe heat loss (BTU/h per 100 ft)",
        "{:.0f}".format(float(
            shared_params.DEFAULT_RECIRC_HEAT_LOSS_BTU_PER_100FT)))
    tb_delta_t = _field(
        "Temperature drop (F), {:.0f} maximum".format(
            float(shared_params.MAX_RECIRC_DELTA_T_F)),
        "{:.0f}".format(float(shared_params.DEFAULT_RECIRC_DELTA_T_F)))
    if use_recirc:
        _note("Sets the flow at each circuit setter: effective length / 100 "
              "x heat loss / (500 x temperature drop), {:.1f} gpm minimum. "
              "The return is sized on the flows added up toward the water "
              "heater.".format(float(shared_params.MIN_CIRCUIT_SETTER_GPM)) +
              (" Used only if the piping picked next has hot water return "
               "piping." if return_detection is None else ""),
              indent=2)
    else:
        tb_heat_loss.IsEnabled = False
        tb_delta_t.IsEnabled = False
        _note("No hot water return piping was found on this network, so "
              "these are not used.", indent=2)

    tb_recirc_error = TextBlock()
    tb_recirc_error.TextWrapping = TextWrapping.Wrap
    tb_recirc_error.Width = CONTENT_W
    tb_recirc_error.Foreground = SolidColorBrush(Colors.Firebrick)
    tb_recirc_error.Margin = Thickness(2, 0, 0, 0)
    tb_recirc_error.Visibility = Visibility.Collapsed
    outer.Children.Add(tb_recirc_error)

    _separator()

    # -- 4. basis of design, at the bottom, matching Calculation Basis ------
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

    # The firm WSFU-to-pipe-size chart, the rows the sizing looks up.
    chart = water_tables.sizing_chart()
    chart_grid = Grid()
    chart_grid.Margin = Thickness(0, 6, 0, 0)
    for _ in chart["headers"]:
        chart_grid.ColumnDefinitions.Add(ColumnDefinition())
    for r, cells in enumerate([chart["headers"]] + chart["rows"]):
        chart_grid.RowDefinitions.Add(RowDefinition())
        for c, value in enumerate(cells):
            cell = TextBlock()
            cell.Text = value
            cell.Foreground = SolidColorBrush(Colors.DimGray)
            if r == 0:
                cell.FontWeight = FontWeights.Bold
            Grid.SetRow(cell, r)
            Grid.SetColumn(cell, c)
            chart_grid.Children.Add(cell)
    outer.Children.Add(chart_grid)

    # The detected return system is NOT shown here. The pyRevit window prints
    # which System Type was identified and why. It is carried through to the
    # result so the engine sizes those pipes on circulation flow.

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

    def _recirc_error(text):
        tb_recirc_error.Text = text
        tb_recirc_error.Visibility = Visibility.Visible

    def _positive_number(box):
        try:
            value = float(box.Text.strip())
        except Exception:
            return None
        return value if value > 0 else None

    def on_ok(sender, e):
        heat_loss = shared_params.DEFAULT_RECIRC_HEAT_LOSS_BTU_PER_100FT
        delta_t = shared_params.DEFAULT_RECIRC_DELTA_T_F
        if use_recirc:
            heat_loss = _positive_number(tb_heat_loss)
            delta_t = _positive_number(tb_delta_t)
            if heat_loss is None:
                _recirc_error(
                    "Pipe heat loss shall be a number greater than zero.")
                return
            if delta_t is None or delta_t > shared_params.MAX_RECIRC_DELTA_T_F:
                _recirc_error(
                    "Temperature drop shall be a number greater than zero "
                    "and not more than {:.0f} F.".format(
                        float(shared_params.MAX_RECIRC_DELTA_T_F)))
                return
        result[0] = {
            "job": tb_job.Text.strip(),
            "job_number": tb_job_no.Text.strip(),
            "by": tb_by.Text.strip(),
            "wsfu_calcs": bool(cb_wsfu.IsChecked),
            # Detected, not chosen. The dialog never asked.
            "return_system_type_ids": set(detected),
            "recirc_heat_loss_btu_per_100ft": float(heat_loss),
            "recirc_delta_t_f": float(delta_t),
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

    Selecting a material instantly repopulates the table list. This is the
    no-regulator dialog (One-Line uses it); Size Gas uses
    show_size_gas_dialog(), which adds the Mid Stream PRV option.

    Args:
        title: Window title string.

    Returns:
        (pipe_material, short_table_label, heat_content_btu_per_cf) or
        (None, None, None) if cancelled.
    """
    res = _show_picker(title, False)
    if res is None:
        return None, None, None
    return res["pipe_material"], res["table_label"], res["heat_content"]


def show_size_gas_dialog(title):
    """show_table_picker() plus a "Mid Stream PRV" check box.

    Checking the box reveals a second dropdown for the IFGC table that governs
    the piping after each pressure regulating valve (same pipe material and
    gas as the first table; it may not be a higher pressure). The user picks
    the table rather than it being inferred: several tables share one
    pressure (four different drops under 2 psi), and the regulator family
    carries no pressure of its own. Only a MID-STREAM regulator (more than
    shared_params.PRV_MIDSTREAM_MIN_DOWNSTREAM_FT of pipe after it) switches
    tables; the traversal decides which regulators those are.

    Returns:
        dict with "pipe_material", "table_label", "heat_content",
        "mid_stream_prv" (bool) and "downstream_table_label" (short label, or
        None when the box is off), or None if cancelled.
    """
    return _show_picker(title, True)


def _show_picker(title, prv_option):
    """Shared implementation behind the two public dialogs.

    Opens the RJA themed dialog (sizer_dialogs), and falls back to the plain
    one below if that cannot be built.
    """
    import sizer_dialogs
    return sizer_dialogs.themed_or(
        lambda use_theme: sizer_dialogs.show_gas_picker(
            title, prv_option, use_theme),
        lambda: _show_picker_legacy(title, prv_option))


def _show_picker_legacy(title, prv_option):
    """FALLBACK ONLY. The plain material / table picker."""
    from System.Windows import Visibility

    window          = XamlReader.Parse(_PICKER_XAML)
    window.Title    = title
    cb_material     = window.FindName('cbMaterial')
    cb_table        = window.FindName('cbTable')
    cb_prv          = window.FindName('cbPrv')
    pnl_prv         = window.FindName('pnlPrv')
    cb_downstream   = window.FindName('cbDownstream')
    tb_prv_error    = window.FindName('tbPrvError')
    tb_heat_content = window.FindName('tbHeatContent')
    btn_ok          = window.FindName('btnOK')
    btn_cancel      = window.FindName('btnCancel')

    materials = gas_tables.get_material_labels()
    for m in materials:
        cb_material.Items.Add(m)
    cb_material.SelectedIndex = 0

    def _upstream_option():
        if cb_material.SelectedItem is None or cb_table.SelectedItem is None:
            return None
        try:
            return gas_tables.get_table_option_by_material_and_short_label(
                cb_material.SelectedItem, cb_table.SelectedItem)
        except ValueError:
            return None

    def populate_downstream():
        """Offer same-material, same-gas tables; pre-select the first one
        that is a real step down from the upstream table."""
        cb_downstream.Items.Clear()
        up = _upstream_option()
        if up is None:
            return
        labels = gas_tables.get_table_option_labels_for_material_and_gas(
            up["material"], up["gas"])
        pick = 0
        found_lower = False
        for i, lbl in enumerate(labels):
            cb_downstream.Items.Add(lbl)
            if not found_lower:
                opt = gas_tables.get_table_option_by_material_and_short_label(
                    up["material"], lbl)
                if opt["inlet_pressure_psi"] < up["inlet_pressure_psi"]:
                    pick = i
                    found_lower = True
        if cb_downstream.Items.Count > 0:
            cb_downstream.SelectedIndex = pick

    def populate_table(mat):
        cb_table.Items.Clear()
        for lbl in gas_tables.get_table_option_labels_for_material(mat):
            cb_table.Items.Add(lbl)
        if cb_table.Items.Count > 0:
            cb_table.SelectedIndex = 0

    populate_table(materials[0])
    tb_heat_content.Text = str(int(shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF))

    if prv_option:
        cb_prv.Visibility = Visibility.Visible
        populate_downstream()

    def on_material_changed(sender, e):
        if cb_material.SelectedItem is not None:
            populate_table(cb_material.SelectedItem)

    def on_table_changed(sender, e):
        if prv_option:
            populate_downstream()

    def on_prv_toggled(sender, e):
        on = bool(cb_prv.IsChecked)
        pnl_prv.Visibility = Visibility.Visible if on else Visibility.Collapsed
        tb_prv_error.Visibility = Visibility.Collapsed

    cb_material.SelectionChanged += on_material_changed
    cb_table.SelectionChanged    += on_table_changed
    cb_prv.Checked               += on_prv_toggled
    cb_prv.Unchecked             += on_prv_toggled

    result = [None]

    def on_ok(sender, e):
        try:
            heat_content = float(tb_heat_content.Text.strip())
            if heat_content <= 0:
                raise ValueError
        except ValueError:
            tb_heat_content.Background = Brushes.LightPink
            return
        tb_heat_content.Background = Brushes.White

        use_prv = bool(cb_prv.IsChecked) if prv_option else False
        downstream_label = None
        if use_prv:
            up = _upstream_option()
            downstream_label = cb_downstream.SelectedItem
            if up is None or downstream_label is None:
                tb_prv_error.Text = "Pick the table downstream of the PRV."
                tb_prv_error.Visibility = Visibility.Visible
                return
            down = gas_tables.get_table_option_by_material_and_short_label(
                up["material"], downstream_label)
            if down["inlet_pressure_psi"] > up["inlet_pressure_psi"]:
                tb_prv_error.Text = (
                    "The downstream table is a higher pressure than the "
                    "table above. A regulator steps pressure down.")
                tb_prv_error.Visibility = Visibility.Visible
                return

        result[0] = {
            "pipe_material":          cb_material.SelectedItem,
            "table_label":            cb_table.SelectedItem,
            "heat_content":           heat_content,
            "mid_stream_prv":         use_prv,
            "downstream_table_label": downstream_label,
        }
        window.Close()

    def on_cancel(sender, e):
        window.Close()

    btn_ok.Click     += on_ok
    btn_cancel.Click += on_cancel
    window.ShowDialog()

    res = result[0]
    if res is None or not res["pipe_material"] or not res["table_label"]:
        return None
    return res


# =============================================================================
# RJA theme helpers (added for the Duct Velocity dialogs)
# =============================================================================
# Everything below is NEW and self-contained. Nothing above this line uses it,
# so the gas and water dialogs are unaffected. Imports are done inside each
# function so a problem here can never stop this module from importing.
#
# Every cosmetic step (theme file, logo, icon) degrades silently: a dialog built
# from these helpers is always usable, just plainer.

RJA_COLORS = {
    'blue':       (0x00, 0x70, 0xC0),   # sampled from the RJA logo
    'blue_dark':  (0x00, 0x5A, 0x9C),   # derived
    'tint':       (0xE6, 0xF1, 0xFA),   # derived
    'text':       (0x1F, 0x1F, 0x1F),
    'muted':      (0x59, 0x59, 0x59),
    'rule':       (0xD0, 0xD0, 0xD0),
    'panel':      (0xF5, 0xF7, 0xFA),
    'error':      (0xC6, 0x28, 0x28),
    'error_fill': (0xFD, 0xEC, 0xEA),
}


def rja_brush(name):
    """A new SolidColorBrush for one of the RJA_COLORS names."""
    from System.Windows.Media import SolidColorBrush, Color
    r, g, b = RJA_COLORS[name]
    return SolidColorBrush(Color.FromRgb(r, g, b))


def _rja_lib_dir():
    import os
    return os.path.dirname(os.path.abspath(__file__))


def rja_theme_path():
    """Absolute path of lib/rja_theme.xaml, or None if the file is missing."""
    import os
    p = os.path.join(_rja_lib_dir(), 'rja_theme.xaml')
    return p if os.path.isfile(p) else None


def rja_logo_path():
    """Absolute path of lib/assets/rja_logo.png, or None if the file is missing."""
    import os
    p = os.path.join(_rja_lib_dir(), 'assets', 'rja_logo.png')
    return p if os.path.isfile(p) else None


def rja_window_resources_xaml(use_theme=True):
    """The <Window.Resources> XAML fragment that merges rja_theme.xaml.

    Returns an empty string when the theme is not wanted or not found, so the
    window loads with plain default styling. The theme is referenced by an
    absolute file URI built by System.Uri, which escapes the space in the
    user profile path.
    """
    if not use_theme:
        return ''
    try:
        from System import Uri
        from xml.sax.saxutils import escape
        p = rja_theme_path()
        if p is None:
            return ''
        uri = Uri(p).AbsoluteUri
        return ('<Window.Resources><ResourceDictionary>'
                '<ResourceDictionary.MergedDictionaries>'
                '<ResourceDictionary Source="' + escape(uri, {'"': '&quot;'}) + '"/>'
                '</ResourceDictionary.MergedDictionaries>'
                '</ResourceDictionary></Window.Resources>')
    except Exception:
        return ''


def rja_header_xaml(title, purpose=None):
    """XAML for the white header band: logo (x:Name img_logo), title, purpose.

    Sits in Grid.Row 0. A 3 px RJA blue rule runs under it. The rule color is
    written literally so it shows even when the theme file did not load.
    """
    from xml.sax.saxutils import escape
    q = {'"': '&quot;'}
    purpose_xaml = ''
    if purpose:
        purpose_xaml = ('<TextBlock Style="{DynamicResource RjaHint}" TextWrapping="Wrap"'
                        ' FontSize="11" Margin="0,2,0,0" Text="' + escape(purpose, q) + '"/>')
    return (
        '<Border Grid.Row="0" Background="White" BorderBrush="#0070C0"'
        ' BorderThickness="0,0,0,3" Padding="16,10,16,10">'
        '<Grid>'
        '<Grid.ColumnDefinitions>'
        '<ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>'
        '</Grid.ColumnDefinitions>'
        '<Image x:Name="img_logo" Height="44" Stretch="Uniform"'
        ' RenderOptions.BitmapScalingMode="HighQuality"'
        ' VerticalAlignment="Center" Margin="0,0,14,0"/>'
        '<StackPanel Grid.Column="1" VerticalAlignment="Center">'
        '<TextBlock FontSize="16" FontWeight="Bold" Text="' + escape(title, q) + '"/>'
        + purpose_xaml +
        '</StackPanel>'
        '</Grid>'
        '</Border>')


def rja_logo_bitmap():
    """The logo as a frozen BitmapImage, or None if missing or unreadable.

    OnLoad caching releases the file handle straight away, so the file is never
    left locked.
    """
    try:
        p = rja_logo_path()
        if p is None:
            return None
        from System import Uri, UriKind
        from System.Windows.Media.Imaging import (
            BitmapImage, BitmapCacheOption, BitmapCreateOptions)
        bmp = BitmapImage()
        bmp.BeginInit()
        bmp.UriSource = Uri(p, UriKind.Absolute)
        bmp.CacheOption = BitmapCacheOption.OnLoad
        bmp.CreateOptions = BitmapCreateOptions.IgnoreImageCache
        bmp.EndInit()
        bmp.Freeze()
        return bmp
    except Exception:
        return None


def rja_apply_logo(window, image_element):
    """Show the logo in image_element and use it as the window icon.

    On any failure the Image is collapsed (title shifts left) and nothing is
    raised.
    """
    bmp = rja_logo_bitmap()
    try:
        if bmp is None:
            from System.Windows import Visibility
            image_element.Visibility = Visibility.Collapsed
            return False
        image_element.Source = bmp
    except Exception:
        return False
    try:
        window.Icon = bmp
    except Exception:
        pass
    return True


def rja_hint_block(element, text):
    """A muted, wrapping hint TextBlock. Uses the theme's RjaHint style when it
    loaded on `element`, otherwise sets the same look directly."""
    from System.Windows.Controls import TextBlock
    from System.Windows import TextWrapping
    tb = TextBlock()
    tb.Text = text
    tb.TextWrapping = TextWrapping.Wrap
    style = None
    try:
        style = element.TryFindResource('RjaHint')
    except Exception:
        style = None
    if style is not None:
        tb.Style = style
    else:
        tb.FontSize = 11
        tb.Foreground = rja_brush('muted')
    return tb


def rja_mark_error(box, expanders=None):
    """Red border and light red fill on a TextBox, and open every Expander in
    `expanders` so the user can see it."""
    try:
        box.BorderBrush = rja_brush('error')
        box.Background = rja_brush('error_fill')
    except Exception:
        pass
    for exp in (expanders or ()):
        try:
            exp.IsExpanded = True
        except Exception:
            pass


def rja_clear_error(box):
    """Remove the error highlight and fall back to the theme or default look."""
    try:
        from System.Windows.Controls import Control
        box.ClearValue(Control.BorderBrushProperty)
        box.ClearValue(Control.BackgroundProperty)
    except Exception:
        pass


def rja_attach_clear_on_edit(box):
    """Clear the error highlight the moment the user edits the box."""
    def _on_changed(sender, args):
        rja_clear_error(sender)
    box.TextChanged += _on_changed
