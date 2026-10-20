# -*- coding: ascii -*-
# sizer_dialogs.py
# RJA themed dialogs for the gas and domestic water sizers.
#
# Same idiom as the Duct Performance settings dialog (Colin, 2026-10-13):
# inline XAML through pyrevit forms.WPFWindow, lib/rja_theme.xaml, the logo
# header band, section headers, inline validation in the footer, and a
# Calculation basis block at the bottom.
#
# Nothing here is called directly by a pushbutton. ui_helpers tries these
# first and falls back to its older plain dialogs, so a tool always opens
# even if a themed window cannot be built. Every function returns exactly
# what its plain counterpart in ui_helpers returns.
#
# IronPython 2.7 / pyRevit

from xml.sax.saxutils import escape as _xml_escape

import gas_tables
import shared_params
import sizing_engine
import ui_helpers
import water_tables


# ---------------------------------------------------------------------------
# Shared window shell
# ---------------------------------------------------------------------------

_SHELL_XAML = '''
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="@@TITLE@@"
        Width="@@WIDTH@@" MinWidth="460" MaxWidth="900"
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

        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto"
                      HorizontalScrollBarVisibility="Disabled">
            <StackPanel Margin="16,4,16,12">
                @@BODY@@
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

_BASIS_XAML = '''
                <Expander x:Name="exp_ref" IsExpanded="True" Margin="0,14,0,0">
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
'''


def _q(text):
    """Escape text for use inside a double-quoted XAML attribute."""
    return _xml_escape(text, {'"': '&quot;'})


def _section(title, top=14):
    return ('<Border Style="{DynamicResource RjaSectionHeader}" Margin="0,%d,0,6">'
            '<TextBlock Style="{DynamicResource RjaSectionTitle}" FontSize="13"'
            ' FontWeight="SemiBold" Text="%s"/></Border>' % (top, _q(title)))


def _hint(text, margin="0,4,0,0", name=None):
    name_attr = ' x:Name="%s"' % name if name else ''
    return ('<TextBlock%s Style="{DynamicResource RjaHint}" FontSize="11"'
            ' TextWrapping="Wrap" Margin="%s" Text="%s"/>'
            % (name_attr, margin, _q(text)))


def _build_window(title, heading, purpose, body_xaml, use_theme, width=560):
    """The themed window, built but not shown."""
    from pyrevit import forms
    from System.Windows import SystemParameters

    xaml = (_SHELL_XAML
            .replace('@@TITLE@@', _q(title))
            .replace('@@WIDTH@@', str(int(width)))
            .replace('@@RESOURCES@@', ui_helpers.rja_window_resources_xaml(use_theme))
            .replace('@@HEADER@@', ui_helpers.rja_header_xaml(heading, purpose))
            .replace('@@BODY@@', body_xaml))
    win = forms.WPFWindow(xaml, literal_string=True)
    ui_helpers.rja_apply_logo(win, _named(win, 'img_logo'))
    try:
        win.MaxHeight = SystemParameters.WorkArea.Height * 0.9
    except Exception:
        win.MaxHeight = 800.0
    return win


def _named(win, name):
    """The element carrying x:Name `name`. Raises if it is missing, which
    sends the caller to the fallback dialog instead of failing later."""
    el = getattr(win, name, None)
    if el is None:
        el = win.FindName(name)
    if el is None:
        raise RuntimeError('XAML element not found: ' + name)
    return el


def _fill_basis(win, rows):
    """Label / value rows in the Calculation basis block."""
    from System.Windows import (FontWeights, GridLength, Thickness,
                                VerticalAlignment)
    from System.Windows.Controls import Grid, RowDefinition, TextBlock

    grid = _named(win, 'grid_ref')
    for r, (label, value) in enumerate(rows):
        rd = RowDefinition()
        rd.Height = GridLength.Auto
        grid.RowDefinitions.Add(rd)

        a = TextBlock()
        a.Text = label.rstrip(':')
        a.FontWeight = FontWeights.SemiBold
        a.Margin = Thickness(0, 2, 12, 2)
        a.VerticalAlignment = VerticalAlignment.Top
        Grid.SetColumn(a, 0)
        Grid.SetRow(a, r)
        grid.Children.Add(a)

        v = ui_helpers.rja_hint_block(win, value)
        v.Margin = Thickness(0, 2, 0, 2)
        Grid.SetColumn(v, 1)
        Grid.SetRow(v, r)
        grid.Children.Add(v)


def _number(box, label, errors, kind='pos', maximum=None):
    """Read a number from a TextBox, or record why it cannot be used."""
    import math
    try:
        value = float(box.Text.strip())
        if math.isnan(value) or math.isinf(value):
            raise ValueError('not finite')
    except (ValueError, TypeError):
        errors.append((box, label + ' shall be a number.'))
        return None
    if kind == 'pos' and value <= 0:
        errors.append((box, label + ' shall be greater than 0.'))
        return None
    if kind == 'nonneg' and value < 0:
        errors.append((box, label + ' cannot be negative.'))
        return None
    if maximum is not None and value > maximum:
        errors.append((box, label + ' shall not be more than %g.' % float(maximum)))
        return None
    return value


def _run(win, boxes, collect):
    """Wire OK / Cancel and show the window.

    collect(errors) returns the result, appending (box, message) pairs to
    errors when something cannot be used. The window stays open on an error
    and says what is wrong in the footer.
    """
    tb_error = _named(win, 'tb_error')
    result = [None]

    for box in boxes:
        ui_helpers.rja_attach_clear_on_edit(box)

    def on_ok(sender, args):
        try:
            for box in boxes:
                ui_helpers.rja_clear_error(box)
            tb_error.Text = ''
            errors = []
            values = collect(errors)
            if errors:
                for box, _message in errors:
                    if box is not None:
                        ui_helpers.rja_mark_error(box)
                message = errors[0][1]
                if len(errors) > 1:
                    message += '  (%d more highlighted)' % (len(errors) - 1)
                tb_error.Text = message
                try:
                    if errors[0][0] is not None:
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

    def on_cancel(sender, args):
        win.Close()

    _named(win, 'btn_ok').Click += on_ok
    _named(win, 'btn_cancel').Click += on_cancel
    win.ShowDialog()
    return result[0]


def themed_or(build, legacy):
    """Run the themed dialog, then the same without the theme file, then the
    plain one, so the tool always opens. None from a themed dialog means the
    user cancelled and is returned as is."""
    import traceback
    for use_theme in (True, False):
        try:
            return build(use_theme)
        except Exception:
            try:
                from pyrevit import script
                script.get_logger().warning(
                    'Themed dialog failed (theme=%s), falling back:\n%s'
                    % (use_theme, traceback.format_exc()))
            except Exception:
                pass
    return legacy()


# ---------------------------------------------------------------------------
# Size Water
# ---------------------------------------------------------------------------

# Shared by the Size Water and Water Performance dialogs.
_PROJECT_XAML = (
    _section('Project information') +
    '<Grid>'
    '<Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>'
    '</Grid.ColumnDefinitions>'
    '<Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>'
    '<RowDefinition Height="Auto"/></Grid.RowDefinitions>'
    '<TextBlock Grid.Row="0" Text="Job name" VerticalAlignment="Center" Margin="0,3,12,3"/>'
    '<TextBox x:Name="tb_job" Grid.Row="0" Grid.Column="1" Margin="0,3,0,3"/>'
    '<TextBlock Grid.Row="1" Text="Job number" VerticalAlignment="Center" Margin="0,3,12,3"/>'
    '<TextBox x:Name="tb_job_no" Grid.Row="1" Grid.Column="1" Margin="0,3,0,3"/>'
    '<TextBlock Grid.Row="2" Text="By" VerticalAlignment="Center" Margin="0,3,12,3"/>'
    '<TextBox x:Name="tb_by" Grid.Row="2" Grid.Column="1" Margin="0,3,0,3"/>'
    '</Grid>'
)

_RECIRC_XAML = (
    _section('Hot water recirculation') +
    '<StackPanel x:Name="pnl_recirc">'
    '<StackPanel Orientation="Horizontal">'
    '<TextBlock Text="Pipe heat loss" Width="110" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_heat_loss" Width="72" Margin="0,0,6,0"'
    ' ToolTip="Heat lost by the hot water piping. Shall be greater than 0."/>'
    '<TextBlock Text="BTU/h per 100 ft" VerticalAlignment="Center"/>'
    '</StackPanel>'
    '<StackPanel Orientation="Horizontal" Margin="0,6,0,0">'
    '<TextBlock Text="Temperature drop" Width="110" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_delta_t" Width="72" Margin="0,0,6,0"'
    ' ToolTip="Temperature drop round the loop. May be lowered, never raised."/>'
    '<TextBlock x:Name="tb_delta_t_unit" VerticalAlignment="Center"/>'
    '</StackPanel>'
    '</StackPanel>' +
    _hint('', '0,4,0,0', name='tb_recirc_note')
)

_WATER_BODY = (
    _section('This run will', 10) +
    '<CheckBox x:Name="cb_size" IsChecked="True" IsEnabled="False" FontWeight="SemiBold"'
    ' Content="Size the piping and write the sizes into the model"/>' +
    _hint('Always on. Cold water is sized on TOTAL fixture units and hot water on HOT '
          'fixture units. Drawn sizes are overwritten and the fittings are resized to '
          'match.', '20,0,0,6') +
    '<CheckBox x:Name="cb_wsfu" IsChecked="True" Content="WSFU Calculations"'
    ' ToolTip="The take-off prints in the pyRevit window either way."/>' +
    _hint('Puts the water supply fixture unit take-off on a drafting view and a new '
          'sheet, so the fixture unit count behind every pipe size can be checked '
          'against the model.', '20,0,0,0') +

    _PROJECT_XAML +
    _RECIRC_XAML +
    _BASIS_XAML
)


def _fill_chart(win, title, chart):
    """A small table under the Calculation basis rows, spanning both columns.

    chart is {"headers": [...], "rows": [[...]]}, e.g. water_tables.sizing_chart().
    """
    from System.Windows import FontWeights, GridLength, Thickness
    from System.Windows.Controls import (ColumnDefinition, Grid, RowDefinition,
                                         TextBlock)

    grid = _named(win, 'grid_ref')

    def _span(element):
        rd = RowDefinition()
        rd.Height = GridLength.Auto
        grid.RowDefinitions.Add(rd)
        Grid.SetRow(element, grid.RowDefinitions.Count - 1)
        Grid.SetColumn(element, 0)
        Grid.SetColumnSpan(element, 2)
        grid.Children.Add(element)

    heading = TextBlock()
    heading.Text = title
    heading.FontWeight = FontWeights.SemiBold
    heading.Margin = Thickness(0, 8, 0, 2)
    _span(heading)

    table = Grid()
    for _ in chart["headers"]:
        cd = ColumnDefinition()
        cd.Width = GridLength.Auto
        table.ColumnDefinitions.Add(cd)
    for r, cells in enumerate([chart["headers"]] + chart["rows"]):
        rd = RowDefinition()
        rd.Height = GridLength.Auto
        table.RowDefinitions.Add(rd)
        for c, value in enumerate(cells):
            cell = ui_helpers.rja_hint_block(win, value)
            if r == 0:
                cell.FontWeight = FontWeights.SemiBold
            cell.Margin = Thickness(0, 1, 18, 1)
            Grid.SetRow(cell, r)
            Grid.SetColumn(cell, c)
            table.Children.Add(cell)
    _span(table)


def show_water_dialog(title, project_info, return_detection, use_theme=True):
    """Themed Size Water dialog. Same arguments and return value as
    ui_helpers.show_water_dialog."""
    win = _build_window(
        title, 'Size Water',
        'Sizes domestic cold water, hot water and hot water return piping per '
        'the 2024 IPC, and writes the sizes into the model.',
        _WATER_BODY, use_theme)

    cb_wsfu = _named(win, 'cb_wsfu')
    project = _init_project(win, project_info)
    recirc = _init_recirc(win, return_detection)

    _fill_basis(win, water_tables.basis_of_design_rows())
    _fill_chart(win, "Firm pipe sizing chart", water_tables.sizing_chart())

    def collect(errors):
        values = recirc(errors)
        if errors:
            return None
        values.update(project())
        values["wsfu_calcs"] = bool(cb_wsfu.IsChecked)
        return values

    return _run(win, [_named(win, 'tb_heat_loss'), _named(win, 'tb_delta_t')],
                collect)


def _init_project(win, project_info):
    """Prefill the Project information fields. Returns a reader."""
    tb_job = _named(win, 'tb_job')
    tb_job_no = _named(win, 'tb_job_no')
    tb_by = _named(win, 'tb_by')
    tb_job.Text = project_info.get("job", "") or ""
    tb_job_no.Text = project_info.get("job_number", "") or ""
    tb_by.Text = project_info.get("by", "") or ""

    def read():
        return {"job": tb_job.Text.strip(),
                "job_number": tb_job_no.Text.strip(),
                "by": tb_by.Text.strip()}
    return read


def _init_recirc(win, return_detection):
    """Prefill the Hot water recirculation section. Returns a reader that
    takes the errors list and gives the return ids and recirc inputs."""
    pnl_recirc = _named(win, 'pnl_recirc')
    tb_heat_loss = _named(win, 'tb_heat_loss')
    tb_delta_t = _named(win, 'tb_delta_t')
    tb_recirc_note = _named(win, 'tb_recirc_note')

    max_dt = float(shared_params.MAX_RECIRC_DELTA_T_F)
    tb_heat_loss.Text = "{:.0f}".format(
        float(shared_params.DEFAULT_RECIRC_HEAT_LOSS_BTU_PER_100FT))
    tb_delta_t.Text = "{:.0f}".format(
        float(shared_params.DEFAULT_RECIRC_DELTA_T_F))
    _named(win, 'tb_delta_t_unit').Text = "F, {:.0f} maximum".format(max_dt)

    # Detected, not chosen. With return_detection None the dialog opens
    # before the network is picked, so the inputs stay live and are used only
    # if the picked network turns out to have hot water return piping.
    detected = set((return_detection or {}).get("detected") or set())
    use_recirc = return_detection is None or bool(detected)
    if use_recirc:
        tb_recirc_note.Text = (
            "Sets the flow at each circuit setter: effective length / 100 x "
            "heat loss / (500 x temperature drop), {:.1f} gpm minimum. The "
            "return is sized on the flows added up toward the water "
            "heater.".format(float(shared_params.MIN_CIRCUIT_SETTER_GPM)))
        if return_detection is None:
            tb_recirc_note.Text += (" Used only if the piping picked next has "
                                    "hot water return piping.")
    else:
        pnl_recirc.IsEnabled = False
        pnl_recirc.Opacity = 0.55
        tb_recirc_note.Text = ("No hot water return piping was found on this "
                               "network, so these are not used.")

    def read(errors):
        heat_loss = shared_params.DEFAULT_RECIRC_HEAT_LOSS_BTU_PER_100FT
        delta_t = shared_params.DEFAULT_RECIRC_DELTA_T_F
        if use_recirc:
            heat_loss = _number(tb_heat_loss, 'Pipe heat loss', errors)
            delta_t = _number(tb_delta_t, 'Temperature drop', errors,
                              maximum=max_dt)
            if errors:
                return {}
        return {
            "return_system_type_ids": set(detected),
            "recirc_heat_loss_btu_per_100ft": float(heat_loss),
            "recirc_delta_t_f": float(delta_t),
        }
    return read


# ---------------------------------------------------------------------------
# Water Performance
# ---------------------------------------------------------------------------

_PERF_BODY = (
    _section('Checks', 10) +
    _hint('Read only: nothing is written to the pipes. Tick at least one. The '
          'results go on a new sheet; with Pipe size and velocity ticked, that '
          'sheet carries a colored copy of the active plan and the color '
          'legend.', '0,0,0,6') +
    '<CheckBox x:Name="cb_velocity" IsChecked="True"'
    ' Content="Pipe size and velocity"/>'
    '<StackPanel Orientation="Horizontal" Margin="20,4,0,0">'
    '<TextBlock Text="Yellow tolerance" Width="110" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_tolerance" Width="72" Margin="0,0,6,0"'
    ' ToolTip="How far computed speed may run over the limit before a pipe at table size turns yellow."/>'
    '<TextBlock Text="% over the speed limit" VerticalAlignment="Center"/>'
    '</StackPanel>' +
    _hint('', '20,4,0,6', name='tb_limits_note') +
    '<CheckBox x:Name="cb_pressure" IsChecked="False" Content="Pressure loss"'
    ' ToolTip="2024 IPC E103.3 segmented loss method, checked against Table 604.3."/>'
    '<StackPanel x:Name="pnl_pressure" Margin="20,4,0,0" IsEnabled="False" Opacity="0.55">'
    '<StackPanel Orientation="Horizontal">'
    '<TextBlock Text="Pressure at start point" Width="150" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_start_psi" Width="72" Margin="0,0,6,0"'
    ' ToolTip="Static pressure at the element you pick next. Shall be greater than 0."/>'
    '<TextBlock Text="psi" VerticalAlignment="Center"/>'
    '</StackPanel>' +
    _hint('The pressure AT THE POINT YOU PICK: what the RPZ delivers (after its '
          'loss), or the pressure in the existing main at the tee. Street, tap, '
          'meter and RPZ losses are upstream of it and are not added again. '
          'Each fixture outlet is checked against its Table 604.3 minimum.',
          '0,4,0,0') +
    '</StackPanel>' +

    '<CheckBox x:Name="cb_pump" IsChecked="False" Margin="0,6,0,0"'
    ' Content="Recirculation pump duty point (preliminary)"'
    ' ToolTip="Flow and head for the pump at each water heater with a return."/>'
    '<StackPanel x:Name="pnl_pump" Margin="20,4,0,0" IsEnabled="False" Opacity="0.55">'
    '<StackPanel Orientation="Horizontal">'
    '<TextBlock Text="Circuit setter drop" Width="150" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_setter_psi" Width="72" Margin="0,0,6,0"'
    ' ToolTip="Design pressure drop across one circuit setter, from its submittal."/>'
    '<TextBlock Text="psi" VerticalAlignment="Center"/>'
    '</StackPanel>'
    '<StackPanel Orientation="Horizontal" Margin="0,6,0,0">'
    '<TextBlock Text="Safety factor on head" Width="150" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_pump_sf" Width="72" Margin="0,0,6,0"/>'
    '<TextBlock Text="%" VerticalAlignment="Center"/>'
    '</StackPanel>' +
    _hint('Preliminary until the RJA pump method is set. '
          'Flow is the circuit setter flows added at the heater. Head is the '
          'friction round the worst circuit at circulation flow, plus the circuit '
          'setter and devices. Used only if the piping picked has a return.',
          '0,4,0,0') +
    '</StackPanel>' +
    '<CheckBox x:Name="cb_wsfu" IsChecked="True" Content="WSFU Calculations"'
    ' Margin="0,6,0,0" ToolTip="The take-off prints in the pyRevit window either way."/>' +
    _hint('Puts the water supply fixture unit take-off on a drafting view and a new '
          'sheet.', '20,0,0,0') +

    _PROJECT_XAML +
    _RECIRC_XAML +
    _BASIS_XAML
)


def show_water_performance_dialog(title, project_info, use_theme=True):
    """Themed Water Performance dialog.

    Returns a dict, or None when cancelled:
        job, job_number, by, wsfu_calcs, speed_tolerance (fraction),
        return_system_type_ids (empty, the run detects it after the pick),
        recirc_heat_loss_btu_per_100ft, recirc_delta_t_f.
    """
    import water_performance

    win = _build_window(
        title, 'Water Performance',
        'Checks domestic cold water, hot water and hot water return piping '
        'against the 2024 IPC firm sizing table. Read only.',
        _PERF_BODY, use_theme)

    cb_wsfu = _named(win, 'cb_wsfu')
    tb_tolerance = _named(win, 'tb_tolerance')
    tb_tolerance.Text = "{:.0f}".format(
        float(water_performance.SPEED_TOLERANCE) * 100.0)
    _named(win, 'tb_limits_note').Text = (
        "Pass or fail is the drawn size against the firm table size. Speed "
        "limits from the firm table: {:.0f} fps cold, {:.0f} fps hot and "
        "return.".format(water_tables.velocity_limit_fps("cold"),
                         water_tables.velocity_limit_fps("hot")))

    import water_pressure
    cb_pressure = _named(win, 'cb_pressure')
    pnl_pressure = _named(win, 'pnl_pressure')
    tb_start_psi = _named(win, 'tb_start_psi')
    tb_start_psi.Text = "{:.1f}".format(water_pressure.DEFAULT_START_PSI)

    def on_pressure_toggled(sender, args):
        on = bool(cb_pressure.IsChecked)
        pnl_pressure.IsEnabled = on
        pnl_pressure.Opacity = 1.0 if on else 0.55
    cb_pressure.Checked += on_pressure_toggled
    cb_pressure.Unchecked += on_pressure_toggled

    import water_pump
    cb_pump = _named(win, 'cb_pump')
    pnl_pump = _named(win, 'pnl_pump')
    tb_setter_psi = _named(win, 'tb_setter_psi')
    tb_pump_sf = _named(win, 'tb_pump_sf')
    tb_setter_psi.Text = "{:.1f}".format(water_pump.DEFAULT_CIRCUIT_SETTER_PSI)
    tb_pump_sf.Text = "{:.0f}".format(water_pump.DEFAULT_SAFETY_FACTOR * 100.0)

    def on_pump_toggled(sender, args):
        on = bool(cb_pump.IsChecked)
        pnl_pump.IsEnabled = on
        pnl_pump.Opacity = 1.0 if on else 0.55
    cb_pump.Checked += on_pump_toggled
    cb_pump.Unchecked += on_pump_toggled

    project = _init_project(win, project_info)
    recirc = _init_recirc(win, None)

    _fill_basis(win, water_tables.basis_of_design_rows())
    _fill_chart(win, "Firm pipe sizing chart", water_tables.sizing_chart())

    def collect(errors):
        cb_velocity = _named(win, 'cb_velocity')
        if not (cb_velocity.IsChecked or cb_pressure.IsChecked or
                cb_pump.IsChecked or cb_wsfu.IsChecked):
            errors.append((None, 'Tick at least one check: pipe size and '
                                 'velocity, pressure loss, the pump duty '
                                 'point or WSFU Calculations.'))
            return None
        tolerance = _number(tb_tolerance, 'Yellow tolerance', errors,
                            kind='nonneg', maximum=100)
        start_psi = None
        if cb_pressure.IsChecked:
            start_psi = _number(tb_start_psi, 'Pressure at start point',
                                errors)
        setter_psi, pump_sf = None, None
        if cb_pump.IsChecked:
            setter_psi = _number(tb_setter_psi, 'Circuit setter drop', errors,
                                 kind='nonneg')
            pump_sf = _number(tb_pump_sf, 'Safety factor', errors,
                              kind='nonneg', maximum=100)
        values = recirc(errors)
        if errors:
            return None
        values.update(project())
        values["wsfu_calcs"] = bool(cb_wsfu.IsChecked)
        values["speed_tolerance"] = float(tolerance) / 100.0
        values["velocity"] = bool(cb_velocity.IsChecked)
        values["pressure_loss"] = bool(cb_pressure.IsChecked)
        values["start_psi"] = start_psi
        values["pump_duty"] = bool(cb_pump.IsChecked)
        values["circuit_setter_psi"] = setter_psi
        values["pump_safety_factor"] = (pump_sf / 100.0
                                        if pump_sf is not None else None)
        return values

    return _run(win, [tb_tolerance, tb_start_psi, tb_setter_psi, tb_pump_sf,
                      _named(win, 'tb_heat_loss'), _named(win, 'tb_delta_t')],
                collect)


# ---------------------------------------------------------------------------
# Start point (Size Water, Water Performance)
# ---------------------------------------------------------------------------

_START_BODY = (
    _section('Where does the domestic water system start?', 10) +
    '<RadioButton x:Name="rb_rpz" GroupName="start" IsChecked="True"'
    ' FontWeight="SemiBold" Content="RPZ"/>' +
    _hint('Pick the RPZ / backflow preventer.', '20,0,0,8') +
    '<RadioButton x:Name="rb_tee" GroupName="start" FontWeight="SemiBold"'
    ' Content="New pipe connection (TEE)"/>' +
    _hint('Existing building: pick the tee on the existing cold water main where '
          'the new piping branches off. The existing main is not checked.',
          '20,0,0,0')
)


def show_start_point_dialog(title, rpz_label, tee_label, use_theme=True):
    """Themed start point choice. Returns rpz_label, tee_label or None."""
    win = _build_window(title, 'Select Start Point',
                        'Choose the start point, then pick it in the model.',
                        _START_BODY, use_theme, width=480)
    rb_rpz = _named(win, 'rb_rpz')

    def collect(errors):
        return rpz_label if rb_rpz.IsChecked else tee_label

    return _run(win, [], collect)


# ---------------------------------------------------------------------------
# Size Duct
# ---------------------------------------------------------------------------

# Systems in the limits grid, in display order. Keys match the duct System
# Type names hvac_graph.FIRM_DEFAULTS uses, so the same limits drive Duct
# Performance and Size Duct. No Outside Air row: Size Duct walks at equipment
# level, which never reaches the outside air ductwork.
DUCT_SYSTEMS = ['Supply Air', 'Return Air', 'Exhaust Air', 'Transfer Air']


def _duct_limits_xaml():
    rows = ''.join('<RowDefinition Height="Auto"/>'
                   for _ in range(len(DUCT_SYSTEMS) + 1))
    cells = ('<TextBlock Grid.Row="0" Grid.Column="1" FontWeight="SemiBold"'
             ' Text="Max velocity (fpm)" Margin="0,0,12,4"/>'
             '<TextBlock Grid.Row="0" Grid.Column="2" FontWeight="SemiBold"'
             ' Text="Max friction (in. wc / 100 ft)" Margin="0,0,0,4"/>')
    for i, name in enumerate(DUCT_SYSTEMS):
        r = i + 1
        cells += ('<TextBlock Grid.Row="%d" Text="%s" VerticalAlignment="Center"'
                  ' Margin="0,2,12,2"/>'
                  '<TextBox x:Name="tb_v_%d" Grid.Row="%d" Grid.Column="1"'
                  ' Width="72" HorizontalAlignment="Left" Margin="0,2,12,2"/>'
                  '<TextBox x:Name="tb_f_%d" Grid.Row="%d" Grid.Column="2"'
                  ' Width="72" HorizontalAlignment="Left" Margin="0,2,0,2"/>'
                  % (r, _q(name), i, r, i, r))
    return ('<Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/>'
            '<ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>'
            '</Grid.ColumnDefinitions><Grid.RowDefinitions>' + rows +
            '</Grid.RowDefinitions>' + cells + '</Grid>')


_DUCT_BODY = (
    _section('This run', 10) +
    _hint('Sizes every main (two or more terminals downstream) and writes the '
          'sizes into the model: the ducts first, then the elbows, dampers and '
          'endcaps between two mains. Takeoff boots, branches and diffusers are '
          'never touched, because the auto-sizing diffusers set the branch. '
          'Existing phase ductwork is left alone. Every joint is audited '
          'afterwards.', '0,0,0,0') +

    _section('Main duct limits') +
    _duct_limits_xaml() +
    _hint('Firm defaults, the same limits Duct Performance checks, so a sized '
          'system comes back green. Each main gets the smallest even size meeting '
          'both. Rectangular mains keep their drawn height and are sized on width '
          '(4:1 maximum); the height steps up only when no width fits. A main '
          'steps down after any tap where the reduced airflow allows, with the '
          'transition just past the tap.',
          '0,6,0,0') +

    _BASIS_XAML
)


def show_size_duct_dialog(title, use_theme=True):
    """Themed Size Duct dialog.

    Returns {"limits": {system: (max fpm, max friction)}} or None.
    """
    import duct_sizing
    import hvac_graph

    win = _build_window(
        title, 'Size Duct',
        'Sizes every main duct to the RJA velocity and friction limits and '
        'writes the sizes into the model.', _DUCT_BODY, use_theme)

    boxes = []
    for i, name in enumerate(DUCT_SYSTEMS):
        fpm, fric = hvac_graph.FIRM_DEFAULTS[name]
        tb_v = _named(win, 'tb_v_%d' % i)
        tb_f = _named(win, 'tb_f_%d' % i)
        tb_v.Text = "{:.0f}".format(float(fpm))
        tb_f.Text = "{:.2f}".format(float(fric))
        boxes.append((name, tb_v, tb_f))

    _fill_basis(win, [
        ("Sizing", "Smallest even size, %d in minimum, meeting BOTH the max "
                   "velocity and the max friction." % duct_sizing.MIN_DUCT_DIM),
        ("Round", "%d to %d in, 2 in steps." % (duct_sizing.ROUND_SIZES[0],
                                               duct_sizing.ROUND_SIZES[-1])),
        ("Rectangular", "Drawn height held, width solved, %d:1 maximum "
                        "aspect." % int(duct_sizing.MAX_ASPECT)),
        ("Friction", "Darcy-Weisbach, Altshul-Tsal friction factor, "
                     "galvanized steel, standard air (same as Duct "
                     "Performance)."),
        ("Not sized", "Branches, takeoff boots, diffusers, flex and oval "
                      "duct, existing phase ductwork, outside air ductwork, "
                      "and any main with a terminal downstream that has no "
                      "Flow value."),
    ])

    def collect(errors):
        limits = {}
        for name, tb_v, tb_f in boxes:
            fpm = _number(tb_v, name + ' velocity', errors)
            fric = _number(tb_f, name + ' friction', errors)
            if fpm is not None and fric is not None:
                limits[name] = (fpm, fric)
        if errors:
            return None
        return {"limits": limits}

    return _run(win, [b for _n, v, f in boxes for b in (v, f)], collect)


# ---------------------------------------------------------------------------
# Size Gas and One-Line: pipe material, IFGC table, heat content
# ---------------------------------------------------------------------------

_GAS_BODY = (
    _section('Pipe and table', 10) +
    '<TextBlock Text="Pipe material" Margin="0,0,0,3"/>'
    '<ComboBox x:Name="cb_material" Margin="0,0,0,8"/>'
    '<TextBlock Text="IFGC table" Margin="0,0,0,3"/>'
    '<ComboBox x:Name="cb_table"/>'
    '<CheckBox x:Name="cb_prv" Content="Mid Stream PRV" Margin="0,10,0,0"'
    ' Visibility="Collapsed"'
    ' ToolTip="Tick when a pressure regulating valve steps the pressure down part way along the piping."/>'
    '<Border x:Name="pnl_prv" Style="{DynamicResource RjaGroupBorder}" Margin="20,4,0,0"'
    ' Visibility="Collapsed">'
    '<StackPanel>'
    '<TextBlock Text="IFGC table downstream of the PRV" Margin="0,0,0,3"/>'
    '<ComboBox x:Name="cb_downstream"/>' +
    _hint('Same material and gas as above. Pipe after each mid-stream regulator is '
          'sized on this table and on its own longest run. Pipe before it uses the '
          'table above. A regulator with %g ft of pipe or less after it is treated as '
          'an equipment regulator and ignored.'
          % float(shared_params.PRV_MIDSTREAM_MIN_DOWNSTREAM_FT)) +
    '</StackPanel>'
    '</Border>' +

    _section('Gas') +
    '<StackPanel Orientation="Horizontal">'
    '<TextBlock Text="Heat content of gas" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_heat" Width="72" Margin="8,0,6,0"'
    ' ToolTip="Get this value from the utility. Shall be greater than 0."/>'
    '<TextBlock Text="BTU/CF" VerticalAlignment="Center"/>'
    '</StackPanel>' +
    _hint('RJA standard: CFH = BTU/h / Heat Content of Gas. Get this value from the '
          'utility. Default %g BTU/CF is Denver. Sea level is 1000 BTU/CF.'
          % float(shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF)) +

    _BASIS_XAML
)


def show_gas_picker(title, prv_option, use_theme=True):
    """Themed pipe material / IFGC table dialog for Size Gas and One-Line.

    Same return value as ui_helpers._show_picker: a dict with
    "pipe_material", "table_label", "heat_content", "mid_stream_prv" and
    "downstream_table_label", or None if cancelled.
    """
    from System.Windows import Visibility

    if prv_option:
        heading = 'Size Gas'
        purpose = ('Sizes gas piping by the IFGC Longest Run Method and writes '
                   'the sizes into the model.')
    else:
        heading = 'Gas One-Line'
        purpose = ('Draws the gas piping one-line diagram on a drafting view, '
                   'with the pipe sizes already in the model.')
    win = _build_window(title, heading, purpose, _GAS_BODY, use_theme)

    cb_material = _named(win, 'cb_material')
    cb_table = _named(win, 'cb_table')
    cb_prv = _named(win, 'cb_prv')
    pnl_prv = _named(win, 'pnl_prv')
    cb_downstream = _named(win, 'cb_downstream')
    tb_heat = _named(win, 'tb_heat')

    materials = gas_tables.get_material_labels()
    for m in materials:
        cb_material.Items.Add(m)
    cb_material.SelectedIndex = 0

    def upstream_option():
        if cb_material.SelectedItem is None or cb_table.SelectedItem is None:
            return None
        try:
            return gas_tables.get_table_option_by_material_and_short_label(
                cb_material.SelectedItem, cb_table.SelectedItem)
        except ValueError:
            return None

    def populate_downstream():
        """Same-material, same-gas tables, with the first real step down
        from the table above pre-selected."""
        cb_downstream.Items.Clear()
        up = upstream_option()
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

    def populate_table(material):
        cb_table.Items.Clear()
        for lbl in gas_tables.get_table_option_labels_for_material(material):
            cb_table.Items.Add(lbl)
        if cb_table.Items.Count > 0:
            cb_table.SelectedIndex = 0

    populate_table(materials[0])
    tb_heat.Text = "{:.0f}".format(
        float(shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF))

    if prv_option:
        cb_prv.Visibility = Visibility.Visible
        populate_downstream()

    def on_material_changed(sender, args):
        if cb_material.SelectedItem is not None:
            populate_table(cb_material.SelectedItem)

    def on_table_changed(sender, args):
        if prv_option:
            populate_downstream()

    def on_prv_toggled(sender, args):
        on = bool(cb_prv.IsChecked)
        pnl_prv.Visibility = Visibility.Visible if on else Visibility.Collapsed

    cb_material.SelectionChanged += on_material_changed
    cb_table.SelectionChanged += on_table_changed
    cb_prv.Checked += on_prv_toggled
    cb_prv.Unchecked += on_prv_toggled

    _fill_basis(win, sizing_engine.basis_of_design_rows())

    def collect(errors):
        heat_content = _number(tb_heat, 'Heat content of gas', errors)
        if not cb_material.SelectedItem or not cb_table.SelectedItem:
            errors.append((None, 'Pick a pipe material and an IFGC table.'))

        use_prv = bool(cb_prv.IsChecked) if prv_option else False
        downstream_label = None
        if use_prv:
            up = upstream_option()
            downstream_label = cb_downstream.SelectedItem
            if up is None or downstream_label is None:
                errors.append((None, 'Pick the table downstream of the PRV.'))
            else:
                down = gas_tables.get_table_option_by_material_and_short_label(
                    up["material"], downstream_label)
                if down["inlet_pressure_psi"] > up["inlet_pressure_psi"]:
                    errors.append((
                        None,
                        'The downstream table is a higher pressure than the '
                        'table above. A regulator steps pressure down.'))
        if errors:
            return None
        return {
            "pipe_material":          cb_material.SelectedItem,
            "table_label":            cb_table.SelectedItem,
            "heat_content":           heat_content,
            "mid_stream_prv":         use_prv,
            "downstream_table_label": downstream_label,
        }

    return _run(win, [tb_heat], collect)


# ---------------------------------------------------------------------------
# Size Gas and One-Line: existing conditions at a new connection
# ---------------------------------------------------------------------------

_EXISTING_BODY = (
    _hint('A pipe was picked, so the run starts at the connection to existing gas '
          'piping. The existing piping is assumed not modeled. Enter it below.',
          '0,10,0,0') +
    _section('Existing gas piping') +
    '<TextBlock Text="Existing pipe size at the connection" Margin="0,0,0,3"/>'
    '<ComboBox x:Name="cb_size" Margin="0,0,0,8"/>'
    '<StackPanel Orientation="Horizontal">'
    '<TextBlock Text="Length, meter to new connection" Width="190" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_length" Width="72" Margin="0,0,6,0"'
    ' ToolTip="Developed length of the existing pipe from the meter to the new connection. Cannot be negative."/>'
    '<TextBlock Text="ft" VerticalAlignment="Center"/>'
    '</StackPanel>'
    '<StackPanel Orientation="Horizontal" Margin="0,6,0,0">'
    '<TextBlock Text="Total existing load" Width="190" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_load" Width="72" Margin="0,0,6,0"'
    ' ToolTip="Connected load of the existing appliances. Cannot be negative."/>'
    '<TextBlock Text="MBH" VerticalAlignment="Center"/>'
    '</StackPanel>' +

    _BASIS_XAML
)

_EXISTING_BASIS_ROWS = [
    ("Existing pipe",
     "Checked, never resized. It passes if its table capacity covers the "
     "existing load plus the new load. A shortfall is a warning."),
    ("Developed length",
     "The entered length is added to the new longest run. That total picks "
     "the table row for the check and for every new pipe."),
    ("Longest run",
     "The new piping is assumed to be the longest run in the system. The "
     "engineer shall verify that existing appliances farther from the meter "
     "still receive adequate gas."),
]


def show_existing_gas_dialog(title, pipe_sizes, defaults=None, use_theme=True):
    """Themed Existing Conditions box. Same arguments and return value as
    ui_helpers.show_existing_gas_dialog."""
    win = _build_window(
        title, 'Existing Conditions',
        'The existing gas piping from the meter to the new connection.',
        _EXISTING_BODY, use_theme, width=520)

    cb_size = _named(win, 'cb_size')
    tb_length = _named(win, 'tb_length')
    tb_load = _named(win, 'tb_load')

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

    _fill_basis(win, _EXISTING_BASIS_ROWS)

    def collect(errors):
        length = _number(tb_length, 'Length', errors, kind='nonneg')
        load = _number(tb_load, 'Total existing load', errors, kind='nonneg')
        if cb_size.SelectedItem is None:
            errors.append((None, 'Pick the existing pipe size.'))
        if errors:
            return None
        return {
            "size_nominal": cb_size.SelectedItem,
            "length_ft":    length,
            "load_mbh":     load,
        }

    return _run(win, [tb_length, tb_load], collect)
