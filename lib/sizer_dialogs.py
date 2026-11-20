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
    win = ui_helpers.rja_wpf_window(xaml)
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

# No "This run will" check boxes: sizing is the only job, and the WSFU
# take-off drafting view belongs to Water Performance (Colin, 2026-10-21).
_WATER_BODY = (
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

    project = _init_project(win, project_info)
    recirc = _init_recirc(win, return_detection)

    _fill_basis(win, water_tables.basis_of_design_rows())
    _fill_chart(win, "Firm pipe sizing chart", water_tables.sizing_chart())

    def collect(errors):
        values = recirc(errors)
        if errors:
            return None
        values.update(project())
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
    _hint('Puts the water supply fixture unit take-off on a drafting view, placed on '
          'the Water Performance sheet under the flagged pipes table.',
          '20,0,0,0') +

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
# Duct Performance: ceiling exhaust fan
# ---------------------------------------------------------------------------

_FAN_BODY = (
    _section('Ceiling exhaust fan', 10) +
    '<TextBlock x:Name="tb_fan" FontWeight="SemiBold" TextWrapping="Wrap"/>' +
    _hint('Its only duct connection is Exhaust Air, out, so it is read as a '
          'ceiling exhaust fan. The fan carries no airflow in the model.',
          '0,2,0,8') +
    '<StackPanel Orientation="Horizontal" Margin="0,4,0,0">'
    '<TextBlock Text="Fan airflow" Width="150" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_cfm" Width="72" Margin="0,0,6,0"'
    ' ToolTip="The fan schedule airflow. Shall be greater than 0."/>'
    '<TextBlock Text="CFM" VerticalAlignment="Center"/>'
    '</StackPanel>' +

    _BASIS_XAML
)

_FAN_BASIS_ROWS = [
    ("Discharge duct",
     "Judged against the Exhaust Air main limits. The one end of the run is "
     "the wall or roof cap, priced as a known drop (Advanced settings, "
     "component drops)."),
    ("External static",
     "Duct friction, fittings at the discharge velocity, the fan system "
     "effect at the fan outlet, the cap, plus the safety factor. The engineer "
     "checks it against the fan selection."),
]


def show_exhaust_fan_dialog(title, fan_label, cfm_default, use_theme=True):
    """Themed ceiling exhaust fan airflow (Colin, 2026-11-03: no rated
    static; the engineer checks the fan selection). Returns {'cfm'} or None."""
    win = _build_window(title, 'Ceiling Exhaust Fan',
                        'Enter the fan airflow.',
                        _FAN_BODY, use_theme, width=480)
    tb_cfm = _named(win, 'tb_cfm')
    _named(win, 'tb_fan').Text = fan_label
    if cfm_default:
        tb_cfm.Text = '{:g}'.format(round(float(cfm_default), 1))
    _fill_basis(win, _FAN_BASIS_ROWS)

    def collect(errors):
        cfm = _number(tb_cfm, 'Fan airflow', errors)
        if errors:
            return None
        return {'cfm': cfm}

    return _run(win, [tb_cfm], collect)


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


def _duct_vav_rows_xaml():
    """The MP velocity and friction rows, from vav_checks.VAV_FIELDS[:2] so
    the labels and units live in one place. Box names tb_vav_<key>."""
    import vav_checks
    out = ''
    for i, (key, label, unit, _default, _kind) in enumerate(vav_checks.VAV_FIELDS[:2]):
        out += ('<StackPanel Orientation="Horizontal" Margin="0,%d,0,0">'
                '<TextBlock Text="%s" Width="170" VerticalAlignment="Center"/>'
                '<TextBox x:Name="tb_vav_%s" Width="72" Margin="0,0,6,0"'
                ' ToolTip="%s (%s). Shall be greater than 0."/>'
                '<TextBlock Text="%s" VerticalAlignment="Center"/>'
                '</StackPanel>' % (0 if i == 0 else 6, _q(label), key, _q(label), _q(unit),
                                   _q(unit)))
    return out


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

    # VAV (plan approved 2026-10-22): the MP supply between the unit and the
    # box inlets is sized to its own limits. Greyed out until ticked.
    _section('VAV') +
    '<CheckBox x:Name="cb_vav" Content="VAV System"'
    ' ToolTip="Sizes the supply mains between the unit and the VAV box inlets (medium pressure) to the MP limits below."/>'
    '<Border x:Name="pnl_vav" Style="{DynamicResource RjaGroupBorder}" Margin="20,4,0,0"'
    ' IsEnabled="False" Opacity="0.55">'
    '<StackPanel>' +
    _duct_vav_rows_xaml() +
    _hint('Found by position: every supply main between the unit and the box '
          'inlets with two or more boxes downstream is sized to these limits (EDR '
          'VAV Guide p3, p100). A box inlet runout (one box downstream) is sized to '
          'the box by the engineer and left as drawn. Ductwork after the box is '
          'sized as above. Boxes, branches and diffusers are never touched.',
          '0,6,0,0') +
    '</StackPanel>'
    '</Border>' +

    _BASIS_XAML
)


def show_size_duct_dialog(title, use_theme=True):
    """Themed Size Duct dialog.

    Returns {"limits": {system: (max fpm, max friction)}, "vav": bool} or
    None. With VAV ticked, limits also carries vav_checks.MP_KEY.
    """
    import duct_sizing
    import hvac_graph
    import vav_checks

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

    cb_vav = _named(win, 'cb_vav')
    pnl_vav = _named(win, 'pnl_vav')
    mp_boxes = []
    for key, label, _unit, default, _kind in vav_checks.VAV_FIELDS[:2]:
        tb = _named(win, 'tb_vav_' + key)
        tb.Text = vav_checks.format_default(default)
        mp_boxes.append((label, tb))

    def on_vav_toggled(sender, args):
        on = bool(cb_vav.IsChecked)
        pnl_vav.IsEnabled = on
        pnl_vav.Opacity = 1.0 if on else 0.55
    cb_vav.Checked += on_vav_toggled
    cb_vav.Unchecked += on_vav_toggled

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
        ("Start", "Pick the RTU or AHU. A duct or a VAV box is refused."),
    ] + [r for r in vav_checks.basis_rows()
         if r[0] in ('VAV, MP supply', 'VAV, LP', 'VAV, inlet')])

    def collect(errors):
        limits = {}
        for name, tb_v, tb_f in boxes:
            fpm = _number(tb_v, name + ' velocity', errors)
            fric = _number(tb_f, name + ' friction', errors)
            if fpm is not None and fric is not None:
                limits[name] = (fpm, fric)
        vav = bool(cb_vav.IsChecked)
        if vav:
            mp = [_number(tb, label, errors) for label, tb in mp_boxes]
            if None not in mp:
                limits[vav_checks.MP_KEY] = (mp[0], mp[1])
        if errors:
            return None
        return {"limits": limits, "vav": vav}

    return _run(win, [b for _n, v, f in boxes for b in (v, f)] +
                [tb for _l, tb in mp_boxes], collect)


# ---------------------------------------------------------------------------
# Size Gas and One-Line: scope, pipe material, IFGC table, heat content
# ---------------------------------------------------------------------------
#
# One dialog before the pick (Colin, 2026-11-20). The Scope section follows
# the Duct Performance Scope: radio choices, with a Scope options box under
# the selected choice. Only Connect into Existing has options (the existing
# gas piping, assumed not modeled), so the box shows under it and is
# collapsed under New Construction. The old Existing Conditions popup after
# the pick is gone.

# The heat content range, read_heat_content and the table kind rule live in
# gas_tables (no WPF), imported here, never copied (2026-11-20).
HEAT_CONTENT_MIN_BTU_PER_CF = gas_tables.HEAT_CONTENT_MIN_BTU_PER_CF
HEAT_CONTENT_MAX_BTU_PER_CF = gas_tables.HEAT_CONTENT_MAX_BTU_PER_CF
read_heat_content = gas_tables.read_heat_content
_gas_table_kind = gas_tables.table_kind


def default_propane_heat_content():
    """shared_params.DEFAULT_PROPANE_HEAT_CONTENT_BTU_PER_CF, as a float."""
    return float(shared_params.DEFAULT_PROPANE_HEAT_CONTENT_BTU_PER_CF)


def gas_heat_boxes(table_id, downstream_table_id=None):
    """Which heat content boxes the gas dialog uses (n=24/25).

    Args:
        table_id: the IFGC table picked above.
        downstream_table_id: the table after a mid-stream PRV, only when
            Mid Stream PRV is ticked.

    Returns:
        (natural_note, propane_used). natural_note is "" when the natural gas
        box is used, else the note saying why it is greyed out. propane_used
        is True when Table 402.4(37) (propane in CFH) is either table.
    """
    kinds = [(tid, _gas_table_kind(tid))
             for tid in (table_id, downstream_table_id) if tid]
    propane_used = any(k == "propane" for _t, k in kinds)
    if not kinds or any(k in ("natural", None) for _t, k in kinds):
        return "", propane_used
    parts = []
    for tid, kind in kinds:
        if kind == "mbh":
            parts.append("Table {} is published in MBH".format(tid))
        else:
            parts.append("Table {} is propane, its heat content is entered "
                         "below".format(tid))
    return "Not used: {}.".format(" and ".join(parts)), propane_used


def gas_heat_contents(natural_text, propane_text, table_id,
                      downstream_table_id=None):
    """Read the dialog's two heat content boxes (n=24/25, 34, 39, 40).

    A box in use is validated with read_heat_content. A box not in use
    (greyed or hidden) is never validated (n=40): its value is its text when
    usable, else its default, and the sizer does not use it.

    Returns:
        (natural, propane, problems): problems is a list of
        ("natural" or "propane", message).
    """
    natural_note, propane_used = gas_heat_boxes(table_id, downstream_table_id)
    problems = []

    def _one(text, label, used, default, key):
        value, message = read_heat_content(text, label)
        if used:
            if message:
                problems.append((key, message))
            return value
        return value if message is None else default

    natural = _one(natural_text, 'Heat content of gas', not natural_note,
                   float(shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF),
                   'natural')
    propane = _one(propane_text, 'Propane heat content', propane_used,
                   default_propane_heat_content(), 'propane')
    return natural, propane, problems


_GAS_SCOPE_XAML = (
    '<StackPanel x:Name="pnl_scope">' +
    _section('Scope', 10) +
    _hint('Chosen before the pick. The scope sets what the next pick shall be.',
          '0,0,0,6') +
    '<RadioButton x:Name="rb_scope_new" GroupName="gas_scope"'
    ' Content="New Construction" IsChecked="True"'
    ' ToolTip="Full new construction with a new meter. Pick the gas meter, or the first new pipe."/>' +
    _hint('Pick the gas meter, or the first new pipe. A pipe start is full new '
          'construction with a new meter.', '20,0,0,0') +
    '<RadioButton x:Name="rb_scope_existing" GroupName="gas_scope"'
    ' Content="Connect into Existing" Margin="0,8,0,0"'
    ' ToolTip="New piping connected into existing gas piping that is not modeled. Pick the first new pipe."/>' +
    _hint('Pick the first new pipe at the connection to the existing gas '
          'piping. The existing piping is assumed not modeled.', '20,0,0,0') +
    '<Border x:Name="pnl_scope_opts" Style="{DynamicResource RjaGroupBorder}"'
    ' Margin="20,4,0,4" Visibility="Collapsed">'
    '<StackPanel>'
    '<TextBlock FontWeight="SemiBold" Text="Existing gas piping"/>'
    '<StackPanel Orientation="Horizontal" Margin="0,6,0,0">'
    '<TextBlock Text="Existing pipe size at the connection" Width="210" VerticalAlignment="Center"/>'
    '<ComboBox x:Name="cb_size" Width="90" Margin="0,0,6,0"'
    ' ToolTip="Nominal size of the existing pipe at the new connection. Shall be chosen."/>'
    '<TextBlock Text="in." VerticalAlignment="Center"/>'
    '</StackPanel>'
    '<StackPanel Orientation="Horizontal" Margin="0,6,0,0">'
    '<TextBlock Text="Length, meter to new connection" Width="210" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_length" Width="90" Margin="0,0,6,0"'
    ' ToolTip="Developed length of the existing pipe from the meter to the new connection. Cannot be negative."/>'
    '<TextBlock Text="ft" VerticalAlignment="Center"/>'
    '</StackPanel>'
    '<StackPanel Orientation="Horizontal" Margin="0,6,0,0">'
    '<TextBlock Text="Total existing load" Width="210" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_load" Width="90" Margin="0,0,6,0"'
    ' ToolTip="Connected load of the existing appliances. Cannot be negative."/>'
    '<TextBlock Text="MBH" VerticalAlignment="Center"/>'
    '</StackPanel>' +
    _hint('The existing pipe is checked for the existing plus new load and is '
          'never resized.') +
    '</StackPanel>'
    '</Border>'
    '</StackPanel>'
)

_GAS_BODY = (
    _GAS_SCOPE_XAML +

    _section('Pipe and table') +
    '<TextBlock Text="Pipe material" Margin="0,0,0,3"/>'
    '<ComboBox x:Name="cb_material" Margin="0,0,0,8"/>'
    '<TextBlock Text="IFGC table" Margin="0,0,0,3"/>'
    '<ComboBox x:Name="cb_table"/>'
    '<StackPanel Orientation="Horizontal" Margin="0,8,0,0">'
    '<TextBlock Text="Heat content of gas" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_heat" Width="72" Margin="8,0,6,0"'
    ' ToolTip="Get this value from the utility. Shall be greater than 0."/>'
    '<TextBlock Text="BTU/CF" VerticalAlignment="Center"/>'
    '</StackPanel>'
    '<TextBlock x:Name="tb_heat_note" Style="{DynamicResource RjaHint}" FontSize="11"'
    ' FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,4,0,0" Visibility="Collapsed"/>' +
    _hint('RJA standard: CFH = BTU/h / Heat Content of Gas. Get this value from the '
          'utility. Default %g BTU/CF is Denver. Sea level is 1000 BTU/CF.'
          % float(shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF)) +
    # Propane Table 402.4(37) is published in CFH (n=24/25, Colin
    # 2026-11-20): its own heat content box, shown only when that table is
    # picked above or downstream of a mid-stream PRV.
    '<StackPanel x:Name="pnl_propane" Margin="0,8,0,0" Visibility="Collapsed">'
    '<StackPanel Orientation="Horizontal">'
    '<TextBlock Text="Propane heat content" VerticalAlignment="Center"/>'
    '<TextBox x:Name="tb_propane_heat" Width="72" Margin="8,0,6,0"'
    ' ToolTip="Get this value from the propane supplier. Shall be between %g and %g BTU/CF."/>'
    '<TextBlock Text="BTU/CF" VerticalAlignment="Center"/>'
    '</StackPanel>' % (HEAT_CONTENT_MIN_BTU_PER_CF, HEAT_CONTENT_MAX_BTU_PER_CF) +
    _hint('Table 402.4(37) is undiluted propane published in CFH. CFH = BTU/h / '
          'propane heat content. Default %g BTU/CF; get the value from the '
          'propane supplier.' % default_propane_heat_content()) +
    '</StackPanel>' +
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

    _BASIS_XAML
)


# Appended to the gas Calculation basis rows (Size Gas and One-Line).
_EXISTING_BASIS_ROWS = [
    ("New Construction",
     "Scope default: full new construction with a new meter. The pick is the "
     "gas meter or the first new pipe. A pipe start begins the system at the "
     "picked pipe, with no existing pipe, length or load and no existing pipe "
     "check."),
    ("Connect into Existing",
     "The pick is the first new pipe at the connection to existing gas "
     "piping, assumed not modeled. The existing pipe size, the length from "
     "the meter and the existing load are entered under the scope."),
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


def heat_content_note(table_id):
    """The note shown under the heat content box, or "" when it is used.

    Propane Tables 402.4(25) to (36) publish capacity in MBH, so the sizer
    compares the MBH load directly and heat content is not used there
    (Colin, 2026-11-04). The box is greyed out and this note says why.
    Propane Table 402.4(37) is published in CFH but takes the propane heat
    content box instead (n=24/25), so the natural gas box is greyed there
    too. With a mid-stream PRV use gas_heat_boxes(), which reads both tables.
    """
    return gas_heat_boxes(table_id)[0]


def existing_pipe_sizes(table_id):
    """Nominal sizes offered for the existing pipe: the sizes of the selected
    IFGC table, smallest first, or [] when the table cannot be read."""
    if not table_id:
        return []
    try:
        return gas_tables.list_pipe_sizes(table_id)
    except ValueError:
        return []


def show_gas_picker(title, prv_option, use_theme=True):
    """Themed scope / pipe material / IFGC table dialog for Size Gas.

    Same return value as ui_helpers._show_picker: a dict with
    "pipe_material", "table_label", "heat_content", "mid_stream_prv",
    "downstream_table_label" and "existing" (None for New Construction, else
    {"size_nominal", "length_ft", "load_mbh"}), or None if cancelled. Also
    "propane_heat_content_btu_per_cf": the propane heat content entered
    when Table 402.4(37) is picked above or after the PRV, else the default
    (not used by the sizer then). "heat_content" is likewise the default
    when the natural gas box is greyed out and its text is not usable.
    """
    return _gas_picker(title, "size" if prv_option else "plain", use_theme)


def show_oneline_picker(title, use_theme=True):
    """Themed One-Line dialog (Plan B: a tool gets its own dialog). The
    scope, material, table, heat content and Mid Stream PRV choice shall be
    the ones Size Gas ran on, so the gate's size check matches. Shown BEFORE
    the pick, like Size Gas (Colin, 2026-11-18), so it carries the same
    Scope and Mid Stream PRV tick box. Same return value as
    show_gas_picker."""
    return _gas_picker(title, "oneline", use_theme)


def _gas_picker(title, mode, use_theme):
    """The gas dialog. mode: "size" (Size Gas, Scope and Mid Stream PRV tick
    box), "plain" (no Scope, no PRV), "oneline" (One-Line heading and basis,
    the same Scope and Mid Stream PRV tick box)."""
    from System.Windows import Visibility

    oneline = mode == "oneline"
    prv_option = mode in ("size", "oneline")
    scope_option = mode in ("size", "oneline")
    if oneline:
        heading = 'Gas One-Line'
        purpose = ('Draws the gas piping one-line diagram on a drafting view '
                   'from the sizes Size Gas wrote. Pick the scope, table, heat '
                   'content and Mid Stream PRV setting Size Gas ran on.')
    elif prv_option:
        heading = 'Size Gas'
        purpose = ('Sizes gas piping by the IFGC Longest Run Method and writes '
                   'the sizes into the model.')
    else:
        heading = 'Gas One-Line'
        purpose = ('Draws the gas piping one-line diagram on a drafting view, '
                   'with the pipe sizes already in the model.')
    win = _build_window(title, heading, purpose, _GAS_BODY, use_theme)

    pnl_scope = _named(win, 'pnl_scope')
    rb_scope_new = _named(win, 'rb_scope_new')
    rb_scope_existing = _named(win, 'rb_scope_existing')
    pnl_scope_opts = _named(win, 'pnl_scope_opts')
    cb_size = _named(win, 'cb_size')
    tb_length = _named(win, 'tb_length')
    tb_load = _named(win, 'tb_load')
    cb_material = _named(win, 'cb_material')
    cb_table = _named(win, 'cb_table')
    cb_prv = _named(win, 'cb_prv')
    pnl_prv = _named(win, 'pnl_prv')
    cb_downstream = _named(win, 'cb_downstream')
    tb_heat = _named(win, 'tb_heat')
    tb_heat_note = _named(win, 'tb_heat_note')
    pnl_propane = _named(win, 'pnl_propane')
    tb_propane_heat = _named(win, 'tb_propane_heat')

    if not scope_option:
        pnl_scope.Visibility = Visibility.Collapsed

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

    def populate_sizes():
        """The existing pipe sizes follow the table above. No size is
        preselected (Colin, 2026-11-20); a size already chosen is kept when
        the new table lists it."""
        chosen = cb_size.SelectedItem
        cb_size.Items.Clear()
        up = upstream_option()
        keep = -1
        for i, size in enumerate(existing_pipe_sizes(
                up["table_id"] if up else None)):
            cb_size.Items.Add(size)
            if chosen is not None and size == chosen:
                keep = i
        cb_size.SelectedIndex = keep

    def populate_table(material):
        cb_table.Items.Clear()
        for lbl in gas_tables.get_table_option_labels_for_material(material):
            cb_table.Items.Add(lbl)
        if cb_table.Items.Count > 0:
            cb_table.SelectedIndex = 0

    def table_ids():
        """(table above, table after the PRV or None). The downstream table
        counts only while Mid Stream PRV is ticked."""
        up = upstream_option()
        up_id = up["table_id"] if up else None
        down_id = None
        if (up is not None and prv_option and bool(cb_prv.IsChecked) and
                cb_downstream.SelectedItem is not None):
            try:
                down_id = gas_tables.get_table_option_by_material_and_short_label(
                    up["material"], cb_downstream.SelectedItem)["table_id"]
            except ValueError:
                down_id = None
        return up_id, down_id

    def update_heat_box():
        """Grey the natural gas box when no natural gas CFH table is
        picked, and show the propane box when 402.4(37) is (n=24/25)."""
        note, propane_used = gas_heat_boxes(*table_ids())
        tb_heat.IsEnabled = not note
        if note:
            ui_helpers.rja_clear_error(tb_heat)
        tb_heat_note.Text = note
        tb_heat_note.Visibility = (Visibility.Visible if note
                                   else Visibility.Collapsed)
        pnl_propane.Visibility = (Visibility.Visible if propane_used
                                  else Visibility.Collapsed)

    populate_table(materials[0])
    tb_heat.Text = "{:.0f}".format(
        float(shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF))
    tb_propane_heat.Text = "{:.0f}".format(default_propane_heat_content())
    update_heat_box()
    populate_sizes()

    if prv_option:
        cb_prv.Visibility = Visibility.Visible
        populate_downstream()

    def on_material_changed(sender, args):
        if cb_material.SelectedItem is not None:
            populate_table(cb_material.SelectedItem)
        update_heat_box()

    def on_table_changed(sender, args):
        populate_sizes()
        if prv_option:
            populate_downstream()
        update_heat_box()

    def on_prv_toggled(sender, args):
        on = bool(cb_prv.IsChecked)
        pnl_prv.Visibility = Visibility.Visible if on else Visibility.Collapsed
        update_heat_box()

    def on_downstream_changed(sender, args):
        update_heat_box()

    def on_scope_changed(sender, args):
        """Scope options sit under Connect into Existing, the only scope
        that has any, and are always live there."""
        on = bool(rb_scope_existing.IsChecked)
        pnl_scope_opts.Visibility = (Visibility.Visible if on
                                     else Visibility.Collapsed)

    cb_material.SelectionChanged += on_material_changed
    cb_table.SelectionChanged += on_table_changed
    cb_prv.Checked += on_prv_toggled
    cb_prv.Unchecked += on_prv_toggled
    cb_downstream.SelectionChanged += on_downstream_changed
    for rb in (rb_scope_new, rb_scope_existing):
        rb.Checked += on_scope_changed
        rb.Unchecked += on_scope_changed
    on_scope_changed(None, None)

    basis = (sizing_engine.oneline_basis_rows() if oneline
             else sizing_engine.basis_of_design_rows())
    if scope_option:
        basis = basis + _EXISTING_BASIS_ROWS
    _fill_basis(win, basis)

    def on_size_changed(sender, args):
        ui_helpers.rja_clear_error(cb_size)

    cb_size.SelectionChanged += on_size_changed

    def collect(errors):
        ui_helpers.rja_clear_error(cb_size)
        existing = None
        if scope_option and bool(rb_scope_existing.IsChecked):
            if cb_size.SelectedItem is None:
                errors.append((cb_size, 'Pick the existing pipe size at the '
                                        'connection.'))
            length = _number(tb_length, 'Length, meter to new connection',
                             errors, kind='nonneg')
            load = _number(tb_load, 'Total existing load', errors,
                           kind='nonneg')
            existing = {
                "size_nominal": cb_size.SelectedItem,
                "length_ft":    length,
                "load_mbh":     load,
            }
        # n=24/25, 34, 39, 40: only the boxes in use are validated, in the
        # 500 to 3000 BTU/CF range; a greyed or hidden box is never checked.
        heat_content, propane_heat, heat_problems = gas_heat_contents(
            tb_heat.Text, tb_propane_heat.Text, *table_ids())
        for key, message in heat_problems:
            errors.append((tb_heat if key == 'natural' else tb_propane_heat,
                           message))
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
            "propane_heat_content_btu_per_cf": propane_heat,
            "mid_stream_prv":         use_prv,
            "downstream_table_label": downstream_label,
            "existing":               existing,
        }

    return _run(win, [tb_length, tb_load, tb_heat, tb_propane_heat], collect)



# ---------------------------------------------------------------------------
# In-run messages (Plan B): the themed replacement for forms.alert
# ---------------------------------------------------------------------------

_MESSAGE_BODY = (
    '<Border x:Name="bd_kind" BorderThickness="4,0,0,0" Padding="10,6,6,6"'
    ' Margin="0,10,0,0">'
    '<TextBlock x:Name="tb_message" TextWrapping="Wrap" FontSize="12"/>'
    '</Border>'
)

# kind: (heading, accent color), the console callout colors
MESSAGE_KINDS = {
    "info": ("Note", "#0070C0"),
    "warn": ("Warning", "#E0A800"),
    "error": ("Error", "#D22828"),
}


def show_message(title, text, kind="info", yes_no=False, heading=None,
                 use_theme=True):
    """Themed message. OK only returns True; Yes/No returns True for Yes and
    None for No or close (ui_helpers.rja_alert turns None into False)."""
    from System.Windows import Visibility
    from System.Windows.Media import BrushConverter

    kind_heading, color = MESSAGE_KINDS.get(kind, MESSAGE_KINDS["info"])
    win = _build_window(title, heading or kind_heading, None, _MESSAGE_BODY,
                        use_theme, width=480)
    _named(win, 'tb_message').Text = text
    _named(win, 'bd_kind').BorderBrush = BrushConverter().ConvertFromString(color)
    btn_ok = _named(win, 'btn_ok')
    btn_cancel = _named(win, 'btn_cancel')
    if yes_no:
        btn_ok.Content = 'Yes'
        btn_cancel.Content = 'No'
    else:
        btn_cancel.Visibility = Visibility.Collapsed

    def collect(errors):
        return True

    return _run(win, [], collect)

