# -*- coding: utf-8 -*-
"""
Duct Velocity.pushbutton/script.py  --  HVAC Phase 1

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
          Slot diffusers are the exception: the design standard publishes no
          duct/neck size breakpoints for them, so a slot branch falls back to
          the same velocity/friction check a main gets.

The diffuser/duct height clearance check still applies to both, and still
overrides either verdict - it is an installability problem, not a sizing one.

Which columns appear in the output tables is chosen in the settings dialog;
the same selection drives both the drafting-view schedule placed on the sheet
and the console table.

Run HVAC Diagnose first to verify the network and CFM values, then run this.

IronPython 2.7 / pyRevit  --  no f-strings, no walrus, no nonlocal.
"""

import os
import sys
import math
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
    Viewport, ViewDuplicateOption, ElementId, XYZ,
    OverrideGraphicSettings, Color,
    TextNote, TextNoteOptions, TextNoteType,
    Line, ViewDrafting, ViewFamilyType, ViewFamily,
    FamilySymbol, StorageType,
    FilledRegion, FilledRegionType, CurveLoop
)
from Autodesk.Revit.UI.Selection import ObjectType

from System.Windows import (
    Window, WindowStartupLocation, Thickness,
    HorizontalAlignment, VerticalAlignment, SizeToContent, TextWrapping
)
from System.Windows.Controls import (
    Grid, Label, TextBox, Button, StackPanel,
    ColumnDefinition, RowDefinition, Orientation,
    Separator, TextBlock, CheckBox
)
from System.Windows.Media import SolidColorBrush, Colors
from System.Windows import FontWeights
from System.Windows import GridLength

doc    = __revit__.ActiveUIDocument.Document
uidoc  = __revit__.ActiveUIDocument
output = script.get_output()

# Add lib/ to path
_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

import hvac_graph
from revit_helpers import eid_int

# ── SMACNA colors ─────────────────────────────────────────────────────────────
GREEN  = Color(0,   200,  0)
YELLOW = Color(255, 215,  0)
RED    = Color(210,  40, 40)
PURPLE = Color(140,  40, 200)
GRAY   = Color(160, 160, 160)

_COLOR_MAP = {'GREEN': GREEN, 'YELLOW': YELLOW, 'RED': RED, 'PURPLE': PURPLE, 'GRAY': GRAY}


# ── output table columns ───────────────────────────────────────────────────────
# ONE definition list drives BOTH output tables: the drafting-view schedule
# placed on the sheet and the console print_code table. They are rendered by
# different code (detail lines + TextNotes vs. padded monospace text) and have
# drifted apart before, so the column set, the headers and the cell values all
# come from here and from _row_cells() rather than being written out twice.
#
#   view_w     column width in feet at the schedule view's 1:1 scale
#              (= printed inches / 12)
#   console_w  column width in monospace characters
#   default_on starting checkbox state in the settings dialog (selectable
#              columns only - see _FIXED_COLUMN_KEYS)
#
# Some columns are always in the tables and have no checkbox; the rest are
# optional. Defaults for the optional ones are the firm's chosen starting
# state, not a stored preference.
_COLUMN_DEFS = [
    # key          header                              view_w  console_w  default_on
    ('num',       '#',                                  0.050,    3,  True),
    ('status',    'Status',                             0.120,    7,  True),
    ('reason',    'Reason',                             0.230,   32,  True),
    ('role',      'Duct Role',                          0.150,   18,  True),
    ('size',      'Size',                               0.100,    9,  True),
    ('required',  'Required Size',                      0.140,   14,  True),
    ('cfm',       'Total CFM through Duct',             0.200,   22,  False),
    ('fpm',       'Actual FPM / Max FPM',               0.220,   21,  True),
    ('fric',      'Actual Fric / Max Fric (iwc/100)',   0.300,   32,  False),
    ('length',    'Length (ft)',                        0.090,   11,  False),
    ('fricloss',  'Friction Loss (iwc)',                0.130,   19,  False),
]

# Always shown, no checkbox. '#' is also the number printed in the keynote
# circles placed in the view, so it stays in the table to keep that
# correlation.
_FIXED_COLUMN_KEYS = frozenset(
    ['num', 'status', 'reason', 'role', 'size', 'required', 'fpm'])


def _optional_columns():
    """The _COLUMN_DEFS entries that get a checkbox, in definition order."""
    return [c for c in _COLUMN_DEFS if c[0] not in _FIXED_COLUMN_KEYS]


def _selected_columns(selected_keys):
    """The _COLUMN_DEFS entries to render: every fixed column plus the
    optional ones the user checked, in definition order.

    Both renderers call this so column ORDER is defined once too, not just
    the set - a dialog that returned an unordered set would otherwise let the
    two tables order themselves differently.
    """
    return [c for c in _COLUMN_DEFS
            if c[0] in _FIXED_COLUMN_KEYS or c[0] in selected_keys]


# ── velocity settings dialog ───────────────────────────────────────────────────
def show_velocity_settings_dialog():
    """WPF dialog — Outside Air scope, per-system max velocity + friction,
    a yellow/red tolerance %, and which columns the output tables show.

    Returns ({sys_class: (max_fpm, max_friction_inwc)}, tol_pct, include_oa,
    selected_column_keys) or None.

    The velocity and friction limits here apply to MAIN ducts only. Branch
    ducts (a run feeding exactly one terminal) are sized against the published
    diffuser tables instead and ignore every number on this dialog — see the
    module docstring. Slot-diffuser branches are the one exception and do use
    these limits.

    include_oa is False by default (equipment-level: Supply Air +
    Return Air only — any mechanical equipment node's Outside Air branch is
    skipped entirely by hvac_graph.traverse(), not sized/colored, for
    projects where the OA network isn't modeled to completion). Checking
    the Outside Air box switches to system-level: full traversal including
    OA, for VAV/FCU networks where OA is actually being traced.

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
    ROWS = [
        ('Supply Air',   800,  0.08),
        ('Return Air',   600,  0.05),
        ('Exhaust Air',  600,  0.05),
        ('Outside Air',  600,  0.05),
        ('Transfer Air', 400,  0.05),
    ]
    DEFAULT_TOL_PCT = 10     # yellow band: ±this % around max

    result    = [None]
    vel_boxes  = {}   # row_idx -> TextBox (velocity)
    fric_boxes = {}   # row_idx -> TextBox (friction)
    gpct_box   = [None]

    WIN_WIDTH   = 520
    CONTENT_W   = WIN_WIDTH - 28 - 20   # win width minus outer margin minus a little slack

    win = Window()
    win.Title  = 'Duct Velocity Settings'
    win.Width  = WIN_WIDTH
    win.SizeToContent = SizeToContent.Height
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    outer = StackPanel()
    outer.Margin = Thickness(14)

    # ── Outside Air checkbox: equipment-level (SA+RA only, default) vs ──────
    # ── system-level (traces OA too) ─────────────────────────────────────
    cb_oa = CheckBox()
    cb_oa_text = TextBlock()
    cb_oa_text.Text = ('Outside Air / system-level (AHU/DOAS) — PENDING, under development. '
                        'Not available yet: the tool runs equipment-level only '
                        '(Supply + Return Air, never goes upstream).')
    cb_oa_text.TextWrapping = TextWrapping.Wrap
    cb_oa_text.Width = CONTENT_W - 20
    cb_oa.Content = cb_oa_text
    cb_oa.IsChecked = False
    cb_oa.IsEnabled = False
    cb_oa.Margin  = Thickness(2, 0, 0, 2)
    outer.Children.Add(cb_oa)

    oa_sep = Separator()
    oa_sep.Margin = Thickness(0, 8, 0, 10)
    outer.Children.Add(oa_sep)

    intro = Label()
    intro.Content = 'Max velocity (FPM) and pressure drop (in. wc/100 ft) per system:'
    intro.Margin  = Thickness(0, 0, 0, 8)
    outer.Children.Add(intro)

    # 3-column grid: system | velocity | friction
    grid = Grid()
    for w in (150, 130, 150):
        cd = ColumnDefinition()
        cd.Width = GridLength(w)
        grid.ColumnDefinitions.Add(cd)
    for _ in range(len(ROWS) + 1):
        rd = RowDefinition()
        rd.Height = GridLength(32)
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

    col_hdr = TextBlock()
    col_hdr.Text = 'Optional output table columns'
    col_hdr.FontWeight = FontWeights.Bold
    col_hdr.Foreground = SolidColorBrush(Colors.DimGray)
    col_hdr.Margin = Thickness(0, 0, 0, 4)
    outer.Children.Add(col_hdr)

    col_grid = Grid()
    _COL_PICKER_COLS = 2
    for _ in range(_COL_PICKER_COLS):
        cd = ColumnDefinition()
        cd.Width = GridLength(CONTENT_W / float(_COL_PICKER_COLS))
        col_grid.ColumnDefinitions.Add(cd)
    # Ceiling division, written the Python 2 way (// on ints floors).
    _optional = _optional_columns()
    _col_rows = (len(_optional) + _COL_PICKER_COLS - 1) // _COL_PICKER_COLS
    for _ in range(_col_rows):
        rd = RowDefinition()
        rd.Height = GridLength(22)
        col_grid.RowDefinitions.Add(rd)

    col_boxes = {}   # column key -> CheckBox
    for i, (key, header, _vw, _cw, default_on) in enumerate(_optional):
        cb = CheckBox()
        cb_txt = TextBlock()
        cb_txt.Text = header
        cb_txt.TextWrapping = TextWrapping.NoWrap
        cb.Content   = cb_txt
        cb.IsChecked = default_on
        cb.Margin    = Thickness(2, 2, 2, 2)
        cb.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(cb, i // _col_rows)
        Grid.SetRow(cb, i % _col_rows)
        col_grid.Children.Add(cb)
        col_boxes[key] = cb

    outer.Children.Add(col_grid)

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

    hdr = TextBlock()
    hdr.Text = 'Calculation Basis'
    hdr.FontWeight = FontWeights.Bold
    hdr.Foreground = SolidColorBrush(Colors.DimGray)
    hdr.Margin = Thickness(0, 0, 0, 4)
    outer.Children.Add(hdr)

    _info_row('Applies to:',      'main ducts only — branches are sized against the diffuser tables')
    _info_row('Pressure drop:',   'Darcy-Weisbach')
    _info_row('Friction factor:', 'Altshul-Tsal  (ASHRAE approx. to Colebrook-White)')
    _info_row('Air density:',     u'0.0750 lb/ft³  (standard, 68°F, sea level)')
    _info_row('Duct roughness:',  u'ε = 0.0003 ft  (galvanized steel)')
    _info_row('Purple (oversized):', 'a smaller standard size exists that stays within max FPM + friction')

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
                forms.alert('Green threshold must be between 0 and 100.', title='Invalid Input')
                return
            out = {}
            for i, (sys_class, _, _) in enumerate(ROWS):
                max_fpm  = float(vel_boxes[i].Text)
                max_fric = float(fric_boxes[i].Text)
                if max_fpm <= 0 or max_fric <= 0:
                    forms.alert('All values must be greater than 0.', title='Invalid Input')
                    return
                out[sys_class] = (max_fpm, max_fric)
            include_oa = bool(cb_oa.IsChecked)
            selected_cols = set(k for k, cb in col_boxes.items() if bool(cb.IsChecked))
            result[0] = (out, gpct, include_oa, selected_cols)
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

    win.Content = outer
    win.ShowDialog()
    return result[0]


# ── helpers ────────────────────────────────────────────────────────────────────
_PRIORITY = {'RED': 4, 'YELLOW': 3, 'PURPLE': 2, 'GREEN': 1, 'GRAY': 0}

# Minimum transition clearance between a duct and a genuinely smaller real
# downstream neighbour, inches.
_CLEARANCE_MARGIN_IN = 2.0


def _fails_height_clearance(dr, downstream_height_in):
    """True if this duct fails the diffuser/duct height clearance rule.

    Only fires where the real downstream neighbour is actually SMALLER than
    this duct (a genuine size reduction) and the step down is too shallow to
    build a transition into — an equal or larger downstream connection is a
    continuous run and needs no margin at all.

    Shared by the main and the branch labelling paths deliberately: this is a
    hard installability problem, orthogonal to whichever sizing method judged
    the duct, and it overrides both of them. Keeping it in one function is
    what stops the two paths' clearance behaviour drifting apart.
    """
    my_height = hvac_graph.effective_height_in(dr.elem)
    if my_height is None or downstream_height_in is None:
        return False
    return (downstream_height_in < my_height
            and my_height < downstream_height_in + _CLEARANCE_MARGIN_IN)


def _duct_label(dr, custom_limits, tol_pct, downstream_height_in=None):
    """Worst of velocity and friction checks (one-sided, unchanged), then
    downgraded to PURPLE if the duct is GREEN but oversized (a smaller
    standard size would still satisfy max FPM and max friction), then
    overridden to RED regardless of the above if the duct fails the
    diffuser/duct height clearance rule (its height doesn't clear its real
    downstream neighbor's height by the required 2" margin, wherever that
    neighbor is actually smaller — a hard installability problem, checked
    independent of velocity/friction).

    MAIN DUCTS ONLY as of the branch/diffuser split, with one exception: a
    branch feeding a slot diffuser also comes through here, because the design
    standard publishes no duct/neck size table to check a slot branch against.
    Behaviour is unchanged from before the split — see _branch_duct_label()
    for what every other branch gets instead, and _label_duct() for the
    dispatch between them.

      Green  : cfm <= max_cap
      Yellow : max_cap < cfm <= max_cap * (1 + tol_pct/100)
      Red    : cfm > max_cap * (1 + tol_pct/100), OR fails height clearance
      Purple : GREEN AND a smaller standard size exists (see _suggest_size_info)

    Returns (label, max_cap_cfm, reason) — reason is 'Velocity', 'Friction',
    'Oversized', 'Diffuser/Duct Clearance', or '' (GREEN/GRAY, nothing to flag).
    """
    defaults = hvac_graph.FIRM_DEFAULTS.get(dr.sys_class, (600, 0.05))
    max_fpm, max_friction = custom_limits.get(dr.sys_class, defaults)
    tol_fac = tol_pct / 100.0

    # Velocity check (CFM vs capacity at max FPM)
    if dr.cfm <= 0 or dr.area_ft2 <= 0:
        vel_label = 'GRAY'
        max_cap   = 0.0
    else:
        max_cap = max_fpm * dr.area_ft2
        red_cap = max_cap * (1.0 + tol_fac)
        if dr.cfm <= max_cap:
            vel_label = 'GREEN'
        elif dr.cfm <= red_cap:
            vel_label = 'YELLOW'
        else:
            vel_label = 'RED'

    # Friction check
    if dr.friction_per_100ft <= 0 or max_friction <= 0:
        fric_label = 'GRAY'
    else:
        red_fric = max_friction * (1.0 + tol_fac)
        if dr.friction_per_100ft <= max_friction:
            fric_label = 'GREEN'
        elif dr.friction_per_100ft <= red_fric:
            fric_label = 'YELLOW'
        else:
            fric_label = 'RED'

    fric_wins = _PRIORITY.get(fric_label, 0) > _PRIORITY.get(vel_label, 0)
    combined  = fric_label if fric_wins else vel_label
    reason    = ('Friction' if fric_wins else 'Velocity') if combined in ('YELLOW', 'RED') else ''

    # GREEN ducts get downgraded to PURPLE if a smaller standard size — both
    # width and height at least one nominal size down — would still satisfy
    # both max FPM and max friction. i.e. it's oversized.
    if combined == 'GREEN':
        _, sugg_area = _suggest_shrink_info(dr, custom_limits, downstream_height_in)
        if sugg_area is not None:
            combined = 'PURPLE'
            reason   = 'Oversized'

    # Diffuser/duct height clearance — hard override, independent of
    # everything above.
    if _fails_height_clearance(dr, downstream_height_in):
        combined = 'RED'
        reason   = 'Diffuser/Duct Clearance'

    return combined, max_cap, reason


# ── branch labelling (diffuser-table method) ──────────────────────────────────

_ROLE_LABELS = {
    'ceiling':  'Branch (Ceiling)',
    'sidewall': 'Branch (Sidewall)',
    'slot':     'Branch (Slot)',
}


def _duct_role_label(is_branch, branch_kind):
    """Display string for the Duct Role column.

    The diffuser category is carried into the label rather than just
    'Main'/'Branch' because it is the single most useful thing to see when a
    branch verdict looks wrong: it says which table the row was judged
    against, or that no table could be chosen at all.
    """
    if not is_branch:
        return 'Main'
    return _ROLE_LABELS.get(branch_kind, 'Branch (Unknown)')


def _branch_duct_label(dr, terminal_elem, terminal_cfm, downstream_height_in=None):
    """Label a branch duct by diffuser-table size match instead of velocity.

    All of the actual sizing logic lives in hvac_graph.branch_diffuser_check();
    this only adds the height-clearance override that applies to every duct
    regardless of how it was judged.

    Returns (label, max_cap_cfm, reason, BranchDiffuserResult). max_cap_cfm is
    always 0.0 — a branch has no velocity capacity figure, and the callers
    that carry that slot only ever read the label out of it. Status can be
    GREEN, RED or GRAY (converted to RED by _label_duct) but never YELLOW or PURPLE: there is no tolerance band
    on a size match, and a branch is never checked for oversize (the
    auto-sizing diffuser family sets the branch size).
    """
    res = hvac_graph.branch_diffuser_check(
        terminal_elem, terminal_cfm, hvac_graph.duct_installed_size_in(dr.elem))
    label  = res.status
    reason = res.reason
    if _fails_height_clearance(dr, downstream_height_in):
        label  = 'RED'
        reason = 'Diffuser/Duct Clearance'
    return label, 0.0, reason, res


def _label_duct(dr, custom_limits, tol_pct, downstream_height_in, branch_ctx):
    """Label one duct GREEN / YELLOW / RED / PURPLE, or GRAY.

    GRAY means the duct could not be judged (no airflow reaching it, no size,
    unreadable diffuser, no table). That is nearly always a broken or
    disconnected system, so it is shown gray in the view and explained in the
    legend rather than passed or failed. Gray ducts are not in the flagged
    table.
    """
    label, cap, reason, branch_res = _label_duct_checked(
        dr, custom_limits, tol_pct, downstream_height_in, branch_ctx)
    if label == 'GRAY' and not reason:
        if dr.cfm <= 0:
            reason = 'No airflow data'
        elif dr.area_ft2 <= 0:
            reason = 'No duct size data'
        else:
            reason = 'Cannot check'
    return label, cap, reason, branch_res


def _label_duct_checked(dr, custom_limits, tol_pct, downstream_height_in, branch_ctx):
    """Send one duct to whichever check applies to it.

    branch_ctx is None for a main, or (diffuser_category, terminal_elem,
    terminal_cfm) for a branch — built once per duct in main() from the
    downstream-terminal map, so the graph is only walked once.

    Returns (label, max_cap_cfm, reason, branch_result_or_None). The fourth
    element is None for anything judged the ductulator way, which is also how
    the output tables know to print real FPM/friction numbers rather than N/A.
    May return GRAY.
    """
    if branch_ctx is None:
        label, cap, reason = _duct_label(dr, custom_limits, tol_pct, downstream_height_in)
        return label, cap, reason, None

    kind, terminal_elem, terminal_cfm = branch_ctx

    if kind == 'slot':
        # No published duct/neck size breakpoints exist for slot diffusers
        # (the standard gives capacity per slot length and open-slot count
        # instead), so there is nothing to size-match against. Fall back to
        # the main check rather than report the branch as unverifiable.
        label, cap, reason = _duct_label(dr, custom_limits, tol_pct, downstream_height_in)
        return label, cap, reason, None

    if kind is None:
        # Connector geometry unreadable — no table can be chosen. Gray, not a
        # pass: an unverifiable branch must never look like a clean one.
        label  = 'GRAY'
        reason = 'Cannot classify diffuser'
        if _fails_height_clearance(dr, downstream_height_in):
            label  = 'RED'
            reason = 'Diffuser/Duct Clearance'
        return label, 0.0, reason, None

    return _branch_duct_label(dr, terminal_elem, terminal_cfm, downstream_height_in)


def _elem_name(elem):
    try:
        return elem.Name
    except Exception:
        return str(eid_int(elem.Id))


def _duct_size_label(elem):
    """Return readable size: '10"' for round/spiral, '18"x12"' for rectangular."""
    try:
        d = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if d is not None and d.AsDouble() > 0:
            return '{:.0f}"'.format(d.AsDouble() * 12.0)
        w = elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        h = elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if w and h and w.AsDouble() > 0 and h.AsDouble() > 0:
            return '{:.0f}"x{:.0f}"'.format(w.AsDouble() * 12.0, h.AsDouble() * 12.0)
    except Exception:
        pass
    return '?'


# Standard spiral/round sizes in 2" increments (inches)
_ROUND_SIZES = [6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36]
_MIN_DUCT_DIM = 6  # firm standard: never suggest below 6" for rect or spiral


def _snap_even(inches):
    """Round up to the nearest even (2" interval) duct dimension — odd sizes
    like 17" aren't stocked/installed; the next even size up (18") is used.

    Subtracts a small epsilon before ceiling so floating-point noise from the
    Revit feet<->inch round-trip (e.g. a true 10" duct stored internally as
    10.000000000000002") doesn't get bumped up to the next even size (12").
    Real fractional dimensions (e.g. 8.5") still round up correctly since the
    epsilon is far smaller than any real duct dimension tolerance.
    """
    n = int(math.ceil(inches - 1e-6))
    if n % 2 != 0:
        n += 1
    return max(n, 2)


def _suggest_shrink_info(dr, custom_limits, downstream_height_in=None):
    """Return (label, area_ft2) for a SMALLER standard duct size than what's
    installed, satisfying both velocity AND friction limits — no tolerance
    band, straight against max FPM / max friction. Returns ('-', None) if no
    strictly-smaller standard size works.

    Used for the oversized/PURPLE suggestion only. Round/spiral: smaller
    standard diameter only. Rectangular: BOTH width and height must drop by
    at least one nominal (2") step from installed — a candidate that only
    shrinks one dimension (e.g. 12x12 -> 12x10) is not a valid suggestion,
    since that's a marginal height-only trim, not a genuine downsize. Only
    a candidate like 12x12 -> 10x10, where both dimensions are a nominal
    size smaller, qualifies. All dimensions snapped to 2" intervals.

    Never suggests below the firm's 6" absolute minimum (_ROUND_SIZES/
    _MIN_DUCT_DIM). If downstream_height_in is given (the effective height of
    whatever this duct's real downstream neighbor is, from
    hvac_graph.max_downstream_height_in), also never suggests a candidate
    that would leave less than the required 2" clearance margin wherever
    downstream_height_in is actually smaller than the candidate — a
    same-size or larger downstream connection needs no transition margin.
    """
    defaults = hvac_graph.FIRM_DEFAULTS.get(dr.sys_class, (600, 0.05))
    max_fpm, max_friction = custom_limits.get(dr.sys_class, defaults)
    if dr.cfm <= 0 or max_fpm <= 0:
        return '-', None

    def _clears(candidate_in):
        if downstream_height_in is None:
            return True
        if downstream_height_in >= candidate_in:
            return True  # no reduction at this candidate, no margin needed
        return candidate_in >= downstream_height_in + 2.0

    try:
        d = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if d is not None and d.AsDouble() > 0:
            installed_d = _snap_even(d.AsDouble() * 12.0)
            for std_d in _ROUND_SIZES:
                if std_d >= installed_d:
                    break
                if not _clears(std_d):
                    continue
                area_ft2 = math.pi * (std_d / 24.0) ** 2
                vel      = dr.cfm / area_ft2
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, float(std_d))
                if vel <= max_fpm and fric <= max_friction:
                    return '{}"'.format(std_d), area_ft2
            return '-', None

        w_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        h_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if w_param and h_param and w_param.AsDouble() > 0 and h_param.AsDouble() > 0:
            w_in = _snap_even(w_param.AsDouble() * 12.0)
            h_in = _snap_even(h_param.AsDouble() * 12.0)
            # Both width and height must be at least one nominal (2") step
            # below what's installed — a candidate that only shrinks one
            # dimension isn't a genuine downsize suggestion.
            best = None
            for new_w in range(_MIN_DUCT_DIM, w_in - 2 + 1, 2):
                min_h = max(_snap_even(new_w / 4.0), _MIN_DUCT_DIM)
                max_h = min(h_in - 2, new_w * 4)
                for new_h in range(min_h, max_h + 1, 2):
                    if not _clears(new_h):
                        continue
                    area_ft2 = new_w * new_h / 144.0
                    vel      = dr.cfm / area_ft2
                    d_h      = 4.0 * new_w * new_h / (2.0 * (new_w + new_h))
                    fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h)
                    if vel <= max_fpm and fric <= max_friction:
                        if best is None or area_ft2 < best[1]:
                            best = ('{}"x{}"'.format(new_w, new_h), area_ft2)
            if best is not None:
                return best
    except Exception:
        pass
    return '-', None


def _suggest_size_info(dr, custom_limits, downstream_height_in=None):
    """Return (label, area_ft2) for the smallest standard duct size satisfying
    both velocity AND friction limits — no tolerance band, straight against
    max FPM / max friction. area_ft2 is None when no standard size could be
    found (duct needs to grow past the largest round size, or past what
    width-expansion covers for rectangular).

    Used for the undersized/YELLOW/RED suggestion (grow direction) — a truly
    undersized duct can never satisfy at a smaller height than installed
    either (smaller height only makes velocity/friction worse), so searching
    the full range from the AR floor upward is safe and never suggests a
    shrink here.
    Round/spiral: iterates standard diameters smallest-first.
    Rectangular: keeps the installed width fixed, iterates height from the
    smallest AR-valid value upward. Expands width only if no height at the
    installed width satisfies both limits within AR 4:1.
    All suggested dimensions are snapped to 2" intervals (2, 4, 6, 8...) —
    odd sizes are not stocked/installed.
    Both constraints must be satisfied — takes the binding (larger) of the two requirements.

    Never suggests below the firm's 6" absolute minimum (_ROUND_SIZES/
    _MIN_DUCT_DIM), and (same as _suggest_shrink_info) never suggests a
    candidate that would leave less than the required 2" clearance margin
    wherever downstream_height_in is actually smaller than the candidate.
    """
    defaults = hvac_graph.FIRM_DEFAULTS.get(dr.sys_class, (600, 0.05))
    max_fpm, max_friction = custom_limits.get(dr.sys_class, defaults)
    if dr.cfm <= 0 or max_fpm <= 0:
        return '-', None

    def _clears(candidate_in):
        if downstream_height_in is None:
            return True
        if downstream_height_in >= candidate_in:
            return True
        return candidate_in >= downstream_height_in + 2.0

    try:
        d = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if d is not None and d.AsDouble() > 0:
            # Round / spiral — first standard diameter satisfying both vel and friction
            for std_d in _ROUND_SIZES:
                if not _clears(std_d):
                    continue
                area_ft2 = math.pi * (std_d / 24.0) ** 2
                vel      = dr.cfm / area_ft2
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, float(std_d))
                if vel <= max_fpm and fric <= max_friction:
                    return '{}"'.format(std_d), area_ft2
            return '>{}"'.format(_ROUND_SIZES[-1]), None

        # Rectangular — keep width fixed, find smallest satisfying height
        w_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        h_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if w_param and h_param and w_param.AsDouble() > 0 and h_param.AsDouble() > 0:
            w_in = _snap_even(w_param.AsDouble() * 12.0)
            h_in = _snap_even(h_param.AsDouble() * 12.0)

            # AR 4:1 both ways bounds the height search at this fixed width:
            # too-flat (w > 4h) below, too-tall (h > 4w) above.
            min_h = max(_snap_even(w_in / 4.0), _MIN_DUCT_DIM)
            max_h = w_in * 4
            for new_h in range(min_h, max_h + 2, 2):
                if not _clears(new_h):
                    continue
                area_ft2 = w_in * new_h / 144.0
                vel      = dr.cfm / area_ft2
                d_h      = 4.0 * w_in * new_h / (2.0 * (w_in + new_h))
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h)
                if vel <= max_fpm and fric <= max_friction:
                    return '{}"x{}"'.format(w_in, new_h), area_ft2

            # No height at the installed width works within AR — expand width
            for new_w in range(int(w_in) + 2, int(w_in) + 60, 2):
                min_h2 = max(_snap_even(new_w / 4.0), _MIN_DUCT_DIM)
                max_h2 = new_w * 4
                for new_h in range(min_h2, max_h2 + 2, 2):
                    if not _clears(new_h):
                        continue
                    area_ft2 = new_w * new_h / 144.0
                    vel      = dr.cfm / area_ft2
                    d_h      = 4.0 * new_w * new_h / (2.0 * (new_w + new_h))
                    fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h)
                    if vel <= max_fpm and fric <= max_friction:
                        return '{}"x{}"'.format(new_w, new_h), area_ft2
    except Exception:
        pass
    return '-', None


def _suggest_size(dr, custom_limits, tol_pct, duct_label, downstream_height_in=None):
    """Formatted-string wrapper for schedule/table display. PURPLE (oversized)
    ducts get a shrink-only suggestion (both width and height at least one
    nominal size down); YELLOW/RED (undersized) ducts get the grow-capable
    suggestion (may need a wider duct)."""
    if duct_label == 'PURPLE':
        size, _ = _suggest_shrink_info(dr, custom_limits, downstream_height_in)
    else:
        size, _ = _suggest_size_info(dr, custom_limits, downstream_height_in)
    return size


def _row_cells(label, dr, reason, role, branch_res,
               custom_limits, tol_pct, downstream_height_in):
    """Every cell value for one flagged duct, keyed by _COLUMN_DEFS key.

    Built once per duct and handed to both renderers, so the drafting-view
    schedule and the console table cannot report different numbers for the
    same duct no matter which columns each is showing.

    '#' is deliberately NOT produced here. The two tables number their rows
    differently on purpose — the schedule's numbers match the keynote circles
    placed in the view (element-id order), while the console list is re-sorted
    worst-first — so each renderer supplies its own row index.

    A branch judged against the diffuser tables reports N/A for velocity and
    friction rather than the real numbers: those numbers played no part in its
    verdict, and printing them beside it invites the reader to conclude they
    were what failed.
    """
    defaults          = hvac_graph.FIRM_DEFAULTS.get(dr.sys_class, (600, 0.05))
    max_fpm, max_fric = custom_limits.get(dr.sys_class, defaults)

    if branch_res is not None:
        fpm_cell  = 'N/A'
        fric_cell = 'N/A'
        required  = branch_res.required_size
    else:
        required  = _suggest_size(dr, custom_limits, tol_pct, label, downstream_height_in)
        if reason == 'Diffuser/Duct Clearance':
            fpm_cell  = 'N/A'
            fric_cell = 'N/A'
        else:
            fpm_cell  = '{:.0f}/{:.0f}'.format(float(dr.fpm), float(max_fpm))
            fric_cell = '{:.3f}/{:.3f}'.format(float(dr.friction_per_100ft), float(max_fric))

    return {
        'status':    label,
        'reason':    reason,
        'role':      role,
        'size':      _duct_size_label(dr.elem),
        'required':  required,
        'cfm':       '{:.0f}'.format(dr.cfm),
        'fpm':       fpm_cell,
        'fric':      fric_cell,
        'length':    '{:.1f}'.format(dr.length_ft),
        'fricloss':  '{:.3f}'.format(dr.friction_loss_inwc),
    }


# ── Schedule table in a Drafting View ─────────────────────────────────────────

def _hline(doc, view, x0, x1, y):
    doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(x0, y, 0.0), XYZ(x1, y, 0.0)))

def _vline(doc, view, x, y0, y1):
    doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(x, y0, 0.0), XYZ(x, y1, 0.0)))


def _legend_swatch(doc, view, filled_region_type_id, x0, y0, size, color, fill_id):
    """Draw a small solid-color square at (x0, y0), size x size (ft), colored
    the same way ducts/fittings are colored in the view (matching fill pattern)."""
    pts = [XYZ(x0, y0, 0.0), XYZ(x0 + size, y0, 0.0),
           XYZ(x0 + size, y0 + size, 0.0), XYZ(x0, y0 + size, 0.0)]
    loop = CurveLoop()
    for i in range(4):
        loop.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
    fr = FilledRegion.Create(doc, filled_region_type_id, view.Id, [loop])
    ogs = OverrideGraphicSettings()
    ogs.SetSurfaceForegroundPatternColor(color)
    if fill_id != ElementId.InvalidElementId:
        ogs.SetSurfaceForegroundPatternId(fill_id)
    view.SetElementOverrides(fr.Id, ogs)
    return fr


# Color key + meaning shown in the legend, in display order.
# Wording covers both check methods, since one view mixes mains and branches.
_LEGEND_ROWS = [
    ('GREEN',  'Main: within limit.  Branch: diffuser and duct both correctly sized'),
    ('PURPLE', 'Main only: oversized — smaller standard size available'),
    ('YELLOW', 'Main only: approaching limit'),
    ('RED',    'Main: exceeds limit.  Branch: diffuser or duct undersized.  '
               'Either: fails diffuser/duct height clearance (see Reason column)'),
    ('GRAY',   'No airflow data — no CFM reaches this element; broken or disconnected system'),
]


def _build_summary_view(doc, summary_lines, flagged_rows, selected_cols,
                        source_sheet_num, tn_type_id, ts, fill_id):
    """Create a Drafting View with a System Summary block + flagged-duct table
    + color legend.

    flagged_rows: list of per-duct cell dicts from _row_cells(), already in
    the order they should appear — the same order the numbered keynote circles
    were placed in the plan view, so row 3 here is circle 3 there.
    selected_cols: set of _COLUMN_DEFS keys the user kept in the dialog.

    Returns (ViewDrafting, total_content_height_ft, table_width_ft), or
    (None, 0.0, 0.0) on failure. The width is returned rather than hardcoded
    at the call site because the column selection now changes it on every run,
    and the viewport is positioned by its centre — the caller needs the width
    to keep the table's left edge on its hand-placed anchor.

    At scale 1:1, model feet = paper feet, so all dims below are paper inches / 12.
    """
    drafting_type_id = None
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements():
        if vft.ViewFamily == ViewFamily.Drafting:
            drafting_type_id = vft.Id
            break
    if drafting_type_id is None:
        return None, 0.0, 0.0

    try:
        sched_view = ViewDrafting.Create(doc, drafting_type_id)
        sched_view.Scale = 1
        base_name = 'Duct Schedule - DV-{} - {}'.format(source_sheet_num, ts)
        try:
            sched_view.Name = base_name
        except Exception:
            sched_view.Name = base_name + ' (2)'

        # ── Layout (ft at 1:1 = inches on paper / 12) ─────────────────────
        ox, oy = 0.0, 0.0   # top-left origin
        PAD     = 0.004      # text inset from cell edge (~1/24")
        HEAD_H  = 0.030      # header row height  (~3/8")
        ROW_H   = 0.022      # data row height    (~1/4")
        SUM_ROW_H = 0.018    # summary line height (~1/5")

        # Columns come from the shared _COLUMN_DEFS, filtered to the user's
        # selection — same list, same order, same headers as the console
        # table. The total width also sets the summary block width, and is
        # returned to the caller so the viewport can stay left-anchored.
        cols        = _selected_columns(selected_cols)
        col_headers = [c[1] for c in cols]
        col_widths  = [c[2] for c in cols]
        # sum(..., 0.0) forces float — sum([]) returns int 0 in Python 2.7
        total_w     = sum(col_widths, 0.0)

        opts = TextNoteOptions(tn_type_id)
        y_cursor = oy

        # ── System Summary block ────────────────────────────────────────────
        if summary_lines:
            sum_top = y_cursor
            for line in summary_lines:
                y_cursor -= SUM_ROW_H
            sum_bottom = y_cursor

            _hline(doc, sched_view, ox, ox + total_w, sum_top)
            _hline(doc, sched_view, ox, ox + total_w, sum_bottom)
            _vline(doc, sched_view, ox, sum_top, sum_bottom)
            _vline(doc, sched_view, ox + total_w, sum_top, sum_bottom)

            row_y = sum_top
            for line in summary_lines:
                row_y -= SUM_ROW_H
                TextNote.Create(doc, sched_view.Id,
                                XYZ(ox + PAD, row_y + SUM_ROW_H - PAD, 0.0),
                                line, opts)

        # ── Flagged ducts table ──────────────────────────────────────────────
        if flagged_rows:
            table_top = y_cursor
            total_h   = HEAD_H + ROW_H * len(flagged_rows)

            col_xs = [ox]
            for w in col_widths:
                col_xs.append(col_xs[-1] + w)

            row_tops = [table_top]
            row_tops.append(table_top - HEAD_H)
            for _ in range(len(flagged_rows)):
                row_tops.append(row_tops[-1] - ROW_H)

            for y in row_tops:
                _hline(doc, sched_view, ox, ox + total_w, y)
            for x in col_xs:
                _vline(doc, sched_view, x, table_top, table_top - total_h)

            for ci, header in enumerate(col_headers):
                TextNote.Create(doc, sched_view.Id,
                                XYZ(col_xs[ci] + PAD, table_top - PAD, 0.0),
                                header, opts)

            for ri, cells in enumerate(flagged_rows):
                row_y = row_tops[ri + 1] - PAD
                for ci, col in enumerate(cols):
                    # '#' is this table's own row index — it matches the
                    # numbered keynote circle placed on the same duct in the
                    # plan view, which is why it is not baked into the cell
                    # dict (the console table numbers the same ducts
                    # differently).
                    cell_text = str(ri + 1) if col[0] == 'num' else cells.get(col[0], '')
                    # Blank cells are real (e.g. no Reason on a passing row), and
                    # TextNote.Create rejects an empty string — leave the cell
                    # empty rather than write a placeholder into the drawing.
                    if not cell_text:
                        continue
                    TextNote.Create(doc, sched_view.Id,
                                    XYZ(col_xs[ci] + PAD, row_y, 0.0),
                                    cell_text, opts)
            y_cursor = table_top - total_h

        # ── Color legend ─────────────────────────────────────────────────────
        LEGEND_HDR_H = 0.026     # header row height (~5/16")
        LEGEND_ROW_H = 0.020     # legend row height (~1/4")
        SWATCH       = 0.014     # swatch square size (~1/6")

        frt_id = None
        for frt in FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements():
            frt_id = frt.Id
            break

        TextNote.Create(doc, sched_view.Id, XYZ(ox + PAD, y_cursor - PAD, 0.0), 'COLOR LEGEND', opts)
        y_cursor -= LEGEND_HDR_H

        if frt_id is not None:
            for color_key, meaning in _LEGEND_ROWS:
                color = _COLOR_MAP[color_key]
                sw_y0 = y_cursor - SWATCH - (LEGEND_ROW_H - SWATCH) / 2.0
                _legend_swatch(doc, sched_view, frt_id, ox + PAD, sw_y0, SWATCH, color, fill_id)
                TextNote.Create(doc, sched_view.Id,
                                XYZ(ox + PAD + SWATCH + PAD * 2.0, y_cursor - PAD, 0.0),
                                '{} — {}'.format(color_key.title(), meaning), opts)
                y_cursor -= LEGEND_ROW_H

        return sched_view, (oy - y_cursor), total_w

    except Exception:
        return None, 0.0, 0.0


# ── main ───────────────────────────────────────────────────────────────────────
def main():
    output.print_md('## Duct Velocity Visualizer')
    output.print_md('_Tip: run HVAC Diagnose first to verify CFM values and network._')
    output.print_md('---')

    # 1. Validate active view
    active_view = doc.ActiveView
    if active_view.ViewType != ViewType.FloorPlan:
        forms.alert(
            'Open a floor plan view first, then run Duct Velocity.',
            title='Wrong View Type', exitscript=True
        )

    # 2. Velocity + friction settings dialog
    dialog_result = show_velocity_settings_dialog()
    if dialog_result is None:
        output.print_md('**Cancelled.**')
        return
    custom_limits, tol_pct, include_oa, selected_cols = dialog_result
    output.print_md('Scope: **{}**'.format(
        'System-level (Supply, Return, Outside Air — upstream and downstream)' if include_oa
        else 'Equipment-level (Supply + Return Air only — never travels upstream)'))

    # 3. Find AHUs in active view and let user pick systems
    equip_in_view = list(FilteredElementCollector(doc, active_view.Id)
                         .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                         .WhereElementIsNotElementType())

    if equip_in_view:
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

        selected_names = forms.SelectFromList.show(
            sorted(name_to_elem.keys()),
            title='Select AHU Systems to Visualize',
            multiselect=True,
            button_name='Run Duct Velocity'
        )
        if not selected_names:
            output.print_md('**Cancelled.**')
            return
        sel_elems = [name_to_elem[n] for n in selected_names]
    else:
        # No equipment found in view — fall back to manual pick
        output.print_md('No mechanical equipment found in active view. Pick an element manually.')
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                'Select any duct, air terminal, or AHU in the system to visualize'
            )
        except Exception:
            output.print_md('**Cancelled.**')
            return
        sel_elems = [doc.GetElement(ref.ElementId)]

    # 4. Traverse each system and merge results
    all_duct_results = {}   # ElementId -> DuctResult  (worst-wins on overlap)
    all_nodes        = {}   # merged for fitting adjacency
    all_children     = {}   # merged
    all_root_ids     = []   # one per successfully traversed system
    ahu_labels       = []
    ahu_totals       = []   # (ahu_name, ahu_id, discharge_cfm, terminal_count, other_class_cfm)
    all_terminals    = {}   # int_id -> (cfm, sys_class, family_name)  (dedup across AHUs)
    all_zero_terms   = set()
    all_missing_flow = set()

    for sel_elem in sel_elems:
        output.print_md('Traversing **{}** (id {})...'.format(
            _elem_name(sel_elem), eid_int(sel_elem.Id)))
        net = hvac_graph.build_network(sel_elem, doc, equipment_level=(not include_oa))

        if net.errors:
            for e in net.errors:
                output.print_md('- :warning: {}'.format(e))
            continue

        ahu_labels.append('{} (id {})'.format(_elem_name(net.root), eid_int(net.root.Id)))

        if net.warnings:
            for w in net.warnings:
                output.print_md(':warning: {}'.format(w))

        output.print_md('  {} ducts  |  {} terminals'.format(
            len(net.duct_results), len(net.terminal_cfms)))

        root_id = eid_int(net.root.Id)
        all_root_ids.append(root_id)
        ahu_totals.append((
            _elem_name(net.root), root_id,
            net.equipment_discharge_cfm(root_id), len(net.terminal_cfms),
            net.equipment_other_class_cfm(root_id)))

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
        output.print_md('**No duct results — check errors above.**')
        return

    output.print_md('**Total: {} systems  |  {} ducts**'.format(
        len(ahu_labels), len(all_duct_results)))

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
    # KNOWN LIMITATION, left as-is deliberately (Colin, 2026-09-23): a piece of
    # equipment feeding only ONE terminal — no AHU/DOAS trunk splitting to
    # multiple VAV/VRF/FCU taps — has its entire run, trunk included,
    # classified as a single branch and gets no velocity/friction check at
    # all. The normal case (an AHU/DOAS main with multiple taps downstream)
    # is unaffected. Not special-cased for now — under development.
    term_ids = set(all_terminals.keys())
    downstream_terms = {}
    for rid in all_root_ids:
        partial = hvac_graph.compute_downstream_terminal_ids(
            rid, all_nodes, all_children, term_ids)
        for nid, tset in partial.items():
            if nid in downstream_terms:
                downstream_terms[nid] = downstream_terms[nid] | tset
            else:
                downstream_terms[nid] = tset

    branch_ctx    = {}   # ElementId -> (category, terminal_elem, terminal_cfm) or None
    duct_roles    = {}   # ElementId -> Duct Role display string
    kind_cache    = {}   # terminal int_id -> category  (many ducts share one terminal)
    role_counts   = {'Main': 0, 'Branch': 0}

    for eid in all_duct_results.keys():
        tset = downstream_terms.get(eid_int(eid), set())
        term_elem = None
        if len(tset) == 1:
            term_elem = all_nodes.get(list(tset)[0])
        if term_elem is None:
            # 0 terminals downstream (a dead-end run), 2+ (a real main), or a
            # terminal that somehow isn't in the merged node map — all keep
            # the pre-existing velocity/friction treatment.
            branch_ctx[eid] = None
            duct_roles[eid] = _duct_role_label(False, None)
            role_counts['Main'] += 1
            continue
        term_id = eid_int(term_elem.Id)
        if term_id not in kind_cache:
            kind_cache[term_id] = hvac_graph.classify_diffuser_branch(term_elem)
        kind = kind_cache[term_id]
        branch_ctx[eid] = (kind, term_elem, all_terminals.get(term_id, (0.0, '', ''))[0])
        duct_roles[eid] = _duct_role_label(True, kind)
        role_counts['Branch'] += 1

    output.print_md('Duct roles: **{} main**, **{} branch** '
                    '(branch = exactly one terminal downstream).'.format(
                        role_counts['Main'], role_counts['Branch']))

    # 4b. System summary — total flow to each equipment, diffuser counts/types
    grand_total_cfm = sum(cfm for cfm, _, _ in all_terminals.values())
    sys_totals = {}    # sys_class -> [total_cfm, count]
    type_totals = {}   # family_name -> [count, total_cfm]
    for cfm, sys_class, family in all_terminals.values():
        s = sys_totals.setdefault(sys_class, [0.0, 0])
        s[0] += cfm
        s[1] += 1
        t = type_totals.setdefault(family, [0, 0.0])
        t[0] += 1
        t[1] += cfm

    summary_lines = ['SYSTEM SUMMARY']
    for name, aid, discharge_cfm, term_count, other_class_cfm in ahu_totals:
        summary_lines.append('Equipment: {} (id {})  -  {:.0f} CFM supply air discharge ({} diffusers)'.format(
            name, aid, discharge_cfm, term_count))
        for cls in sorted(other_class_cfm.keys()):
            summary_lines.append('    + {:.0f} CFM {} feeding this equipment (not counted in discharge total)'.format(
                other_class_cfm[cls], cls))
    summary_lines.append('Grand Total Airflow: {:.0f} CFM  |  {} Diffusers  |  {} Ducts'.format(
        grand_total_cfm, len(all_terminals), len(all_duct_results)))
    for sys_class in sorted(sys_totals.keys()):
        s_cfm, s_cnt = sys_totals[sys_class]
        summary_lines.append('  {}: {:.0f} CFM  ({} diffusers)'.format(sys_class, s_cfm, s_cnt))
    summary_lines.append('Diffuser Types:')
    for family in sorted(type_totals.keys()):
        t_cnt, t_cfm = type_totals[family]
        summary_lines.append('  {}  x{}  ({:.0f} CFM)'.format(family, t_cnt, t_cfm))
    if all_zero_terms:
        summary_lines.append('WARNING: {} diffuser(s) with Flow = 0 (missing CFM data)'.format(
            len(all_zero_terms)))
    if all_missing_flow:
        summary_lines.append('WARNING: {} diffuser(s) missing a Flow parameter entirely'.format(
            len(all_missing_flow)))

    output.print_md('---')
    output.print_md('### System Summary')
    for line in summary_lines[1:]:
        output.print_md(line.strip())

    # 5. Find source sheet number
    source_sheet_num = 'NoSheet'
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        for vpid in sheet.GetAllViewports():
            vp = doc.GetElement(vpid)
            if vp is not None and vp.ViewId == active_view.Id:
                source_sheet_num = sheet.SheetNumber
                break

    output.print_md('Source sheet: **{}**'.format(source_sheet_num))

    # 6. Title block and solid fill
    tb_list = list(FilteredElementCollector(doc)
                   .OfCategory(BuiltInCategory.OST_TitleBlocks)
                   .WhereElementIsElementType())
    tb_id   = tb_list[0].Id if len(tb_list) > 0 else ElementId.InvalidElementId
    fill_id = hvac_graph.solid_fill_pattern_id(doc)

    # 7. Text height: aim for 5/64" printed size at the view's print scale
    view_scale = getattr(active_view, 'Scale', 48)
    text_h_ft  = (5.0 / (64.0 * 12.0)) * float(view_scale)

    # 8. Transaction: copy view → color overrides → FPM annotations → sheet
    t = Transaction(doc, 'Duct Velocity Visualizer')
    t.Start()
    ts = datetime.datetime.now().strftime('%Y-%m-%d-%H%M%S')

    try:
        # Copy floor plan
        new_vid  = active_view.Duplicate(ViewDuplicateOption.Duplicate)
        new_view = doc.GetElement(new_vid)
        base_name = 'Ducting Velocities - {} - {}'.format(source_sheet_num, ts)
        try:
            new_view.Name = base_name
        except Exception:
            new_view.Name = base_name + ' (2)'

        # Color overrides — worst of velocity check and friction check
        counts         = {'GREEN': 0, 'YELLOW': 0, 'RED': 0, 'PURPLE': 0, 'GRAY': 0}
        clearance_count = 0
        # eid -> (label, green_cap_cfm, reason) for fittings + annotations + schedule
        duct_labels  = {}
        # eid -> BranchDiffuserResult, or None for anything judged the
        # ductulator way. Also what tells the output tables whether to print
        # real FPM/friction numbers or N/A.
        branch_results = {}
        # eid -> effective height (in) of this duct's real downstream neighbor,
        # precomputed once here so _duct_label/_suggest_size don't need graph access
        downstream_heights = {}

        for eid, dr in all_duct_results.items():
            downstream_h = hvac_graph.max_downstream_height_in(eid_int(eid), all_nodes, all_children)
            downstream_heights[eid] = downstream_h
            label, green_cap, reason, branch_res = _label_duct(
                dr, custom_limits, tol_pct, downstream_h, branch_ctx.get(eid))
            duct_labels[eid]    = (label, green_cap, reason)
            branch_results[eid] = branch_res
            if reason == 'Diffuser/Duct Clearance':
                clearance_count += 1
            counts[label] = counts.get(label, 0) + 1
            color = _COLOR_MAP[label]
            ogs   = OverrideGraphicSettings()
            ogs.SetSurfaceForegroundPatternColor(color)
            if fill_id != ElementId.InvalidElementId:
                ogs.SetSurfaceForegroundPatternId(fill_id)
            ogs.SetProjectionLineColor(color)
            new_view.SetElementOverrides(eid, ogs)

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
            if not (hvac_graph.is_fitting_or_accessory(elem)
                    or hvac_graph.is_terminal(elem)):
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
            color = _COLOR_MAP[worst]
            ogs   = OverrideGraphicSettings()
            ogs.SetSurfaceForegroundPatternColor(color)
            if fill_id != ElementId.InvalidElementId:
                ogs.SetSurfaceForegroundPatternId(fill_id)
            ogs.SetProjectionLineColor(color)
            new_view.SetElementOverrides(elem.Id, ogs)
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
            if lbl not in ('YELLOW', 'RED', 'PURPLE'):
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

        for idx, (lbl, dr, _cells) in enumerate(flagged_items, 1):
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

        # Output sheet
        new_sheet             = ViewSheet.Create(doc, tb_id)
        new_sheet.SheetNumber = 'DV-{}-{}'.format(source_sheet_num, ts)
        # Sheet/View names share one uniqueness pool in Revit — new_view above
        # already claimed base_name verbatim, so reusing it here always collides.
        new_sheet.Name = base_name + ' (Sheet)'

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

        # System Summary + flagged-duct table + legend, placed as second viewport on sheet
        if tn_type_id is not None:
            sched_view, content_h, total_w = _build_summary_view(
                doc, summary_lines, [item[2] for item in flagged_items],
                selected_cols, source_sheet_num, tn_type_id, ts, fill_id)
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
                # No title/border on this viewport — use the "None" Viewport Type if
                # present. FilteredElementCollector on OST_Viewports + WhereElementIsElementType
                # misses this system-family type entirely (confirmed via MCP inspection).
                # Best-effort only: cosmetic, must never break the whole tool if this
                # lookup misbehaves (confirmed AttributeError on vp_type.Name here even
                # though the identical query works fine outside pyRevit's IronPython).
                try:
                    for vp in FilteredElementCollector(doc).OfClass(Viewport):
                        vp_type = doc.GetElement(vp.GetTypeId())
                        if vp_type is None:
                            continue
                        vp_type_name = vp_type.Name
                        if vp_type_name is not None and vp_type_name.strip().lower() == 'none':
                            sched_vp.ChangeTypeId(vp_type.Id)
                            break
                except Exception:
                    pass

        t.Commit()
    except Exception as ex:
        t.RollBack()
        output.print_md('**Error — transaction rolled back:** {}: {}'.format(
            type(ex).__name__, str(ex)))
        output.print_md('```\n{}\n```'.format(traceback.format_exc()))
        return

    # 9. Summary
    output.print_md('---')
    output.print_md('## Done')
    output.print_md('Sheet **{}** created.'.format(new_sheet.SheetNumber))
    output.print_md('')
    output.print_md('**Mains and branches were checked by different methods.** '
                    'A duct is a branch when exactly one terminal is reachable '
                    'downstream of it, a main when two or more are.')
    output.print_md('')
    output.print_md('**Main ducts — design limits used  (green ≤ max,  yellow = within {}% above max,  red > max+{}%,  purple = a smaller standard size still fits):**'.format(
        int(tol_pct), int(tol_pct)))
    output.print_md('| System | Max Velocity | Max Friction |')
    output.print_md('| --- | --- | --- |')
    for sys_class in ('Supply Air', 'Return Air', 'Exhaust Air', 'Outside Air', 'Transfer Air'):
        defaults = hvac_graph.FIRM_DEFAULTS.get(sys_class, (600, 0.05))
        mx_fpm, mx_fric = custom_limits.get(sys_class, defaults)
        output.print_md('| {} | {:.0f} FPM | {:.3f} iwc/100 |'.format(
            sys_class, mx_fpm, mx_fric))
    output.print_md('')
    output.print_md('**Branch ducts — checked against the diffuser sizing tables '
                    '(MEP Design Standards, section 1), not against velocity or friction.** '
                    'An auto-sizing diffuser is selected on static pressure and NC rather '
                    'than duct velocity, so velocity-checking a run that feeds only that one '
                    'diffuser is redundant. Two things are checked instead: whether the '
                    'diffuser\'s own neck/face is big enough for its own CFM, and whether the '
                    'branch duct is at least as big as the diffuser connects with. Either one '
                    'failing is red; there is no yellow band on a size match. Oversize is '
                    'not checked on branches.')
    output.print_md('')
    output.print_md('Branches feeding a **slot diffuser** fall back to the main '
                    'velocity/friction check above — the design standard publishes no '
                    'duct or neck size breakpoints for slot diffusers to check against.')
    output.print_md('')
    output.print_md('| Color | Ducts | Fittings, Accessories & Diffusers | Meaning |')
    output.print_md('| --- | --- | --- | --- |')
    output.print_md('| Green  | {} | {} | Main: within limit. Branch: diffuser and duct both correctly sized |'.format(
        counts.get('GREEN',  0), fitting_counts.get('GREEN',  0)))
    output.print_md('| Purple | {} | {} | Main only: oversized (low velocity, review for cost) |'.format(
        counts.get('PURPLE', 0), fitting_counts.get('PURPLE', 0)))
    output.print_md('| Yellow | {} | {} | Main only: approaching limit |'.format(
        counts.get('YELLOW', 0), fitting_counts.get('YELLOW', 0)))
    output.print_md('| Red    | {} | {} | Main: exceeds limit. Branch: diffuser or duct undersized |'.format(
        counts.get('RED',    0), fitting_counts.get('RED',    0)))
    output.print_md('| Gray   | {} | {} | No airflow data: broken or disconnected system (not in the flagged table) |'.format(
        counts.get('GRAY',   0), fitting_counts.get('GRAY',   0)))
    output.print_md('')
    output.print_md('**Diffuser/duct height clearance issues (flagged Red): {}**'.format(
        clearance_count))

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
        if label not in ('RED', 'YELLOW', 'PURPLE'):
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

    if flagged:
        flagged.sort(key=_flag_sort_key)

        # Same shared column definitions and same user selection as the
        # schedule view — only the fixed monospace widths differ, since this
        # table is padded text rather than ruled cells.
        cols = _selected_columns(selected_cols)

        def _fmt_row(cells):
            return '  '.join(str(c).ljust(col[3]) for c, col in zip(cells, cols))

        header    = _fmt_row([col[1] for col in cols])
        separator = _fmt_row(['-' * col[3] for col in cols])
        rows      = [header, separator]

        for idx, (label, fpm, cells) in enumerate(flagged, 1):
            rows.append(_fmt_row([
                str(idx) if col[0] == 'num' else cells.get(col[0], '')
                for col in cols
            ]))

        output.print_md('')
        output.print_md('### Flagged Ducts')
        output.print_code('\n'.join(rows))

    uidoc.ActiveView = new_sheet


main()
