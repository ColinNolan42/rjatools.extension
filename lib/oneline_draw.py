# -*- coding: ascii -*-
"""
oneline_draw.py

One-Line rebuild step 6 (2026-11-18): carries out the oneline_layout draw
list in a Revit drafting view. All layout decisions are made upstream; this
module only draws.

    syms, report = ensure_symbols(doc)            # find in model, activate
    view = create_view(doc, name)                  # ViewDrafting, 1:100
    counts = draw(doc, view, prims, syms)          # lines, text, symbols

ensure_symbols and draw open no transaction of their own; the button wraps
them (one TransactionGroup, one undo). The pure helpers at the top
(pick_text_type, fallback_segments) have no Revit import
and are tested in CPython (lib/_test_oneline_draw.py).

Symbols (Colin, 2026-11-19): only families already loaded in the project
are used. The button never loads a family file: loading inside the command
left Revit 2024 and 2025 stuck at the end of the run (Grantham 4, Horse
Creek). A missing symbol is drawn as plain lines and the console lists each
as found or fallback. Insertion points, read from
the families: Equipment, Gate Valve, PRV, Meter Annotation centered; Cap a
single line on its insertion point; Squiggle centered. A project's older copy
may differ, so Equipment, Gate Valve and Meter are re-centered on their view
box after placing (CENTER_BY_BOX). The notes block uses the bordered text
type (border_text_type_id).

IronPython 2.7 and CPython 3 for the pure part. ASCII source.
"""

import math

# role -> family names to look for in the project, in order (first found
# wins). Nothing is ever loaded from disk (Colin, 2026-11-19).
SYMBOL_FAMILIES = [
    ("meter", ["RJA - P Symbols - Meter Annotation"]),
    ("equipment", ["RJA - P Symbols - Equipment"]),
    ("prv", ["RJA - P Symbols - Pressure Regulating Valve",
             "RJA - P Symbols - Pressure Reducing Valve"]),
    ("iso", ["RJA - P Symbols - Gate Valve"]),
    # the hood Ansul shutoff: the RJA solenoid valve symbol (Colin,
    # 2026-11-19), the gate valve when it is missing
    ("ansul", ["RJA - P Symbols - Solenoid Control Valve",
               "RJA - P Symbols - Gate Valve"]),
    ("cap", ["RJA- Equipment Cap", "RJA - Equipment Cap"]),
    ("squiggle", ["RJA - Squiggle", "RJA- Squiggle"]),
]

ROLE_TITLES = {"meter": "Meter", "equipment": "Equipment", "prv": "PRV",
               "iso": "Isolation valve", "ansul": "Ansul valve", "cap": "Cap", "squiggle": "Squiggle"}

NEW_LINE_STYLE = "Line 4"          # new construction pipe (as before)
# existing pipe and an existing meter's service: a named thin style, not the
# default <Lines> (Colin, 2026-11-18: match the weights for existing vs new)
EXISTING_LINE_STYLE = "Line 2"
TEXT_TYPE_NAME = '3/32" Arial'
TEXT_SIZE_FT = 0.09375 / 12.0
# the notes block sits in a bordered text box (Colin, 2026-11-18): the base
# type's " - Border" sibling, made by duplicating the base type when missing
BORDER_SUFFIX = " - Border"

# Symbols centered on their bounding box after placing, so a project's copy
# of the family with a different insertion point still lands where the
# layout put it (G1, 2026-11-18: an Equipment family inserted on its top
# line dropped onto its label). These are symmetric about the pipe; the
# squiggle is centered on its line end, the line running into its center
# (Colin, 2026-11-19). The PRV (dome to one side) and the cap (sits on the
# pipe end) keep the insertion.
CENTER_BY_BOX = ("equipment", "iso", "meter", "squiggle")

# fallback geometry, drafting view feet at 1:100
_EQ_HALF_W = 1.0
_EQ_PITCH = 0.56
_VALVE_HALF_L = 0.98
_VALVE_HALF_T = 0.59
_CAP_HALF_W = 0.84
_METER_R = 1.04


# ---------------------------------------------------------------------------
# pure helpers
# ---------------------------------------------------------------------------

def pick_text_type(types):
    """types: [(id, name, size_ft)]. The '3/32" Arial' type, else any 3/32
    in. type, else the first. Returns the id or None."""
    if not types:
        return None
    for tid, name, _size in types:
        if name == TEXT_TYPE_NAME:
            return tid
    for tid, _name, size in types:
        if size is not None and abs(size - TEXT_SIZE_FT) < 1e-6:
            return tid
    return types[0][0]


def text_type_note(types, picked):
    """Warning when the picked text type is not 3/32 in.: the layout's
    clash boxes assume 3/32 in. text at 1:100 (review n44). None when fine."""
    for tid, name, size in types or []:
        if tid != picked:
            continue
        if name == TEXT_TYPE_NAME or (size is not None and
                                      abs(size - TEXT_SIZE_FT) < 1e-6):
            return None
        return ("No 3/32 in. text type in this project; labels drawn with "
                "'{0}', so they may overlap symbols and other labels. Add "
                "the '{1}' text type and run One-Line again.".format(
                    name, TEXT_TYPE_NAME))
    if types:
        return "Text type for the labels could not be identified."
    return "No text type in this project; labels could not be drawn."


def style_notes(new_found, existing_found, failures):
    """Warnings for the line styles (review n12, n45). failures: {style
    name: count of lines whose LineStyle could not be set}."""
    out = []
    if not new_found:
        out.append("Line style '{0}' not found; pipes drawn thin.".format(
            NEW_LINE_STYLE))
    if not existing_found:
        out.append("Line style '{0}' not found; existing pipe and the "
                   "existing meter service drawn in the default line "
                   "style.".format(EXISTING_LINE_STYLE))
    for name in sorted(failures):
        out.append("Line style '{0}' could not be set on {1} line(s): {2}. "
                   "Those lines are drawn in the default style.".format(
                       name, failures[name][0], failures[name][1]))
    return out


def border_type_name(base_name):
    """Name of the bordered sibling of a text type."""
    base = base_name or TEXT_TYPE_NAME
    if base.endswith(BORDER_SUFFIX):
        return base
    return base + BORDER_SUFFIX


def center_shift(x, y, box):
    """(dx, dy) moving a placed symbol so its view box (minx, miny, maxx,
    maxy) is centered on (x, y)."""
    return (x - (box[0] + box[2]) / 2.0, y - (box[1] + box[3]) / 2.0)


def _rot(pts, x, y, deg):
    a = math.radians(deg)
    c, s = math.cos(a), math.sin(a)
    out = []
    for (px, py, qx, qy) in pts:
        out.append((x + px * c - py * s, y + px * s + py * c,
                    x + qx * c - qy * s, y + qx * s + qy * c))
    return out


def fallback_segments(prim):
    """Drawn stand in for a symbol whose family is missing, as straight
    segments (x0, y0, x1, y1). The meter fallback is a circle, drawn by
    draw() itself; it returns []."""
    role = prim["role"]
    x = prim.get("ix", prim["x"])
    y = prim.get("iy", prim["y"])
    rot = prim.get("rot") or 0
    if role == "equipment":
        segs = [(-_EQ_HALF_W, d, _EQ_HALF_W, d) for d in (-_EQ_PITCH, 0.0, _EQ_PITCH)]
    elif role in ("prv", "iso", "ansul"):
        L, T = _VALVE_HALF_L, _VALVE_HALF_T
        segs = [(-L, -T, -L, T), (L, -T, L, T), (-L, T, L, -T), (L, T, -L, -T)]
        if role == "prv":
            segs += [(0.0, 0.0, 0.0, T + 0.4), (-0.4, T + 0.4, 0.4, T + 0.4)]
    elif role == "cap":
        segs = [(-_CAP_HALF_W, 0.0, _CAP_HALF_W, 0.0)]
    elif role == "squiggle":
        segs = [(-0.35, -0.5, 0.35, 0.0), (0.35, 0.0, -0.35, 0.5)]
    else:
        return []
    return _rot(segs, x, y, rot)


def text_value(prim):
    return "\r".join(prim["lines"])


# ---------------------------------------------------------------------------
# Revit part
# ---------------------------------------------------------------------------

def _db():
    from Autodesk.Revit import DB
    return DB


def _name(el):
    DB = _db()
    try:
        return DB.Element.Name.GetValue(el)
    except Exception:
        return getattr(el, "Name", None)


def _find_family(doc, name):
    DB = _db()
    for f in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        if _name(f) == name:
            return f
    return None


def _first_symbol(doc, fam):
    if fam is None:
        return None
    for sid in fam.GetFamilySymbolIds():
        return doc.GetElement(sid)
    return None


FALLBACK_REASON = ("not loaded in this project, drawn as lines. Load the "
                   "family into the project to use the RJA symbol.")


def ensure_symbols(doc):
    """Find each symbol family in the project and activate its first type.
    Never loads a family file. Call inside a transaction.

    Returns ({role: FamilySymbol or None}, [report rows for
    gas_console.symbols_table])."""
    syms = {}
    report = []
    for role, names in SYMBOL_FAMILIES:
        sym = None
        row = {"role": ROLE_TITLES[role], "family": names[0] if names else "",
               "status": "fallback", "reason": FALLBACK_REASON}
        for name in names:
            fam = _find_family(doc, name)
            if fam is not None:
                sym = _first_symbol(doc, fam)
                if sym is not None:
                    row.update({"family": name, "status": "found",
                                "reason": ""})
                    break
        if sym is not None:
            try:
                if not sym.IsActive:
                    sym.Activate()
            except Exception as ex:
                row["reason"] = "activate failed: {0}".format(ex)
        syms[role] = sym
        report.append(row)
    return syms, report


def find_view_type(doc):
    DB = _db()
    for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
        if v.ViewFamily == DB.ViewFamily.Drafting:
            return v
    return None


def create_view(doc, name, scale=100):
    """A new drafting view. Call inside a transaction."""
    DB = _db()
    vft = find_view_type(doc)
    if vft is None:
        raise ValueError("No Drafting view type in this project.")
    view = DB.ViewDrafting.Create(doc, vft.Id)
    view.Name = name
    view.Scale = scale
    return view


def _text_types(doc):
    DB = _db()
    types = []
    for t in DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType):
        try:
            size = t.get_Parameter(DB.BuiltInParameter.TEXT_SIZE).AsDouble()
        except Exception:
            size = None
        types.append((t.Id, _name(t), size))
    return types


def text_type_id(doc):
    return pick_text_type(_text_types(doc))


def text_type_choice(doc):
    """(id or None, warning or None) for the label text type."""
    types = _text_types(doc)
    tid = pick_text_type(types)
    return tid, text_type_note(types, tid)


def border_text_type_id(doc, base_id):
    """The bordered sibling of the base text type ('3/32" Arial - Border'),
    duplicated from the base with Show Border on when the project has none.
    Call inside a transaction. Returns (id or None, note or None)."""
    DB = _db()
    if base_id is None:
        return None, "No text type for the notes border."
    base = doc.GetElement(base_id)
    want = border_type_name(_name(base))
    for t in DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType):
        if _name(t) == want:
            p = t.get_Parameter(DB.BuiltInParameter.TEXT_BOX_VISIBILITY)
            if p is not None and p.AsInteger() == 1:
                return t.Id, None
            return t.Id, "Text type '{0}' has Show Border off.".format(want)
    try:
        new = base.Duplicate(want)
        new.get_Parameter(DB.BuiltInParameter.TEXT_BOX_VISIBILITY).Set(1)
        return new.Id, "Text type '{0}' created for the notes.".format(want)
    except Exception as ex:
        return None, "Could not make text type '{0}': {1}".format(want, ex)


def line_style(doc, name=NEW_LINE_STYLE):
    """The PROJECTION GraphicsStyle of a line style. A line subcategory
    also has a Cut style of the same name, which DetailCurve.LineStyle
    rejects (review n12). UNVERIFIED Revit API: GraphicsStyle.GraphicsStyleType
    and GraphicsStyleType.Projection; when they cannot be read, the first
    style of that name is used, as before."""
    DB = _db()
    first = None
    for gs in DB.FilteredElementCollector(doc).OfClass(DB.GraphicsStyle):
        if _name(gs) != name:
            continue
        if first is None:
            first = gs
        try:
            if gs.GraphicsStyleType == DB.GraphicsStyleType.Projection:
                return gs
        except Exception:
            return first
    return first


class _Drawer(object):

    def __init__(self, doc, view, syms, tt_id, new_style, border_tt_id=None,
                 existing_style=None):
        self.doc = doc
        self.view = view
        self.syms = syms or {}
        self.tt_id = tt_id
        self.border_tt_id = border_tt_id
        self.new_style = new_style
        self.existing_style = existing_style
        self.counts = {"lines": 0, "texts": 0, "symbols": 0, "fallbacks": 0,
                       "failed": 0}
        self.errors = []
        # style name -> [lines it could not be set on, first error]
        self.style_failures = {}

    def _set_style(self, dc, style):
        """Set a detail curve's line style; a refusal is counted and
        reported, never passed over (Rule 6, review n12)."""
        if style is None:
            return
        try:
            dc.LineStyle = style
        except Exception as ex:
            try:
                key = _name(style) or "?"
            except Exception:
                key = "?"
            rec = self.style_failures.setdefault(key, [0, str(ex)])
            rec[0] += 1

    def line(self, x0, y0, x1, y1, new=False):
        DB = _db()
        if abs(x1 - x0) < 1e-3 and abs(y1 - y0) < 1e-3:
            return None
        dc = self.doc.Create.NewDetailCurve(
            self.view, DB.Line.CreateBound(DB.XYZ(x0, y0, 0), DB.XYZ(x1, y1, 0)))
        self._set_style(dc, self.new_style if new else self.existing_style)
        self.counts["lines"] += 1
        return dc

    def text(self, p):
        DB = _db()
        tid = self.tt_id
        if p.get("role") == "notes" and self.border_tt_id is not None:
            tid = self.border_tt_id
        opts = DB.TextNoteOptions(tid)
        opts.HorizontalAlignment = (DB.HorizontalTextAlignment.Center
                                    if p["align"] == "center"
                                    else DB.HorizontalTextAlignment.Left)
        opts.VerticalAlignment = DB.VerticalTextAlignment.Top
        tn = DB.TextNote.Create(self.doc, self.view.Id,
                                DB.XYZ(p["x"], p["y"], 0), text_value(p), opts)
        if p.get("underline_first") and p["lines"]:
            try:
                ft = tn.GetFormattedText()
                ft.SetUnderlineStatus(DB.TextRange(0, len(p["lines"][0])), True)
                tn.SetFormattedText(ft)
            except Exception as ex:
                self.errors.append("underline: {0}".format(ex))
        self.counts["texts"] += 1
        return tn

    def symbol(self, p):
        DB = _db()
        sym = self.syms.get(p["role"])
        x = p.get("ix", p["x"])
        y = p.get("iy", p["y"])
        if sym is not None:
            inst = self.doc.Create.NewFamilyInstance(DB.XYZ(x, y, 0), sym,
                                                     self.view)
            rot = p.get("rot") or 0
            if inst is not None and rot:
                axis = DB.Line.CreateBound(DB.XYZ(x, y, 0), DB.XYZ(x, y, 1))
                DB.ElementTransformUtils.RotateElement(
                    self.doc, inst.Id, axis, math.radians(rot))
            if inst is not None and p["role"] in CENTER_BY_BOX:
                self.doc.Regenerate()
                bb = inst.get_BoundingBox(self.view)
                if bb is not None:
                    dx, dy = center_shift(x, y, (bb.Min.X, bb.Min.Y,
                                                 bb.Max.X, bb.Max.Y))
                    if abs(dx) > 1e-3 or abs(dy) > 1e-3:
                        DB.ElementTransformUtils.MoveElement(
                            self.doc, inst.Id, DB.XYZ(dx, dy, 0))
            self.counts["symbols"] += 1
            return
        # fallback: drawn in the new pipe line style, Line 4 (Colin,
        # 2026-11-20), never the default <Lines> or the thin existing style
        if p["role"] == "meter":
            for a0 in (0.0, math.pi):
                arc = DB.Arc.Create(DB.XYZ(x, y, 0), _METER_R, a0, a0 + math.pi,
                                    DB.XYZ.BasisX, DB.XYZ.BasisY)
                dc = self.doc.Create.NewDetailCurve(self.view, arc)
                self._set_style(dc, self.new_style)
            opts = DB.TextNoteOptions(self.tt_id)
            opts.HorizontalAlignment = DB.HorizontalTextAlignment.Center
            opts.VerticalAlignment = DB.VerticalTextAlignment.Middle
            DB.TextNote.Create(self.doc, self.view.Id, DB.XYZ(x, y, 0), "M",
                               opts)
        else:
            for s in fallback_segments(p):
                self.line(s[0], s[1], s[2], s[3], new=True)
        self.counts["fallbacks"] += 1

    def draw(self, prims):
        for p in prims:
            try:
                if p["t"] == "line":
                    # new construction pipe and the new meter's utility stub
                    # get Line 4; existing phase pipe stays thin
                    self.line(p["x0"], p["y0"], p["x1"], p["y1"],
                              p.get("new", True))
                elif p["t"] == "text":
                    self.text(p)
                elif p["t"] == "symbol":
                    self.symbol(p)
            except Exception as ex:
                self.counts["failed"] += 1
                self.errors.append("{0} {1}: {2}".format(
                    p["t"], p.get("role") or p.get("seg") or "", ex))
        return self.counts


def draw(doc, view, prims, syms, tt_id=None, new_style=None):
    """Draw every primitive. Call inside a transaction. Returns (counts,
    errors): nothing is skipped silently."""
    tt_note = None
    if tt_id is None:
        tt_id, tt_note = text_type_choice(doc)
    if new_style is None:
        new_style = line_style(doc)
    border_id, border_note = border_text_type_id(doc, tt_id)
    existing_style = line_style(doc, EXISTING_LINE_STYLE)
    d = _Drawer(doc, view, syms, tt_id, new_style, border_id, existing_style)
    counts = d.draw(prims)
    if tt_note:
        d.errors.append(tt_note)
    if border_note and "created" not in border_note:
        d.errors.append(border_note)
    # only warn about Line 2 when something existing was drawn
    any_existing = any(p.get("t") == "line" and not p.get("new", True)
                       for p in prims)
    d.errors += style_notes(new_style is not None,
                            existing_style is not None or not any_existing,
                            d.style_failures)
    return counts, d.errors


# ---------------------------------------------------------------------------
# Developed Length Breakdown table (Colin, 2026-11-20)
# ---------------------------------------------------------------------------

BREAKDOWN_MARK = "DELETE BEFORE PUTTING ON SCHEDULE"
BREAKDOWN_GAP_BLOCKS = 2.0     # gaps (rja_drafting.GAP x view scale) under the diagram


def breakdown_title(bd):
    """Band text: LONGEST RUN TO <farthest outlet>."""
    return "LONGEST RUN TO {0}".format(bd.get("farthest") or "FARTHEST OUTLET")


def draw_breakdown(doc, view, bd, bounds, style=None):
    """The longest run table under the diagram, in the RJA table look.

    bd: oneline_schematic.developed_breakdown. bounds: the layout's
    (x0, y0, x1, y1) in model feet; the table starts at x0, a gap under y0.
    The view is 1:100, so rja_drafting's paper sized cells are scaled by the
    view scale (rja_drafting._Style(doc, scale)). The columns are fitted to
    their longest text first, then the blue band and the table are drawn at
    the same width. Call inside a transaction. Returns (bottom y, width).
    style: a stand in for rja_drafting._Style (tests).
    """
    import oneline_schematic
    import rja_drafting as rd
    k = float(getattr(view, "Scale", 1) or 1)
    st = style if style is not None else rd._Style(doc, k)
    headers, rows = oneline_schematic.breakdown_table(bd)
    needs, band_need = rd.column_needs(headers, rows, True, k)
    title_need = rd.text_width(breakdown_title(bd), True, k)
    needs = list(needs)
    short = max(band_need, title_need) - sum(needs, 0.0)
    if short > 0 and needs:
        needs[-1] += short       # a long outlet name widens the last column
    columns = [(h, w) for h, w in zip(headers, needs)]
    width = sum(needs, 0.0)
    ox = float(bounds[0])
    top = float(bounds[1]) - rd.GAP * k * BREAKDOWN_GAP_BLOCKS
    rd._t(st, view, ox, top, BREAKDOWN_MARK, "bold")
    top -= rd.TEXT_H * k * 1.4
    top = rd._band(st, view, ox, top, width, breakdown_title(bd), None,
                   rd.BAND_H * k)
    bottom = rd._styled_table(st, view, ox, top, columns, rows,
                              totals_row=True, left_cols=(0,), fit=True)
    return bottom, width
