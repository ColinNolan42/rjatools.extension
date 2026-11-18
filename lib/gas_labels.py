# -*- coding: ascii -*-
"""
gas_labels.py

The one home for gas One-Line label text (rebuild plan step 1, 2026-11-18).

    pipe_label("3", 25.2)              -> '3"G, 25\''
    pipe_label("1-1/4 K&L", 6.4)       -> '1-1/4"G, 6\''   (material stripped)
    mbh_text(1615)                     -> '1615 MBH'
    fixture_label("RTU-1", 64.8)       -> ['RTU-1', '64.8 MBH']
    cap_label(None, 0)                 -> ['FUTURE CONNECTION']
    notes_lines("402.4(2)", 1814, 117) -> notes block lines

Rules (Colin):
- Pipe label: size, inch mark, G, comma, length in whole feet with an
  apostrophe. No space before G, no material (2026-11-18, item 5 and F4).
- Copper "K&L" and any other material word after the size is stripped; the
  material is stated once in the notes block instead.
- Line 2 is the cumulative MBH.
- An unnamed cap at 0 MBH reads FUTURE CONNECTION, no MBH line.
- The notes wording lives only in NOTES_TEMPLATE (firm wording pending).

Pure Python, no Revit import. IronPython 2.7 and CPython 3. ASCII source.
"""

import math

import gas_tables


# ---------------------------------------------------------------------------
# Notes block wording. Firm wording pending from Colin: edit here only.
# Fields: inlet_text, drop_text, table_id, material_text, total_mbh,
# developed_ft. The PRV line is added only when a mid stream PRV exists.
# ---------------------------------------------------------------------------

NOTES_TEMPLATE = [
    "CONTRACTOR SHALL SUBMIT APPLICATIONS TO UTILITY AND COORDINATE NEW METER SERVICE",
    "GAS PIPING SIZED FOR {inlet_text}",
    "MAX PRESSURE LOSS OF {drop_text} PER IFGC TABLE {table_id}",
    "PIPING MATERIAL: {material_text}",
    "TOTAL CONNECTED LOAD: {total_mbh} MBH",
    "TOTAL DEVELOPED LENGTH: {developed_ft}'",
]

NOTES_PRV_LINE = ("DOWNSTREAM OF PRV: {inlet_text}, MAX PRESSURE LOSS OF "
                  "{drop_text} PER IFGC TABLE {table_id}")

FUTURE_CONNECTION = "FUTURE CONNECTION"
UNNAMED = "UNNAMED"

# Words that may follow a nominal size in the table names or the model.
_MATERIAL_WORDS = ("K&L", "ACR", "CTS", "IPS", "SDR")


def _half_up(x):
    """Round half up to an int. Python 3 round() is banker's rounding."""
    return int(math.floor(float(x) + 0.5))


def feet(length_ft):
    """Whole feet for a label. Any real length under 0.5 ft shows as 1."""
    if length_ft is None:
        return None
    n = _half_up(length_ft)
    if n < 1 and float(length_ft) > 0.0:
        n = 1
    return n


def size_text(nominal):
    """Nominal size as printed, material removed. None when blank.

    '1-1/4 K&L' -> '1-1/4'; '2-1/2"' -> '2-1/2'; 'EHD-13' -> 'EHD-13'.
    """
    if nominal is None:
        return None
    s = str(nominal).strip()
    if not s:
        return None
    s = s.replace('"', "").strip()
    parts = s.split()
    keep = []
    for p in parts:
        if p.upper() in _MATERIAL_WORDS:
            continue
        keep.append(p)
    s = " ".join(keep).strip()
    return s or None


def _is_inch_size(s):
    """True for '3', '1/2', '1-1/4'; False for CSST 'EHD-13'."""
    return bool(s) and s[0].isdigit()


def pipe_label(nominal, length_ft):
    """Line 1 of a pipe label: 3"G, 25'.

    CSST sizes keep their EHD name with G: EHD-13G, 25'. No size gives the
    length only (the One-Line gate rejects unsized pipe before drawing).
    """
    s = size_text(nominal)
    ft = feet(length_ft)
    length = "{0}'".format(ft) if ft is not None else ""
    if s is None:
        return length
    if _is_inch_size(s):
        size = '{0}"G'.format(s)
    else:
        size = "{0}G".format(s)
    if not length:
        return size
    return "{0}, {1}".format(size, length)


def mbh_number(mbh):
    """MBH as printed: whole when within 0.05 of a whole number, else 1 place."""
    v = float(mbh or 0.0)
    r = _half_up(v)
    if abs(v - r) < 0.05:
        return "{0}".format(r)
    return "{0:.1f}".format(v)


def mbh_text(mbh):
    """Line 2 of a pipe or fixture label: 1615 MBH."""
    return "{0} MBH".format(mbh_number(mbh))


def segment_label(nominal, length_ft, mbh):
    """Both lines of a pipe label."""
    return [pipe_label(nominal, length_ft), mbh_text(mbh)]


def fixture_label(name, mbh):
    """Fixture label lines: name (UNNAMED when blank), then MBH."""
    n = (name or "").strip() or UNNAMED
    return [n, mbh_text(mbh)]


def is_future_name(name):
    """Blank, UNNAMED or starting with FUTURE (One-Line cap symbol)."""
    n = (name or "").strip().upper()
    return (not n) or n == UNNAMED or n.startswith("FUTURE")


def cap_label(name, mbh):
    """Label for a cap drawn with the cap symbol.

    Unnamed or FUTURE at 0 MBH: FUTURE CONNECTION, no MBH line (Colin,
    2026-11-18). FUTURE with a load: FUTURE and its MBH. A named, non FUTURE
    cap at 0 MBH is a data error the gate rejects; it is labelled by name.
    """
    load = float(mbh or 0.0)
    n = (name or "").strip()
    if is_future_name(n):
        if load <= 0.0:
            return [FUTURE_CONNECTION]
        if not n or n.upper() == UNNAMED:
            n = "FUTURE"
        return [n, mbh_text(load)]
    return fixture_label(n, load)


# ---------------------------------------------------------------------------
# Notes block
# ---------------------------------------------------------------------------

def _table_texts(table_id):
    t = gas_tables.get_table(table_id)
    inlet = str(t.get("inlet_pressure") or "").upper()
    drop = str(t.get("pressure_drop") or "").upper()
    material = str(t.get("description") or "").upper()
    return inlet, drop, material


def notes_lines(table_id, total_mbh, developed_ft, downstream_table_id=None,
                template=None):
    """Notes block lines from NOTES_TEMPLATE and the table's own text.

    Inlet pressure and drop come from the table JSON (402.4(2): LESS THAN
    2 PSI, 0.5 IN. W.C.), never from the TABLE_OPTIONS 0.7 placeholder.
    """
    inlet, drop, material = _table_texts(table_id)
    fields = {
        "inlet_text": inlet,
        "drop_text": drop,
        "table_id": table_id,
        "material_text": material,
        "total_mbh": mbh_number(total_mbh),
        "developed_ft": feet(developed_ft) if developed_ft else 0,
    }
    lines = [ln.format(**fields) for ln in (template or NOTES_TEMPLATE)]
    if downstream_table_id:
        d_inlet, d_drop, _m = _table_texts(downstream_table_id)
        lines.append(NOTES_PRV_LINE.format(
            inlet_text=d_inlet, drop_text=d_drop,
            table_id=downstream_table_id))
    return lines
