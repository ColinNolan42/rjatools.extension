# -*- coding: utf-8 -*-
"""
console_report.py  --  RJA themed HTML blocks for the pyRevit output window.

Pure string builders, no Revit imports, so they run under IronPython 2.7 in
pyRevit and under CPython in lib/_test_console_report.py. Every function
returns an HTML string; the caller prints it with output.print_html().

The pyRevit output window is an embedded browser, so the markup is kept to
plain tables, divs and spans with inline styles: no flex or grid CSS, no
classes, no external fonts or images.

Theme: RJA blue #0070C0, Segoe UI, compact (lib/rja_theme.xaml uses the same
blue and font for the dialogs).

Status colors match the plan override colors used by Duct Performance
(script.py GREEN / YELLOW / RED / PURPLE / GRAY).

    >>> html = section('System Summary') + table(['A', 'B'], [['1', '2']])

IronPython 2.7: no f-strings, '{:.0f}'.format(int) raises, so every value is
passed in already formatted as text.
"""

RJA_BLUE = '#0070C0'
FONT = "'Segoe UI', Tahoma, Arial, sans-serif"
TEXT = '#1F1F1F'
MUTED = '#666666'
BORDER = '#C9D6E3'
HEADER_BG = '#DCE9F5'
ZEBRA_BG = '#F4F8FC'
TOTAL_BG = '#C5DBF0'

# Plan override colors from Duct Performance (Color(r, g, b) in script.py).
STATUS_COLORS = {
    'GREEN':  '#00C800',
    'YELLOW': '#FFD700',
    'RED':    '#D22828',
    'PURPLE': '#8C28C8',
    'GRAY':   '#A0A0A0',
}
# Text color that reads on each status background.
STATUS_TEXT = {
    'GREEN':  '#000000',
    'YELLOW': '#000000',
    'RED':    '#FFFFFF',
    'PURPLE': '#FFFFFF',
    'GRAY':   '#000000',
}
# Other status words that map onto a plan color.
STATUS_ALIASES = {
    'OK': 'GREEN', 'PASS': 'GREEN', 'WITHIN': 'GREEN',
    'LOW': 'RED', 'FAIL': 'RED', 'EXCEEDS': 'RED',
    'GREY': 'GRAY', 'N/A': 'GRAY', 'NOT': 'GRAY',
}

# (border, background, text) per callout kind.
CALLOUT_STYLES = {
    'warn':  ('#E0A800', '#FFF6D6', '#5C4300'),
    'error': ('#D22828', '#FDE3E3', '#7A1010'),
    'info':  (RJA_BLUE,  '#E3F0FB', '#00365E'),
    'ok':    ('#1E9E3A', '#E2F5E6', '#0F4D1D'),
}


def esc(value):
    """HTML escape any value as text. None becomes ''."""
    if value is None:
        return ''
    try:
        text = value if isinstance(value, basestring) else str(value)  # noqa: F821
    except NameError:   # CPython 3
        text = value if isinstance(value, str) else str(value)
    return (text.replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;').replace("'", '&#39;'))


def status_key(text):
    """The STATUS_COLORS key for a status cell, or None when it has none.

    Reads the first word, so 'LOW: under the minimum' and 'RED' both color.
    """
    if text is None:
        return None
    words = str(text).strip().upper().replace(':', ' ').replace(',', ' ').split()
    if not words:
        return None
    word = words[0]
    if word in STATUS_COLORS:
        return word
    return STATUS_ALIASES.get(word)


def _wrap(inner, margin='6px 0'):
    return ('<div style="font-family:{};font-size:12px;color:{};margin:{};">'
            '{}</div>').format(FONT, TEXT, margin, inner)


def section(title, subtitle=None):
    """Full width RJA blue title bar, white bold text, optional subtitle."""
    sub = ''
    if subtitle:
        sub = ('<span style="font-weight:normal;font-size:11px;'
               'margin-left:10px;color:#DCE9F5;">{}</span>').format(esc(subtitle))
    return ('<div style="font-family:{};background:{};color:#FFFFFF;'
            'font-weight:bold;font-size:14px;padding:5px 10px;'
            'margin:14px 0 6px 0;">{}{}</div>').format(
                FONT, RJA_BLUE, esc(title), sub)


def subsection(title):
    """Bold title with a blue left rule."""
    return ('<div style="font-family:{};border-left:4px solid {};'
            'padding:2px 8px;margin:10px 0 4px 0;font-weight:bold;'
            'font-size:13px;color:{};">{}</div>').format(
                FONT, RJA_BLUE, RJA_BLUE, esc(title))


def swatch(key):
    """A small filled square in a status color, followed by nothing."""
    color = STATUS_COLORS.get(key, '#FFFFFF')
    return ('<span style="display:inline-block;width:12px;height:12px;'
            'background:{};border:1px solid #555555;vertical-align:middle;'
            'margin-right:6px;"></span>').format(color)


def _cell_style(align, extra=''):
    return ('padding:3px 8px;border:1px solid {};text-align:{};'
            'vertical-align:top;{}').format(BORDER, align, extra)


def table(headers, rows, align=None, total_row=None, status_col=None,
          caption=None, swatch_col=None):
    """A compact bordered table.

    headers     list of header texts
    rows        list of row lists (text values, escaped here)
    align       list of 'l' / 'r' / 'c' per column, default left
    total_row   optional row list, drawn bold on a shaded band at the bottom
    status_col  optional column index whose cell is colored by status_key()
    caption     optional bold line above the table
    swatch_col  optional column index whose cell gets a color swatch, keyed by
                status_key() of the cell text
    """
    ncol = len(headers)
    align = list(align or [])
    align = [({'r': 'right', 'c': 'center'}.get(a, 'left')) for a in align]
    align += ['left'] * (ncol - len(align))

    out = []
    if caption:
        out.append('<div style="font-weight:bold;margin:6px 0 2px 0;">{}</div>'
                   .format(esc(caption)))
    out.append('<table style="border-collapse:collapse;font-family:{};'
               'font-size:12px;margin:2px 0 6px 0;">'.format(FONT))
    out.append('<tr>')
    for i, h in enumerate(headers):
        out.append('<th style="{}">{}</th>'.format(
            _cell_style(align[i], 'background:{};font-weight:bold;color:{};'
                        'white-space:nowrap;'.format(HEADER_BG, TEXT)), esc(h)))
    out.append('</tr>')

    for r, row in enumerate(rows):
        bg = ZEBRA_BG if r % 2 else '#FFFFFF'
        out.append('<tr>')
        cells = list(row) + [''] * (ncol - len(row))
        for i, cell in enumerate(cells[:ncol]):
            extra = 'background:{};'.format(bg)
            body = esc(cell)
            if status_col is not None and i == status_col:
                key = status_key(cell)
                if key is not None:
                    extra = 'background:{};color:{};font-weight:bold;'.format(
                        STATUS_COLORS[key], STATUS_TEXT[key])
            if swatch_col is not None and i == swatch_col:
                key = status_key(cell)
                if key is not None:
                    body = swatch(key) + body
            out.append('<td style="{}">{}</td>'.format(_cell_style(align[i], extra), body))
        out.append('</tr>')

    if total_row is not None:
        cells = list(total_row) + [''] * (ncol - len(total_row))
        out.append('<tr>')
        for i, cell in enumerate(cells[:ncol]):
            out.append('<td style="{}">{}</td>'.format(
                _cell_style(align[i], 'background:{};font-weight:bold;'
                            'border-top:2px solid {};'.format(TOTAL_BG, RJA_BLUE)),
                esc(cell)))
        out.append('</tr>')
    out.append('</table>')
    return ''.join(out)


def callout(text, kind='info'):
    """A highlighted box. kind: 'warn' (amber), 'error' (red), 'info' (blue),
    'ok' (green). An unknown kind falls back to 'info'."""
    border, bg, color = CALLOUT_STYLES.get(kind, CALLOUT_STYLES['info'])
    return ('<div style="font-family:{};border-left:5px solid {};background:{};'
            'color:{};padding:5px 10px;margin:4px 0;font-size:12px;">{}</div>'
            ).format(FONT, border, bg, color, esc(text))


def keyvalue(rows):
    """Two column label / value grid. rows: [(label, value), ...]."""
    out = ['<table style="border-collapse:collapse;font-family:{};'
           'font-size:12px;margin:2px 0 6px 0;">'.format(FONT)]
    for label, value in rows:
        out.append('<tr><td style="padding:2px 12px 2px 0;color:{};'
                   'white-space:nowrap;">{}</td>'
                   '<td style="padding:2px 0;font-weight:bold;">{}</td></tr>'
                   .format(MUTED, esc(label), esc(value)))
    out.append('</table>')
    return ''.join(out)


def note(text):
    """Small muted paragraph."""
    return ('<div style="font-family:{};font-size:11px;color:{};'
            'margin:3px 0;">{}</div>').format(FONT, MUTED, esc(text))


def paragraph(text, bold_lead=None):
    """Normal text, with an optional bold lead in sentence."""
    lead = ''
    if bold_lead:
        lead = '<b>{}</b> '.format(esc(bold_lead))
    return _wrap(lead + esc(text), margin='4px 0')


def markdown_table(headers, rows):
    """The same rows as a markdown table, for the print_md fallback."""
    def _md(v):
        return esc(v).replace('|', '/')
    lines = ['| ' + ' | '.join(_md(h) for h in headers) + ' |',
             '| ' + ' | '.join('---' for _h in headers) + ' |']
    for row in rows:
        lines.append('| ' + ' | '.join(_md(c) for c in row) + ' |')
    return '\n'.join(lines)
