# Parse the 2024 IPC tables out of the PDFs and the firm sizing table out of the
# workbook, cross-check them, and emit lib/ipc_water_tables.json.
# Values are PARSED, never typed. Run with `python` (CPython 3.x).
import fitz, json, re, io, sys, datetime

PDFS = r"C:\Users\Colin Nolan\Desktop\FOR CLAUDE\Skills\2024 I Codes\pdfs"
CALC = r"C:\Users\Colin Nolan\Desktop\FOR CLAUDE\Calc Tools\Plumbing"
IPC3 = PDFS + r"\2024_IPC_3.pdf"
IPC1 = PDFS + r"\2024_IPC_1.pdf"

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
log = io.StringIO()
def say(m):
    log.write(m + "\n")
    print(m)

DASH = "\u2014"          # em dash used for "no value" in the printed tables
PRIME = "\u2033"         # double prime used for inch marks

# The occupancy cell of Table E103.3(2) only ever holds one of these.
OCCUPANCIES = ("Private", "Public", "Offices, etc.", "Hotel, restaurant",
               "Public or private")


def page_lines(path, pno):
    doc = fitz.open(path)
    txt = doc[pno].get_text()
    doc.close()
    return [ln.strip() for ln in txt.split("\n")]


def is_val(tok):
    """A table value cell: a number (maybe with commas) or the em dash."""
    if tok == DASH or tok == "-":
        return True
    return bool(re.match(r"^\d{1,3}(,\d{3})*(\.\d+)?$|^\d+(\.\d+)?$", tok))


def num(tok):
    if tok == DASH or tok == "-":
        return None
    return float(tok.replace(",", ""))


# ---------------------------------------------------------------- E103.3(2)
def parse_e103_3_2():
    """Fixture -> cold/hot/total wsfu. Records are label lines then 3 values."""
    raw = page_lines(IPC3, 35) + page_lines(IPC3, 36)
    # keep only the table body: stop at the SI/footnote lines
    rows, label, vals = [], [], []
    started = False
    for ln in raw:
        if ln.startswith("TABLE E103.3(2)"):
            started = True
            label, vals = [], []
            continue
        if not started or not ln:
            continue
        if ln.startswith("For SI:") or ln.startswith("a. For fixtures"):
            started = False
            continue
        if ln in ("FIXTURE", "OCCUPANCY", "TYPE OF SUPPLY", "CONTROL",
                  "LOAD VALUES, IN WATER SUPPLY FIXTURE UNITS (wsfu)",
                  "Cold", "Hot", "Total"):
            continue
        if is_val(ln):
            vals.append(ln)
            if len(vals) == 3:
                # Anchor on the occupancy cell, counting from the END of the
                # accumulated label lines. Page headers pile up in front of a
                # continued row, and the supply-control cell can wrap onto a
                # second line, so neither end has a fixed offset.
                occ_at = None
                for j in range(len(label) - 1, 0, -1):
                    if label[j] in OCCUPANCIES:
                        occ_at = j
                        break
                if occ_at is not None:
                    rows.append({
                        "fixture": label[occ_at - 1],
                        "occupancy": label[occ_at],
                        "supply_control": " ".join(label[occ_at + 1:]),
                        "cold_wsfu": num(vals[0]),
                        "hot_wsfu": num(vals[1]),
                        "total_wsfu": num(vals[2]),
                    })
                label, vals = [], []
        else:
            if vals:            # values seen but fewer than 3 -> not a real row
                label, vals = [], []
            label.append(ln)
    return rows


# ---------------------------------------------------------------- E103.3(3)
def parse_e103_3_3():
    """Demand curve. Each printed row is 6 cells: tank load/gpm/cfm, valve load/gpm/cfm."""
    # Each page needs its own window. On the first page the data follows the
    # caption; on the continued page the data comes FIRST and the repeated
    # caption sits after it, followed by the friction-chart axis scale, which
    # is numeric and would otherwise be swallowed as table rows.
    def window(pno, start_after_caption):
        toks, live = [], not start_after_caption
        for ln in page_lines(IPC3, pno):
            if start_after_caption and ln.startswith("TABLE E103.3(3)"):
                live = True
                continue
            if ln.startswith("Copyright") or ln.startswith("For SI:"):
                break
            if live and ln and is_val(ln):
                toks.append(ln)
        return toks

    toks = window(37, True) + window(38, False)
    if len(toks) % 6:
        raise SystemExit("E103.3(3): {} tokens is not a whole number of "
                         "6-cell rows".format(len(toks)))
    tank, valve = [], []
    for i in range(0, len(toks), 6):
        tl, tg, tc, vl, vg, vc = (num(x) for x in toks[i:i + 6])
        tank.append({"wsfu": tl, "gpm": tg, "cfm": tc})
        if vl is not None:
            valve.append({"wsfu": vl, "gpm": vg, "cfm": vc})
    # monotonic load is the sanity check that the grouping landed correctly
    for seq, name in ((tank, "flush tank"), (valve, "flushometer")):
        for a, b in zip(seq, seq[1:]):
            if b["wsfu"] <= a["wsfu"]:
                raise SystemExit("E103.3(3) {}: load not increasing at {} -> {}"
                                 .format(name, a, b))
    return tank, valve


# ---------------------------------------------------------------- 604.5
def parse_604_5():
    raw = page_lines(IPC1, 262)
    rows, started, pending = [], False, None
    for ln in raw:
        if ln.startswith("TABLE 604.5"):
            started = True
            continue
        if not started or not ln:
            continue
        if ln.startswith("For SI:") or ln.startswith("a. Where"):
            break
        if ln in ("FIXTURE", "MINIMUM PIPE SIZE (inch)"):
            continue
        if re.match(r"^\d+(/\d+)?$", ln) and pending:
            rows.append({"fixture": pending, "min_supply_in": ln})
            pending = None
        else:
            pending = ln
    return rows


# ---------------------------------------------------------------- 604.3
def parse_604_3():
    raw = page_lines(IPC1, 259) + page_lines(IPC1, 260)
    rows, label, vals = [], None, []
    for ln in raw:
        if not ln or ln.startswith("Copyright") or ln.startswith("For SI"):
            continue
        if ln in ("FIXTURE SUPPLY OUTLET SERVING", "FLOW RATEa", "(gpm)",
                  "FLOW", "PRESSURE", "(psi)"):
            continue
        clean = re.sub(r"[a-z]$", "", ln)      # strip footnote letters like 2.5b
        if is_val(clean):
            vals.append(clean)
            if len(vals) == 2 and label:
                rows.append({"fixture": label,
                             "flow_gpm": num(vals[0]),
                             "flow_pressure_psi": num(vals[1])})
                label, vals = None, []
        else:
            if len(ln) > 3 and not ln[0].isdigit():
                label, vals = ln, []
    return rows


# ---------------------------------------------------------------- firm table
def parse_firm_table():
    import openpyxl
    p = CALC + r"\Domestic Water\Template-Domestic Water Pipe Sizing Calculator.xlsx"
    wb = openpyxl.load_workbook(p, data_only=False)
    ws = wb["Sheet1"]
    sizes, hot, cold = [], [], []
    for r in range(6, 14):                      # G6:I13
        sizes.append(str(ws.cell(row=r, column=7).value))
        hot.append(ws.cell(row=r, column=8).value)
        cold.append(ws.cell(row=r, column=9).value)
    # Upper bounds come from the DHW/DCW IF() formulas in C6/D6, which are the
    # authoritative thing the workbook actually evaluates.
    f_hot = ws["C6"].value
    f_cold = ws["D6"].value
    wb.close()

    def bounds(formula):
        # Capture each '<=LIMIT , SIZE' pair in order, so both the WSFU limit
        # and the nominal size the workbook actually returns come from the
        # spreadsheet rather than being typed here. The first pair has no
        # closing paren before the comma.
        pairs = re.findall(r"<=\s*([\d.]+)\s*\)?\s*,\s*([\d.]+)", formula)
        return [float(a) for a, b in pairs], [float(b) for a, b in pairs]

    hb, h_sizes = bounds(f_hot)
    cb, c_sizes = bounds(f_cold)
    if h_sizes != c_sizes:
        raise SystemExit("hot and cold size ladders differ: {} vs {}".format(
            h_sizes, c_sizes))
    say("  firm size ladder (decimal in): {}".format(h_sizes))
    say("  firm DHW 5fps formula bounds: {}".format(hb))
    say("  firm DCW 8fps formula bounds: {}".format(cb))
    say("  firm printed ranges hot : {}".format(hot))
    say("  firm printed ranges cold: {}".format(cold))
    rows = []
    for i, s in enumerate(sizes):
        rows.append({"nominal_size": s,
                     "size_inches": h_sizes[i],
                     "max_wsfu_5fps_hot": hb[i],
                     "max_wsfu_8fps_cold": cb[i]})
    return rows, f_hot, f_cold


def firm_demand_rows():
    import openpyxl
    p = CALC + r"\Domestic Water\Template-Domestic Water Pipe Sizing Calculator.xlsx"
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb["Sheet1"]
    tank, valve = [], []
    r = 57
    while r <= 112:
        b, c = ws.cell(row=r, column=2).value, ws.cell(row=r, column=3).value
        e, f = ws.cell(row=r, column=5).value, ws.cell(row=r, column=6).value
        if isinstance(b, (int, float)) and isinstance(c, (int, float)):
            tank.append({"wsfu": float(b), "gpm": float(c)})
        if isinstance(e, (int, float)) and isinstance(f, (int, float)):
            valve.append({"wsfu": float(e), "gpm": float(f)})
        r += 1
    wb.close()
    return tank, valve


# ================================================================== run
say("=== 2024 IPC water tables, parsed {} ===".format(datetime.date.today()))

wsfu = parse_e103_3_2()
say("\nTable E103.3(2): {} fixture rows".format(len(wsfu)))
for r in wsfu:
    say("  {:<24} {:<18} {:<24} {!s:>6} {!s:>6} {!s:>6}".format(
        r["fixture"][:24], r["occupancy"][:18], r["supply_control"][:24],
        r["cold_wsfu"], r["hot_wsfu"], r["total_wsfu"]))

tank, valve = parse_e103_3_3()
say("\nTable E103.3(3): {} flush-tank rows, {} flushometer rows".format(
    len(tank), len(valve)))
say("  tank  first={} last={}".format(tank[0], tank[-1]))
say("  valve first={} last={}".format(valve[0], valve[-1]))

t604_5 = parse_604_5()
say("\nTable 604.5: {} rows".format(len(t604_5)))
for r in t604_5:
    say("  {:<50} {}".format(r["fixture"][:50], r["min_supply_in"]))

t604_3 = parse_604_3()
say("\nTable 604.3: {} rows".format(len(t604_3)))
for r in t604_3:
    say("  {:<70} {:>6} gpm @ {:>5} psi".format(
        r["fixture"][:70], r["flow_gpm"], r["flow_pressure_psi"]))

say("\n=== firm sizing table (Template-Domestic Water Pipe Sizing Calculator.xlsx) ===")
firm, f_hot, f_cold = parse_firm_table()
for r in firm:
    say("  {:<8} hot<=5fps {:>6}   cold<=8fps {:>6}".format(
        r["nominal_size"], r["max_wsfu_5fps_hot"], r["max_wsfu_8fps_cold"]))

# ---- cross-check the workbook's demand copy against the IPC PDF ----
say("\n=== DIFF: workbook demand curve vs 2024 IPC Table E103.3(3) ===")
wtank, wvalve = firm_demand_rows()
say("  workbook rows: {} tank, {} valve   PDF rows: {} tank, {} valve".format(
    len(wtank), len(wvalve), len(tank), len(valve)))
diffs = 0
for label, a, b in (("tank", tank, wtank), ("valve", valve, wvalve)):
    n = min(len(a), len(b))
    for i in range(n):
        if abs(a[i]["wsfu"] - b[i]["wsfu"]) > 1e-9 or abs(a[i]["gpm"] - b[i]["gpm"]) > 1e-9:
            diffs += 1
            say("  MISMATCH {} row {}: PDF {} vs workbook {}".format(
                label, i, a[i], b[i]))
    if len(a) != len(b):
        diffs += 1
        say("  ROW COUNT differs for {}: PDF {} workbook {}".format(label, len(a), len(b)))
say("  differences found: {}".format(diffs))

# ---- emit JSON ----
payload = {
    "_source": {
        "generated": str(datetime.date.today()),
        "code": "2024 International Plumbing Code",
        "wsfu_table": "Table E103.3(2), 2024_IPC_3.pdf PDF pages 36-37",
        "demand_table": "Table E103.3(3), 2024_IPC_3.pdf PDF pages 38-39",
        "fixture_flow_table": "Table 604.3, 2024_IPC_1.pdf PDF pages 260-261",
        "fixture_supply_table": "Table 604.5, 2024_IPC_1.pdf PDF page 263",
        "firm_sizing_table": ("Template-Domestic Water Pipe Sizing Calculator.xlsx, "
                              "Sheet1, flush tank, Copper Type L"),
        "firm_formula_hot_5fps": f_hot,
        "firm_formula_cold_8fps": f_cold,
        "note_cfm_typos": ("The printed 2024 IPC cubic-feet-per-minute column has "
                           "typographical errors (1 wsfu prints 0.04104, 2 prints "
                           "0.0684, 16 prints 2.90624). The engine uses gpm only."),
        "note_velocity_basis": ("Firm workbook cites IPC commentary figure 604.1: up to "
                                "5 fps for hot water to 140F and up to 8 fps for cold "
                                "water in copper. The note under IPC Figures E103.3(2) "
                                "and E103.3(3) says velocities above 5 to 8 ft/s are "
                                "not usually recommended."),
    },
    "wsfu_by_fixture": wsfu,
    "demand_flush_tank": tank,
    "demand_flushometer_valve": valve,
    "fixture_flow_604_3": t604_3,
    "fixture_supply_604_5": t604_5,
    "firm_sizing_table": firm,
}
# ensure_ascii so the file is pure ASCII with \u escapes: IronPython 2.7 reads
# it through a plain open() + json.load, exactly like ifgc_gas_sizing_tables.json.
json.dump(payload, open("ipc_water_tables.json", "w", encoding="ascii"),
          indent=2, ensure_ascii=True)
say("\nwrote ipc_water_tables.json")
open("build_report.txt", "w", encoding="utf-8").write(log.getvalue())
