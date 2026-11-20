# -*- coding: ascii -*-
"""Tests for the Size Gas headless entry (gas_run.run_headless, its settings,
the start point checks and the result envelope).

Plain CPython with the Revit assemblies stubbed. The traversal and the IFGC
sizing are covered by _test_gas_traversal, _test_prv_zones, _test_gas_tie_in
and the live bench runs (G1, G4, G6, G6C, G7 on 2026-11-16, Revit 2027).

Run:  python _test_gas_run_headless.py
"""

import json
import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _install_stubs():
    autodesk = types.ModuleType("Autodesk")
    revit = types.ModuleType("Autodesk.Revit")
    db = types.ModuleType("Autodesk.Revit.DB")
    ui = types.ModuleType("Autodesk.Revit.UI")
    selection = types.ModuleType("Autodesk.Revit.UI.Selection")
    for name in ("BuiltInParameter", "ElementId", "FilteredElementCollector",
                 "BuiltInCategory", "FlowDirectionType", "MEPSystem",
                 "SpecTypeId", "Transaction", "Domain", "XYZ",
                 "PartType", "UnitTypeId", "UnitUtils", "MEPCurve"):
        setattr(db, name, type(name, (object,), {}))
    selection.ObjectType = type("ObjectType", (object,), {"Element": 1})
    autodesk.Revit = revit
    revit.DB = db
    revit.UI = ui
    ui.Selection = selection
    sys.modules.update({
        "Autodesk": autodesk, "Autodesk.Revit": revit,
        "Autodesk.Revit.DB": db, "Autodesk.Revit.UI": ui,
        "Autodesk.Revit.UI.Selection": selection,
        "ui_helpers": types.ModuleType("ui_helpers"),
    })
    pyrevit = types.ModuleType("pyrevit")
    pyrevit.HOST_APP = type("HostApp", (object,), {"version": "2024"})()
    sys.modules["pyrevit"] = pyrevit


_install_stubs()

import headless        # noqa: E402
import shared_params   # noqa: E402
import gas_run         # noqa: E402

FAILS = []
COUNT = [0]


def check(name, cond):
    COUNT[0] += 1
    if not cond:
        FAILS.append(name)
        print("FAIL " + name)


# --- settings ----------------------------------------------------------------
d = gas_run.default_settings()
check("default keys", sorted(d) == ["downstream_table_id", "existing",
                                    "heat_content_btu_per_cf",
                                    "pipe_material",
                                    "propane_heat_content_btu_per_cf",
                                    "table_id"])
check("default propane heat", d["propane_heat_content_btu_per_cf"]
      == float(shared_params.DEFAULT_PROPANE_HEAT_CONTENT_BTU_PER_CF))
check("no default table", d["table_id"] is None)
check("default heat", d["heat_content_btu_per_cf"]
      == float(shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF))
check("no table refused", any("table_id is required" in p
                              for p in gas_run.validate_settings(d)))
ok = gas_run.merge_settings({"table_id": "402.4(10)"})
check("copper 1.0 valid", gas_run.validate_settings(ok) == [])
try:
    gas_run.merge_settings({"table": "402.4(10)"})
    check("unknown key raises", False)
except ValueError as exc:
    check("unknown key raises", "table" in str(exc))
check("unknown table", "not an offered" in gas_run.validate_settings(
    dict(ok, table_id="402.4(99)"))[0])
check("material mismatch", "for Schedule 40 Steel" in gas_run.validate_settings(
    dict(ok, pipe_material="Schedule 40 Steel"))[0])
# gas_tables.read_heat_content: a number from 500 to 3,000 BTU/CF (n=34/39)
check("heat zero", gas_run.validate_settings(
    dict(ok, heat_content_btu_per_cf=0)) == [
        "heat_content_btu_per_cf shall be between 500 and 3000 BTU/CF."])
check("heat slipped digit", gas_run.validate_settings(
    dict(ok, heat_content_btu_per_cf=8400)) == [
        "heat_content_btu_per_cf shall be between 500 and 3000 BTU/CF."])
check("heat text", gas_run.validate_settings(
    dict(ok, heat_content_btu_per_cf="x")) == [
        "heat_content_btu_per_cf shall be a number."])
check("heat nan", gas_run.validate_settings(
    dict(ok, heat_content_btu_per_cf=float("nan"))) == [
        "heat_content_btu_per_cf shall be a number."])
check("propane heat checked", gas_run.validate_settings(
    dict(ok, propane_heat_content_btu_per_cf=25000)) == [
        "propane_heat_content_btu_per_cf shall be between 500 and 3000 "
        "BTU/CF."])
check("propane heat valid", gas_run.validate_settings(
    dict(ok, propane_heat_content_btu_per_cf=2516)) == [])

# Mid stream PRV: same material and gas, never a higher pressure (S11).
zoned = gas_run.merge_settings({"table_id": "402.4(12)",
                                "downstream_table_id": "402.4(10)"})
check("MP to LP valid", gas_run.validate_settings(zoned) == [])
s11 = gas_run.validate_settings(dict(zoned, table_id="402.4(10)",
                                     downstream_table_id="402.4(12)"))
check("S11 higher pressure", len(s11) == 1 and "higher pressure" in s11[0])
steel = gas_run.validate_settings(dict(zoned, downstream_table_id="402.4(2)"))
check("downstream other material", len(steel) == 1
      and "shall be a" in steel[0] and "Natural table" in steel[0])

# Existing conditions for a pipe start.
ex = {"size_nominal": "1-1/2 K&L", "length_ft": 60, "load_mbh": 270}
check("existing valid", gas_run.validate_settings(
    dict(ok, table_id="402.4(9)", existing=ex)) == [])
check("existing missing key", gas_run.validate_settings(
    dict(ok, existing={"size_nominal": "1 K&L"})) == [
        "existing is missing length_ft, load_mbh"])
bad_size = gas_run.validate_settings(
    dict(ok, existing=dict(ex, size_nominal="1-1/2")))
check("existing size not in table", len(bad_size) == 1
      and "not in Table 402.4(10)" in bad_size[0])
check("existing negative", gas_run.validate_settings(
    dict(ok, existing=dict(ex, length_ft=-1))) == [
        "existing length_ft shall be 0 or more"])
check("existing not a dict", gas_run.validate_settings(
    dict(ok, existing=5)) == [
        "existing shall be {size_nominal, length_ft, load_mbh}"])
check("existing new construction", gas_run.validate_settings(
    dict(ok, existing={"new_construction": True})) == [])

opts = gas_run.table_options()
check("table options", len(opts) == len(gas_run.gas_tables.TABLE_OPTIONS)
      and set(opts[0]) == set(["table_id", "material", "gas",
                               "inlet_pressure_psi", "label"]))

# Dialog answer to settings (the button path).
choice = {"pipe_material": opts[0]["material"],
          "table_label": gas_run.gas_tables._short_label(
              gas_run.gas_tables.TABLE_OPTIONS[0]),
          "heat_content": 1000, "mid_stream_prv": False,
          "downstream_table_label": None}
sd = gas_run.settings_from_dialog(choice)
check("dialog to settings", sd["table_id"] == opts[0]["table_id"]
      and sd["heat_content_btu_per_cf"] == 1000.0
      and sd["downstream_table_id"] is None
      and sd["existing"] is None
      and sorted(sd) == sorted(gas_run.default_settings()))
check("dialog without propane box gives the default",
      sd["propane_heat_content_btu_per_cf"]
      == shared_params.DEFAULT_PROPANE_HEAT_CONTENT_BTU_PER_CF)
scope = {"size_nominal": "1-1/2", "length_ft": 40.0, "load_mbh": 200.0}
sd = gas_run.settings_from_dialog(dict(
    choice, propane_heat_content_btu_per_cf=2400, existing=scope))
check("dialog propane and scope carried",
      sd["propane_heat_content_btu_per_cf"] == 2400.0
      and sd["existing"] == scope)

# --- refusals before any Revit call -------------------------------------------
r = gas_run.run_headless(None, 1, settings={})
check("refused no table", r["status"] == "refused"
      and "table_id is required" in r["errors"][0])
r = gas_run.run_headless(None, 1, settings={"table_id": "402.4(10)"},
                         detail="some")
check("refused detail", r["status"] == "refused")
r = gas_run.run_headless(None, 1, settings={"typo": 1})
check("refused unknown", r["status"] == "refused" and "typo" in r["errors"][0])
r = gas_run.run_headless(None, 1, settings={"table_id": "402.4(10)"},
                         decisions={"presizing_warnings": "yes"})
check("refused bad decision", r["status"] == "refused")


# --- resolve_start --------------------------------------------------------------
class Elem(object):
    def __init__(self, n, pipe=False):
        self.Id = n
        self.pipe = pipe


gas_run.revit_helpers.is_pipe = lambda e: e.pipe
gas_run.revit_helpers.eid_int = lambda i: int(i)
gas_run.revit_helpers.validate_selected_element = lambda e: (
    {"is_valid": e.Id != 13, "reason": "no Out connector"})
gas_run.pipe_graph.check_gas_pipe = lambda e: (
    {"ok": e.Id != 21, "reason": "system is Domestic Cold Water"})
gas_run.tie_in.resolve_new_pipe_upstream = lambda e, doc: (
    {"status": "ok", "reason": "one end is open"} if e.Id != 22
    else {"status": "error", "reason": "both ends new"})

s = gas_run.resolve_start(None, None, ok)
check("no element", s["stop"] == "Could not retrieve selected element.")
s = gas_run.resolve_start(None, Elem(12), ok)
check("meter ok", s["stop"] is None and s["kind"] == "meter"
      and s["upstream"] is None)
s = gas_run.resolve_start(None, Elem(13), ok)
check("bad meter", "no Out connector" in s["stop"])
s = gas_run.resolve_start(None, Elem(21, pipe=True), ok)
check("not a gas pipe", "not a gas pipe" in s["stop"] and s["kind"] == "pipe")
s = gas_run.resolve_start(None, Elem(22, pipe=True), ok)
check("no upstream end", "both ends new" in s["stop"])
s = gas_run.resolve_start(None, Elem(23, pipe=True), ok)
# Colin 2026-11-20: no existing given = full new construction
check("pipe with no existing is new construction",
      s["stop"] is None and s["existing"] == {"new_construction": True}
      and "New construction" in s["existing_text"])
s = gas_run.resolve_start(None, Elem(23, pipe=True), dict(ok, existing=ex))
check("pipe with existing", s["stop"] is None
      and dict(s["existing"], new_construction=None)
      == dict(ex, new_construction=None)
      and s["existing"]["new_construction"] is False
      and s["upstream"]["reason"] == "one end is open")
nc = {"new_construction": True}
s = gas_run.resolve_start(None, Elem(23, pipe=True), dict(ok, existing=nc))
check("pipe new construction", s["stop"] is None and s["existing"] == nc
      and "New construction" in s["existing_text"]
      and s["check"].startswith("New gas piping connection"))
# Connect into Existing takes the first new pipe only (tie_in.check_gas_start
# pipe_only, main's Scope, 2026-11-20): a meter start is refused.
s = gas_run.resolve_start(None, Elem(12), dict(ok, existing=ex))
check("connect into existing refuses a meter",
      "first new gas pipe" in (s["stop"] or "")
      and "Nothing was changed" in s["stop"])
s = gas_run.resolve_start(None, Elem(12), dict(ok, existing=nc))
check("new construction scope still takes the meter", s["stop"] is None
      and s["kind"] == "meter" and s["existing"] is None)
check("stop says nothing was changed", "Nothing was changed" in
      gas_run.resolve_start(None, Elem(13), ok)["stop"])


# --- envelope_data over a fake outcome -------------------------------------------
class Node(object):
    def __init__(self, n, fixture=False, name="", mbh=0.0):
        self.element_id = n
        self.is_gas_fixture = fixture
        self.fixture_name = name
        self.gas_load_mbh = mbh


class Edge(object):
    def __init__(self, n, to, mbh, dia):
        self.element_id = n
        self.to_node_id = to
        self.cumulative_load_mbh = mbh
        self.diameter_inches = dia
        self.length_feet = 10.0
        self.pipe = None


class Graph(object):
    pass


g = Graph()
g.nodes = {1: Node(1), 2: Node(2, True, "RTU-1", 100.0),
           3: Node(3, True, "WH-1", 76.0)}
g.edges = {10: Edge(10, 1, 176.0, 1.0), 11: Edge(11, 2, 100.0, 1.0),
           12: Edge(12, 3, 76.0, 0.75)}
g.longest_run = {"farthest_fixture_name": "RTU-1"}
g.existing_connections = []
gas_run.sizing_engine.nominal_for_diameter = lambda t, d: (
    "1 K&L" if d == 1.0 else "3/4 K&L")
outcome = {"status": "ok", "stop": None, "graph": g,
           "result": {"sizes": {10: "1-1/4 K&L", 11: "1 K&L", 12: "3/4 K&L"},
                      "table_id": "402.4(10)", "longest_run_ft": 99.45,
                      "table_length_used_ft": 100, "zones": [],
                      "segment_detail": [
                          {"pipe_id": 10, "demand_mbh": 176.0,
                           "demand_cfh": 209.5, "units": "CFH",
                           "capacity_mbh": 445, "note": ""}]},
           "existing_check": None, "warnings": {"prv": []}, "writes": None}
data = gas_run.envelope_data(outcome, 7, "meter", ok, False, "flagged")
check("flagged lists changed only", [p["id"] for p in data["pipes"]] == [10])
check("changed row", data["pipes"][0]["size"] == "1-1/4 K&L"
      and data["pipes"][0]["drawn"] == "1 K&L"
      and data["pipes"][0]["demand"] == 209.5)
check("totals", data["total_load_mbh"] == 176.0
      and data["farthest_fixture"] == "RTU-1"
      and data["table_row_ft"] == 100)
all_data = gas_run.envelope_data(outcome, 7, "meter", ok, False, "all")
check("all lists every pipe", [p["id"] for p in all_data["pipes"]]
      == [10, 11, 12])
check("stub named", all_data["pipes"][1]["stub_to"] == "RTU-1")
env = headless.result("size_gas", "ok", data=data)
check("envelope dumps", isinstance(json.dumps(env), str))

print("{} checks, {} failed".format(COUNT[0], len(FAILS)))
sys.exit(1 if FAILS else 0)
