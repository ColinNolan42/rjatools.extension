# -*- coding: ascii -*-
"""Tests for the Size Water headless entry (water_run.run_headless and its
settings, origin checks and result envelope).

Plain CPython with the Revit assemblies stubbed and size_core replaced by a
fake, so this checks the wrapper only: settings, refusals, decisions and the
JSON envelope. The traversal and sizing themselves are covered by the other
_test_water_* files and need a live Revit run for the model side.

Run:  python _test_water_run_headless.py
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
                 "SpecTypeId", "Transaction"):
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
import water_run       # noqa: E402

FAILS = []
COUNT = [0]


def check(name, cond):
    COUNT[0] += 1
    if not cond:
        FAILS.append(name)
        print("FAIL " + name)


# --- settings --------------------------------------------------------------
d = water_run.default_settings()
check("default keys", sorted(d) == ["by", "job", "job_number",
                                    "recirc_delta_t_f",
                                    "recirc_heat_loss_btu_per_100ft"])
check("default heat loss", d["recirc_heat_loss_btu_per_100ft"]
      == float(shared_params.DEFAULT_RECIRC_HEAT_LOSS_BTU_PER_100FT))
check("default dt", d["recirc_delta_t_f"]
      == float(shared_params.DEFAULT_RECIRC_DELTA_T_F))
check("defaults valid", water_run.validate_settings(d) == [])
m = water_run.merge_settings({"recirc_delta_t_f": 15})
check("merge override", m["recirc_delta_t_f"] == 15)
try:
    water_run.merge_settings({"delta_t": 15})
    check("unknown key raises", False)
except ValueError as exc:
    check("unknown key raises", "delta_t" in str(exc))
bad = dict(d, recirc_delta_t_f=25, recirc_heat_loss_btu_per_100ft=0)
check("two problems", len(water_run.validate_settings(bad)) == 2)
check("not a number", water_run.validate_settings(
    dict(d, recirc_delta_t_f="x")) == ["recirc_delta_t_f is not a number"])

# --- refusals before any Revit call ----------------------------------------
r = water_run.run_headless(None, 1, settings={"recirc_delta_t_f": 30})
check("refused dt", r["status"] == "refused" and "recirc_delta_t_f" in r["errors"][0])
r = water_run.run_headless(None, 1, start="meter")
check("refused start", r["status"] == "refused")
r = water_run.run_headless(None, 1, detail="some")
check("refused detail", r["status"] == "refused")
r = water_run.run_headless(None, 1, settings={"typo": 1})
check("refused unknown", r["status"] == "refused" and "typo" in r["errors"][0])
r = water_run.run_headless(None, 1, decisions={"incomplete_system": "yes"})
check("refused bad decision", r["status"] == "refused")


# --- resolve_origin ----------------------------------------------------------
class Elem(object):
    def __init__(self, n, pipe=False, cold=True):
        self.Id = n
        self.pipe = pipe
        self.cold = cold


water_run.revit_helpers.is_pipe = lambda e: e.pipe
water_run.revit_helpers.eid_int = lambda i: int(i)
water_run.revit_helpers.get_connectors = lambda e: (
    [{"system_type": shared_params.SYSTEM_DOMESTIC_COLD_WATER}] if e.cold else [])

p = headless.HeadlessPrompter()
check("none element", water_run.resolve_origin(None, None, "rpz", p)[2]
      == "Could not read the selected element.")
check("pipe picked", "pipe was picked" in
      water_run.resolve_origin(None, Elem(1, pipe=True), "rpz", p)[2])
ok = water_run.resolve_origin(None, Elem(5), "rpz", p)
check("rpz ok", ok[0].Id == 5 and ok[1] is None and ok[2] is None)
try:
    water_run.resolve_origin(None, Elem(6, cold=False), "rpz", p)
    check("no cold asks", False)
except headless.NeedsDecision as exc:
    check("no cold asks", exc.qid == "no_cold_connector" and exc.suggested == "stop")
stop = water_run.resolve_origin(None, Elem(6, cold=False), "rpz",
                                headless.HeadlessPrompter({"no_cold_connector": "stop"}))
check("no cold stop", stop[2] == water_run._DECLINED)
go = water_run.resolve_origin(None, Elem(6, cold=False), "rpz",
                              headless.HeadlessPrompter({"no_cold_connector": "continue"}))
check("no cold continue", go[0].Id == 6 and go[2] is None)
check("tee on a pipe", "TEE" in water_run.resolve_origin(
    None, Elem(7, pipe=True), "tee", p)[2])


# --- run_headless with a fake core -------------------------------------------
class Seg(object):
    def __init__(self, n, flag="", size='1"', keep=False, resized=False):
        self.element_id = n
        self.system = "DCW"
        self.demand_wsfu = 10.0
        self.demand_gpm = 8.0
        self.circ_gpm = None
        self.is_return = False
        self.existing_size_in = 0.75 if resized else 1.0
        self.nominal_size = size
        self.size_inches = 1.0 if size else None
        self.rule = "table"
        self.flag = flag
        self.length_feet = 12.5
        self.fixture_count = 3
        self.exceeds_table = False
        self.keep_existing = keep
        self.required_nominal = None
        self.upsize_needed = False


class Finding(object):
    def __init__(self, message):
        self.message = message


ELEMENTS = {5: Elem(5), 6: Elem(6, cold=False)}
water_run._element = lambda doc, i: ELEMENTS.get(i)
CALLS = []


def fake_core(doc, settings, origin, upstream, output, prompter, write=True,
              transaction_factory=None):
    CALLS.append((write, transaction_factory))
    output.print_md("core ran")
    if settings["job"] == "incomplete":
        prompter.confirm("incomplete_system", "Open ends. Continue?",
                         suggested="stop")
    prompter.notify("2 pipe(s) over the table")
    return {
        "status": "ok", "graph": None,
        "detection": {"candidates": [{"name": "HWR", "detected": True,
                                      "pipe_count": 4, "reasons": ["pump"]}],
                      "certain": True},
        "checks": {"errors": [], "warnings": [Finding("short stub")],
                   "ready_for_sizing": True},
        "recirc": {"total_gpm": 1.5, "root_gpm": {9: 1.5}, "legs": [1, 2, 3],
                   "unrooted_pipe_ids": [], "heat_loss_btu_per_100ft": 5000.0,
                   "delta_t_f": 20.0},
        "sizing": {"segments": [Seg(11), Seg(12, flag="min 3/4"), Seg(15, resized=True),
                                Seg(13, size=None), Seg(14, keep=True)],
                   "totals": {"pipe_count": 4, "sized_count": 3},
                   "flags": ["min 3/4 on 12"]},
        "writes": None, "errors": []}


water_run.size_core = fake_core

r = water_run.run_headless(None, 5, write=False)
check("ok status", r["status"] == "ok")
check("plan only, no factory", CALLS[-1] == (False, None))
check("flagged pipes", [p["id"] for p in r["data"]["pipes"]] == [12, 15, 13])
check("changed field", [p["changed"] for p in r["data"]["pipes"]] == [False, True, False])
check("origin", r["data"]["origin"] == {"id": 5, "start": "rpz", "tee_reason": None})
check("notify to warnings", r["warnings"] == ["2 pipe(s) over the table"])
check("recirc summary", r["data"]["recirc"]["circuit_setters"] == 3
      and r["data"]["recirc"]["root_gpm"] == {"9": 1.5})
check("check warnings", r["data"]["check_warnings"] == ["short stub"])
check("no log", "log" not in r)
check("json", isinstance(json.dumps(r), str))

r = water_run.run_headless(None, 5, write=False, detail="all", include_log=True)
check("all pipes", len(r["data"]["pipes"]) == 5)
check("log kept", any(k == "md" and v == "core ran" for k, v in r["log"]))

r = water_run.run_headless(None, 6, write=False)
check("needs no_cold", r["status"] == "needs_decision"
      and r["decision"]["qid"] == "no_cold_connector")
r = water_run.run_headless(None, 6, write=False,
                           decisions={"no_cold_connector": "stop"})
check("declined is cancelled", r["status"] == "cancelled")
r = water_run.run_headless(None, 99, write=False)
check("missing element refused", r["status"] == "refused")

r = water_run.run_headless(None, 5, write=False, settings={"job": "incomplete"})
check("needs incomplete", r["status"] == "needs_decision"
      and r["decision"]["suggested"] == "stop")


def boom(*a, **k):
    raise RuntimeError("traversal blew up")


water_run.size_core = boom
r = water_run.run_headless(None, 5, write=False)
check("error envelope", r["status"] == "error" and "blew up" in r["errors"][0])

print("{} checks, {} failed".format(COUNT[0], len(FAILS)))
sys.exit(1 if FAILS else 0)
