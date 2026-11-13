# -*- coding: ascii -*-
"""Tests for the Water Performance headless entry (water_perf_run): settings,
plan view check, refusals, created ids and the result envelope.

Plain CPython with the Revit assemblies stubbed and perf_core replaced by a
fake, so this checks the wrapper only. Judging, pressure and pumps are
covered by _test_water_performance, _test_water_pressure, _test_water_pump;
the model side needs a live Revit run.

Run:  python _test_water_perf_headless.py
"""

import json
import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


class _StubMeta(type):
    """Any class attribute (enum member such as OST_DuctTerminal) is a
    distinct int, so int(BuiltInCategory.X) works at import."""
    _next = [1000]

    def __getattr__(cls, name):
        if name.startswith("__"):
            raise AttributeError(name)
        _StubMeta._next[0] += 1
        value = _StubMeta._next[0]
        setattr(cls, name, value)
        return value


_Stub = _StubMeta("_Stub", (object,), {
    "__init__": lambda self, *args, **kwargs: setattr(self, "args", args)})


class _StubModule(types.ModuleType):
    """Any name imported from it is a stub class."""

    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)
        cls = type(name, (_Stub,), {})
        setattr(self, name, cls)
        return cls


def _install_stubs():
    names = ["Autodesk", "Autodesk.Revit", "Autodesk.Revit.DB",
             "Autodesk.Revit.DB.Mechanical", "Autodesk.Revit.DB.Plumbing",
             "Autodesk.Revit.UI", "Autodesk.Revit.UI.Selection", "System",
             "System.Collections", "System.Collections.Generic"]
    for name in names:
        sys.modules[name] = _StubModule(name)
    sys.modules["Autodesk"].Revit = sys.modules["Autodesk.Revit"]
    sys.modules["Autodesk.Revit"].DB = sys.modules["Autodesk.Revit.DB"]
    sys.modules["Autodesk.Revit"].UI = sys.modules["Autodesk.Revit.UI"]
    sys.modules["Autodesk.Revit.UI"].Selection = sys.modules[
        "Autodesk.Revit.UI.Selection"]
    sys.modules["ui_helpers"] = types.ModuleType("ui_helpers")
    pyrevit = types.ModuleType("pyrevit")
    pyrevit.HOST_APP = type("HostApp", (object,), {"version": "2024"})()
    sys.modules["pyrevit"] = pyrevit


_install_stubs()

import water_perf_run as wp      # noqa: E402
import water_performance         # noqa: E402
import water_pressure            # noqa: E402
import water_pump                # noqa: E402
import water_run                 # noqa: E402

FAILS = []
COUNT = [0]


def check(name, cond):
    COUNT[0] += 1
    if not cond:
        FAILS.append(name)
        print("FAIL " + name)


# --- settings --------------------------------------------------------------
d = wp.default_settings()
check("perf keys", set(["velocity", "speed_tolerance", "wsfu_calcs",
                        "pressure_loss", "start_psi", "pump_duty",
                        "circuit_setter_psi", "pump_safety_factor",
                        "recirc_delta_t_f", "job"]) <= set(d))
check("dialog ticks", d["velocity"] and d["wsfu_calcs"]
      and not d["pressure_loss"] and not d["pump_duty"])
check("start psi default", d["start_psi"] == water_pressure.DEFAULT_START_PSI)
check("setter default", d["circuit_setter_psi"] == water_pump.DEFAULT_CIRCUIT_SETTER_PSI)
check("tolerance default", d["speed_tolerance"] == water_performance.SPEED_TOLERANCE)
check("defaults valid", wp.validate_settings(d) == [])
none = dict(d, velocity=False, wsfu_calcs=False)
check("one check needed", any("at least one" in p for p in wp.validate_settings(none)))
check("tolerance high", wp.validate_settings(dict(d, speed_tolerance=1.5))
      == ["speed_tolerance shall be at most 1.0"])
check("start psi only when on", wp.validate_settings(dict(d, start_psi=0)) == [])
check("start psi checked", wp.validate_settings(
    dict(d, pressure_loss=True, start_psi=0)) == ["start_psi shall be above 0.0"])
check("pump sf", wp.validate_settings(
    dict(d, pump_duty=True, pump_safety_factor=2)) == [
        "pump_safety_factor shall be at most 1.0"])
check("recirc checked too", wp.validate_settings(dict(d, recirc_delta_t_f=40)) != [])
try:
    wp.merge_settings({"velocity_check": True})
    check("unknown key", False)
except ValueError:
    check("unknown key", True)


# --- plan view check --------------------------------------------------------
class Plan(wp.ViewPlan):
    IsTemplate = False


class Sheet(object):
    IsTemplate = False


check("no sheets, no view needed", wp.check_plan_view(None, d, make_sheets=False) is None)
check("velocity off, no view needed", wp.check_plan_view(
    Sheet(), dict(d, velocity=False)) is None)
check("sheet view refused", "floor plan" in wp.check_plan_view(Sheet(), d))
check("plan ok", wp.check_plan_view(Plan(), d) is None)
tpl = Plan()
tpl.IsTemplate = True
check("template refused", wp.check_plan_view(tpl, d) is not None)


# --- _record ----------------------------------------------------------------
class E(object):
    def __init__(self, n, number=None):
        self.Id = n
        if number:
            self.SheetNumber = number


wp.revit_helpers.eid_int = lambda i: int(i)
created = {}
wp._record(created, "sheet", E(7, "WP-M101-1"))
wp._record(created, "wsfu", (E(8), E(9, "WS-1")))
wp._record(created, "tables", None)
wp._record(created, "x", (E(10), None))
check("record", created == {"sheet_id": 7, "sheet_number": "WP-M101-1",
                            "wsfu_view_id": 8, "wsfu_sheet_id": 9,
                            "wsfu_sheet_number": "WS-1", "x_view_id": 10})


# --- run_headless with a fake core -------------------------------------------
class Doc(object):
    ActiveView = Sheet()


class Elem(object):
    def __init__(self, n):
        self.Id = n


wp.water_run._element = lambda doc, i: Elem(i) if i else None
wp.water_run.revit_helpers.is_pipe = lambda e: False
wp.water_run.revit_helpers.get_connectors = lambda e: [
    {"system_type": wp.shared_params.SYSTEM_DOMESTIC_COLD_WATER}]


def res(n, label):
    r = water_performance.PipeResult(n)
    r.label = label
    r.reason = "reason {}".format(n)
    r.gpm = 5.0
    return r


class Item(object):
    pass


CALLS = []


def fake_core(doc, settings, origin, upstream, view, output, make_sheets=True):
    CALLS.append((make_sheets, view))
    if origin.Id == 99:
        return {"status": "refused", "message": "No pipes were found on this "
                "system.", "warnings": [], "created": {}}
    return {
        "status": "ok", "message": "", "graph": None,
        "sizing": {"totals": {"pipe_count": 3}},
        "recirc": None,
        "results": {1: res(1, water_performance.GREEN),
                    2: res(2, water_performance.RED),
                    3: res(3, water_performance.GRAY)},
        "items": [], "checks": {"errors": [], "warnings": [],
                                "ready_for_sizing": True},
        "system_errors": [("WARNING", "1 circuit setter for 2 legs")],
        "pressure": None, "pumps": None,
        "warnings": ["Pump duty point failed, not reported: x"],
        "created": {"colored_view_id": 55} if make_sheets else {}}


wp.perf_core = fake_core

r = wp.run_headless(Doc(), 5)
check("ok", r["status"] == "ok")
check("results only by default, no view", CALLS[-1] == (False, None))
check("flagged pipes", [p["id"] for p in r["data"]["pipes"]] == [2, 3])
check("counts", r["data"]["counts"][water_performance.RED] == 1)
check("system errors", r["data"]["system_errors"] == [
    {"severity": "WARNING", "text": "1 circuit setter for 2 legs"}])
check("core warning carried", r["warnings"] == ["Pump duty point failed, not reported: x"])
check("no pressure key", "pressure" not in r["data"])
check("json", isinstance(json.dumps(r), str))

r = wp.run_headless(Doc(), 5, detail="all")
check("all pipes", len(r["data"]["pipes"]) == 3)

r = wp.run_headless(Doc(), 5, make_sheets=True)
check("sheets need a plan", r["status"] == "refused" and "floor plan" in r["errors"][0])
wp.water_run._element = lambda doc, i: Plan() if i == 300 else (Elem(i) if i else None)
r = wp.run_headless(Doc(), 5, make_sheets=True, view_id=300)
check("sheets with plan", r["status"] == "ok" and r["created"] == {"colored_view_id": 55})
r = wp.run_headless(Doc(), 5, make_sheets=True, settings={"velocity": False})
check("sheets without velocity skip view", r["status"] == "ok" and CALLS[-1][1] is None)

r = wp.run_headless(Doc(), 99)
check("no pipes refused", r["status"] == "refused" and "No pipes" in r["errors"][0])
r = wp.run_headless(Doc(), 0)
check("no element refused", r["status"] == "refused")
r = wp.run_headless(Doc(), 5, settings={"velocity": False, "wsfu_calcs": False})
check("nothing ticked refused", r["status"] == "refused")
r = wp.run_headless(Doc(), 5, start="meter")
check("bad start refused", r["status"] == "refused")

print("{} checks, {} failed".format(COUNT[0], len(FAILS)))
sys.exit(1 if FAILS else 0)
