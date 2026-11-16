# -*- coding: ascii -*-
"""Tests for the Diagnose headless entry (diagnose_run.run_headless and the
outcome it wraps).

Plain CPython with the Revit assemblies stubbed and the discipline diagnoses
replaced by fakes, so this checks the wrapper only: refusals, the envelope
and the shared helpers. The traversals are covered by their own tests and by
the live bench runs (gas G1, G6 pipe start, G7 on 2026-11-16, Revit 2027).

Run:  python _test_diagnose_headless.py
"""

import json
import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


class _AnyMember(type):
    """Class level attributes (BuiltInCategory.OST_..., enums) read as
    distinct ints, which is all module level code does with them."""
    _next = [1000]

    def __getattr__(cls, name):
        if name.startswith("__"):
            raise AttributeError(name)
        _AnyMember._next[0] += 1
        setattr(cls, name, _AnyMember._next[0])
        return _AnyMember._next[0]


class _AnyModule(types.ModuleType):
    """Every attribute is a fresh empty class: the import chain behind
    diagnose_run (hvac_graph, water_report, ...) names many DB types."""

    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)
        cls = _AnyMember(name, (object,), {})
        setattr(self, name, cls)
        return cls


def _install_stubs():
    autodesk = types.ModuleType("Autodesk")
    revit = types.ModuleType("Autodesk.Revit")
    db = _AnyModule("Autodesk.Revit.DB")
    ui = _AnyModule("Autodesk.Revit.UI")
    selection = types.ModuleType("Autodesk.Revit.UI.Selection")
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
import diagnose_run    # noqa: E402

FAILS = []
COUNT = [0]


def check(name, cond):
    COUNT[0] += 1
    if not cond:
        FAILS.append(name)
        print("FAIL " + name)


class Cat(object):
    Name = "Mechanical Equipment"
    Id = 77


class Elem(object):
    def __init__(self, n):
        self.Id = n
        self.Category = Cat()


rh = diagnose_run.revit_helpers
rh.eid_int = lambda i: int(i)
diagnose_run.eid_int = lambda i: int(i)
rh.clear_log = lambda: None
rh.lib_folder = lambda: "lib"
ELEMENTS = {5: Elem(5), 6: Elem(6), 7: Elem(7)}
rh.get_element_by_id = lambda doc, i: ELEMENTS.get(int(i))

# --- element not found ---------------------------------------------------------
r = diagnose_run.run_headless(None, 99)
check("missing element refused", r["status"] == "refused"
      and "99" in r["errors"][0])

# --- discipline cannot be determined -----------------------------------------
def classify(e):
    if e.Id == 6:
        raise ValueError("The selected element has no connectors")
    if e.Id == 7:
        raise RuntimeError("boom")
    return "GAS", "has a piping connector"


diagnose_run.classify_discipline = classify
r = diagnose_run.run_headless(None, 6)
check("no discipline refused", r["status"] == "refused"
      and r["data"]["discipline"] is None
      and "no connectors" in r["errors"][0])
r = diagnose_run.run_headless(None, 7)
check("exception is error", r["status"] == "error" and "boom" in r["errors"][0])

# --- a discipline run ------------------------------------------------------------
diagnose_run.diagnose_gas = lambda doc, e, out: {
    "summary": {"total_load_mbh": 256.0, "disconnected": []},
    "errors": [], "warnings": ["1 disconnected element(s) found."],
    "ready": True, "report": "=== DIAGNOSTIC REPORT ==="}
r = diagnose_run.run_headless(None, 5)
check("gas ok", r["status"] == "ok" and r["data"]["discipline"] == "GAS"
      and r["data"]["ready"] is True
      and r["data"]["summary"]["total_load_mbh"] == 256.0)
check("warnings carried", r["warnings"] == ["1 disconnected element(s) found."])
check("selected", r["data"]["selected"] == {"id": 5,
                                            "category": "Mechanical Equipment"})
check("no report by default", "report" not in r["data"] and "log" not in r)
r = diagnose_run.run_headless(None, 5, include_report=True, include_log=True)
check("report when asked", r["data"]["report"] == "=== DIAGNOSTIC REPORT ===")
check("log when asked", any("Detected discipline" in str(p)
                            for k, p in r["log"]))
check("envelope dumps", isinstance(json.dumps(r), str))

# --- the button's stop path ---------------------------------------------------------
stopped = diagnose_run._stopped("Traversal failed", "Diagnose - Traversal Error")
check("stopped shape", stopped["status"] == "refused"
      and stopped["errors"] == ["Traversal failed"] and not stopped["ready"])
diagnose_run.diagnose_gas = lambda doc, e, out: diagnose_run._stopped(
    "Traversal failed", "Diagnose - Traversal Error")
r = diagnose_run.run_headless(None, 5)
check("traversal stop refused", r["status"] == "refused"
      and r["errors"] == ["Traversal failed"])

# --- _ids ------------------------------------------------------------------------------
check("_ids mixed", diagnose_run._ids([Elem(3), 4, (8, "DCW")]) == [3, 4, 8])

print("{} checks, {} failed".format(COUNT[0], len(FAILS)))
sys.exit(1 if FAILS else 0)
