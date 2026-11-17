# -*- coding: utf-8 -*-
"""One entry point for the MEP design agents (headless step 6).

Agents call these functions through the pyRevit agent runtime MCP with the
headless worktree's lib as workspace. Each function only adapts arguments to
the tool's own run_headless and returns its headless.result envelope; no tool
logic lives here.

- Writes are off by default: a check never writes. A designer passes
  write=True / make_sheets=True on purpose, in run_modify (Colin's Keep or
  Discard).
- Duct Performance needs engine='cpython'; on IronPython duct_perf_run
  refuses it before it runs (the runtime's IronPython fails inside the
  external static walk).
- Gas table_id is required, never defaulted; ask Colin from table_options().
- No reload_tools: the runtime re-imports workspace modules every run and
  each tool resets its own state (calc basis, pipe diameter, log) at entry.

IronPython 2.7 syntax.
"""
import sys

import headless

TOOLS = (
    {"operation": "size_water", "module": "water_run",
     "ids": "start_id: the RPZ (start 'rpz') or the tee on the existing cold "
            "main (start 'tee')",
     "engine": "ironpython", "mcp": "run_query; run_modify when write=True",
     "writes": "pipes, fittings, fixture stubs (write=True)",
     "decisions": [{"qid": "incomplete_system", "answer_with": "decisions"},
                   {"qid": "no_cold_connector", "answer_with": "decisions"}]},
    {"operation": "water_performance", "module": "water_perf_run",
     "ids": "start_id: the RPZ (start 'rpz') or the tee (start 'tee'); "
            "view_id for make_sheets (else the active view)",
     "engine": "ironpython", "mcp": "run_query; run_modify when make_sheets=True",
     "writes": "colored plan copy, sheet, schedule (make_sheets=True)",
     "decisions": [{"qid": "no_cold_connector", "answer_with": "decisions"}]},
    {"operation": "duct_performance", "module": "duct_perf_run",
     "ids": "unit_id: the unit, VAV box or exhaust fan the scope lists "
            "(scope rtu, fancoil, vav, vavbox, exhaust)",
     "engine": "cpython",
     "mcp": "run_query(engine='cpython'); run_modify(engine='cpython') when "
            "make_sheets=True",
     "writes": "colored plan copies, sheet (make_sheets=True)",
     "decisions": [{"qid": "fan_cfm", "answer_with": "fan_cfm"}]},
    {"operation": "size_gas", "module": "gas_run",
     "ids": "start_id: the gas meter, or the first new gas pipe at the "
            "connection to existing (then settings['existing'] is required)",
     "engine": "ironpython", "mcp": "run_query; run_modify when write=True",
     "writes": "pipes, fittings, stubs (write=True)",
     "decisions": [{"qid": "presizing_warnings", "answer_with": "decisions"}]},
    {"operation": "diagnose", "module": "diagnose_run",
     "ids": "start_id: gas meter or first new gas pipe, water RPZ or tee, or "
            "HVAC equipment, duct or terminal (discipline detected)",
     "engine": "ironpython", "mcp": "run_query", "writes": "none",
     "decisions": []},
)


def engine():
    """'ironpython' or 'cpython', the engine this run is on."""
    return "ironpython" if sys.platform == "cli" else "cpython"


def tools():
    """What an agent can call: ids, engine, MCP call, writes, decision qids.
    Settings and their defaults come from default_settings(operation)."""
    out = []
    for t in TOOLS:
        row = dict(t)
        del row["module"]
        out.append(row)
    return out


def _tool(operation):
    for t in TOOLS:
        if t["operation"] == operation:
            return t
    return None


def _module(name):
    return __import__(name)


def _guard(operation, call):
    """Run call(); a question that escapes a tool comes back as
    needs_decision, any other failure (an import error, code outside a
    tool's own try) as an error envelope with the traceback tail, so the
    agent always gets an envelope it can act on."""
    try:
        return call()
    except headless.NeedsDecision as exc:
        return headless.needs_decision(operation, exc)
    except Exception as exc:  # the envelope carries the failure
        import traceback
        try:  # the shared helper; revit_helpers itself may be what failed
            from revit_helpers import exc_text
            text = exc_text(exc)
        except Exception:
            text = repr(exc)
        return headless.result(operation, "error", errors=[
            "{}: {}".format(type(exc).__name__, text)],
            data={"traceback": traceback.format_exc().splitlines()[-12:]})


def default_settings(operation, scope="rtu", doc=None):
    """The tool's own defaults (the dialog's prefill). Duct Performance takes
    the scope; water tools read Project Information when doc is given."""
    t = _tool(operation)
    if t is None:
        raise ValueError("unknown operation: {}".format(operation))
    mod = _module(t["module"])
    if operation == "duct_performance":
        return mod.default_controls(scope)
    if operation in ("size_water", "water_performance"):
        return mod.default_settings(doc)
    if operation == "size_gas":
        return mod.default_settings()
    return {}


def table_options():
    """The IFGC tables to ask Colin from for size_gas."""
    return _module("gas_run").table_options()


def size_water(doc, start_id, start="rpz", settings=None, decisions=None,
               write=False, detail="flagged", include_log=False):
    return _guard("size_water", lambda: _module("water_run").run_headless(
        doc, start_id, start=start, settings=settings, decisions=decisions,
        write=write, detail=detail, include_log=include_log))


def water_performance(doc, start_id, start="rpz", settings=None,
                      decisions=None, make_sheets=False, view_id=None,
                      detail="flagged", include_log=False):
    return _guard("water_performance",
                  lambda: _module("water_perf_run").run_headless(
                      doc, start_id, start=start, settings=settings,
                      decisions=decisions, make_sheets=make_sheets,
                      view_id=view_id, detail=detail,
                      include_log=include_log))


def duct_performance(doc, unit_id, scope="rtu", settings=None, fan_cfm=None,
                     make_sheets=False, view_id=None, detail="flagged",
                     include_log=False):
    return _guard("duct_performance",
                  lambda: _module("duct_perf_run").run_headless(
                      doc, unit_id, scope=scope, settings=settings,
                      fan_cfm=fan_cfm, make_sheets=make_sheets,
                      view_id=view_id, detail=detail,
                      include_log=include_log))


def size_gas(doc, start_id, table_id=None, settings=None, decisions=None,
             write=False, detail="flagged", include_log=False):
    """table_id is required (argument or settings): ask Colin from
    table_options(), never assume. gas_run refuses a missing or unknown one."""
    def call():
        merged = dict(settings or {})
        if table_id is not None:
            if merged.get("table_id") not in (None, table_id):
                return headless.result("size_gas", "refused", errors=[
                    "table_id given twice with different values ({!r} and "
                    "{!r}); pass it once.".format(table_id,
                                                   merged["table_id"])])
            merged["table_id"] = table_id
        return _module("gas_run").run_headless(
            doc, start_id, settings=merged, decisions=decisions, write=write,
            detail=detail, include_log=include_log)
    return _guard("size_gas", call)


def diagnose(doc, start_id, include_report=False, include_log=False):
    return _guard("diagnose", lambda: _module("diagnose_run").run_headless(
        doc, start_id, include_report=include_report,
        include_log=include_log))
