# -*- coding: ascii -*-
"""Smoke tests for diagnose_run.diagnose_gas after the gas review port
(2026-11-20): the one gas start check on a pipe pick (upstream passed to the
walk), the shared summary counts, data errors, the existing appliance table
and the summary data the headless envelope carries.

Plain CPython with the Revit assemblies stubbed (from _test_prv_zones) and a
hand built graph. It does NOT exercise a real Revit model: the Diagnose gas
branch inside Revit needs a live run.

Run:  python -I _test_diagnose_gas_run.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_prv_zones import _build   # installs the stubs


import types   # noqa: E402

# The HVAC and water modules diagnose_run imports are never called by the gas
# branch; empty stand-ins keep their Revit enums out of these stubs.
for _name in ("hvac_graph", "water_graph", "water_checks", "water_recirc",
              "water_run", "water_report"):
    sys.modules.setdefault(_name, types.ModuleType(_name))

import diagnose_run     # noqa: E402
import headless         # noqa: E402

FAILED = []
COUNT = [0]


def check(name, cond):
    COUNT[0] += 1
    if not cond:
        FAILED.append(name)
        print("FAIL " + name)


class _Elem(object):
    def __init__(self, n, pipe=False):
        self.Id = n
        self.pipe = pipe


rh = diagnose_run.revit_helpers
rh.is_pipe = lambda e: e.pipe
rh.get_piping_connectors = lambda e: [{"direction": "Out",
                                       "is_connected": True}]
diagnose_run.report_generator.format_diagnostic_output = (
    lambda graph, element, mid_stream_prv=None, tie_in_check=None:
    "PASTE tie_in={}".format(tie_in_check is not None))

WALKS = []


def _walk(graph):
    def build(element, doc, upstream=None):
        WALKS.append(upstream)
        return graph
    diagnose_run.pipe_graph.build_network = build


def _text(out):
    return "\n".join(str(p) for _k, p in out.log)


# --- meter pick: summary, counts, data errors, existing appliances ----------
g = _build()
# the stand-in PRV is named like a real one, so shared_params.accessory_kind
# classifies it by family name as Diagnose does
g.nodes[3].family_name = "Gas Pressure Regulator"
g.data_errors = [{"element_id": 4, "name": "RTU-1", "kind": "zero_load",
                  "message": "GAS_LOAD_MBH is 0"}]
g.existing_appliances = [5]
_walk(g)
out = headless.RecordingOutput()
part = diagnose_run.diagnose_gas(None, _Elem(1), out)
text = _text(out)
summary = part["summary"]
check("meter walk has no upstream", WALKS[-1] is None)
check("meter is not a pipe start", summary["pipe_start"] is False
      and summary["tie_in_check"] is None)
check("summary counts from report_generator",
      summary["prvs"] == 1 and summary["tees"] >= 1
      and summary["fittings_other"] >= 0)
check("prv accessory row", any(r["kind"] == "PRV" and r["id"] == 3
                               for r in summary["accessories"]))
check("total load from pipe_graph", summary["total_load_mbh"]
      == diagnose_run.pipe_graph.total_connected_mbh(g))
check("data errors listed", len(summary["data_errors"]) == 1
      and "RTU-1 (element 4)" in summary["data_errors"][0])
check("data errors printed", "Data errors (0 MBH gas fixtures)" in text)
check("existing appliance table", "Existing Appliances on New Piping" in text
      and summary["existing_appliances"][0]["id"] == 5)
check("paste block printed", "PASTE tie_in=False" in text)
check("zone runs carried", isinstance(summary["zone_runs"], list))

# --- pipe pick refused by the one start check --------------------------------
diagnose_run.tie_in.check_gas_start = lambda e, doc, pipe_only=False: {
    "ok": False, "kind": "tie_in", "upstream": None, "gas_check": None,
    "message": "This is not a gas pipe: system is Domestic Cold Water."}
before = len(WALKS)
out = headless.RecordingOutput()
part = diagnose_run.diagnose_gas(None, _Elem(21, pipe=True), out)
check("refused pipe stops", part["status"] == "refused"
      and "not a gas pipe" in part["stop"]
      and part["stop_title"] == "Diagnose - Gas")
check("refused pipe never walks", len(WALKS) == before)

# --- pipe pick accepted: upstream reaches the walk, meter checks skipped -----
UP = {"status": "ok", "reason": "one end is open"}
diagnose_run.tie_in.check_gas_start = lambda e, doc, pipe_only=False: {
    "ok": True, "kind": "tie_in", "upstream": UP, "gas_check": {"ok": True},
    "message": "New gas piping connection. Upstream end found because one "
               "end is open."}
seen = []
real_validation = diagnose_run.report_generator.gas_validation


def _validation(graph, origin_connectors=None, mid_stream_prv=None):
    seen.append(origin_connectors)
    return real_validation(graph, origin_connectors, mid_stream_prv)


diagnose_run.report_generator.gas_validation = _validation
_walk(_build())
out = headless.RecordingOutput()
part = diagnose_run.diagnose_gas(None, _Elem(22, pipe=True), out)
diagnose_run.report_generator.gas_validation = real_validation
check("tie in upstream passed to the walk", WALKS[-1] is UP)
check("tie in skips the meter connectors", seen == [None])
check("tie in summary", part["summary"]["pipe_start"] is True
      and "Upstream end found" in part["summary"]["tie_in_check"])
check("tie in paste block", "PASTE tie_in=True" in _text(out))
check("tie in warning printed", "TIE-IN PIPE" in _text(out))

print("{} checks, {} failed".format(COUNT[0], len(FAILED)))
sys.exit(1 if FAILED else 0)
