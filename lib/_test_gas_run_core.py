# -*- coding: ascii -*-
"""Smoke tests for gas_run.size_core and _write_sizes after the main merge
(2026-11-20): the themed gas_console sections, Connect into Existing (new
construction), the transition rules and the write results.

Plain CPython with the Revit assemblies stubbed (from _test_prv_zones) and a
hand built graph. It does NOT exercise a real Revit model: live behaviour of
the Size Gas button through gas_run needs a Revit run.

Run:  python -I _test_gas_run_core.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_prv_zones import _build   # installs the stubs

import gas_run          # noqa: E402
import headless         # noqa: E402

FAILED = []


def check(name, cond):
    if not cond:
        FAILED.append(name)
        print("FAIL " + name)


class _Tx(object):
    def __init__(self, doc, name):
        self.name = name

    def Start(self):
        pass

    def Commit(self):
        pass

    def RollBack(self):
        pass


class _Doc(object):
    def Regenerate(self):
        pass


def _settings():
    s = gas_run.default_settings()
    s["table_id"] = gas_run.table_options()[0]["table_id"]
    s["pipe_material"] = gas_run.table_options()[0]["material"]
    return s


rh = gas_run.revit_helpers
rh.is_new_construction = lambda el, doc: True
rh.is_valid_element = lambda el: True
rh.is_transition = lambda el: False
rh.set_fitting_size = lambda el, inches: (True, "stub")
rh.reset_pipe_diameter_approach = lambda: None


def _run(write, themed, existing=None, decisions=None, upstream=None):
    g = _build(with_prv=False)
    gas_run.pipe_graph.build_network = lambda el, doc, up=None: g
    out = headless.RecordingOutput()
    prompter = headless.HeadlessPrompter(decisions or {})
    oc = gas_run.size_core(
        _Doc(), _settings(), object(), upstream, existing, out, prompter,
        write=write, transaction_factory=_Tx, themed=themed,
        inputs_extra={"selected": "Element ID 1", "check": "Meter ok."})
    return oc, out


# plan only, markdown twins
oc, out = _run(False, False)
check("plan ok", oc["status"] == "ok" and oc["result"] is not None)
kinds = set(k for k, _p in out.log)
check("markdown only when not themed", kinds == set(["md"]))
check("inputs printed", any("Element ID 1" in str(p) for _k, p in out.log))
check("warnings carry the new groups",
      "non_gas" in oc["warnings"] and "bare_end" in oc["warnings"])

# plan only, themed html
oc, out = _run(False, True)
check("themed prints html", "html" in set(k for k, _p in out.log))

# new construction: nothing applied, nothing checked
oc, out = _run(False, False, existing={"new_construction": True})
check("new construction has no existing check", oc["existing_check"] is None)

# write: fake pipes are None so every pipe is a failure row, not a crash
oc, out = _run(True, False, decisions={"presizing_warnings": "continue"})
w = oc["writes"]
check("write ran", w is not None and w["pipe_transaction_error"] is None)
check("pipe failures listed", len(w["pipe_failures"]) > 0
      and w["pipe_failures"][0].startswith("Pipe "))
fit = w["fittings"]
check("fit counts have transitions", "transitions" in fit
      and "transitions_removed" in fit and "failure_rows" in fit)
text = "\n".join(str(p) for _k, p in out.log)
check("write results printed", "Write results" in text)
check("summary printed", "Summary" in text)

if FAILED:
    print("{} failed".format(len(FAILED)))
    sys.exit(1)
print("OK")
