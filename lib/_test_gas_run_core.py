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


RUN_STEPS = []
ROLL_BACK = set()


def _run_step(doc, name, action):
    """Stand-in for revit_helpers.run_in_transaction: runs the action and
    returns the error text for a step named in ROLL_BACK."""
    RUN_STEPS.append(name)
    action()
    for key in ROLL_BACK:
        if key in name:
            return "rolled back by Revit"
    return None


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
rh.set_fitting_size = lambda el, inches, piping_only=False: (True, "stub")
rh.reset_pipe_diameter_approach = lambda: None


def _run(write, themed, existing=None, decisions=None, upstream=None):
    g = _build(with_prv=False)
    gas_run.pipe_graph.build_network = lambda el, doc, up=None: g
    out = headless.RecordingOutput()
    prompter = headless.HeadlessPrompter(decisions or {})
    oc = gas_run.size_core(
        _Doc(), _settings(), object(), upstream, existing, out, prompter,
        write=write, run_step=_run_step, themed=themed,
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
check("three write steps ran", len(RUN_STEPS) >= 2
      and "Pipes" in RUN_STEPS[0] and "Fittings" in RUN_STEPS[1])
check("nothing rolled back", w["rolled_back"] == []
      and w["transaction_errors"] == [])

# pipe step rolled back: stop, later steps skipped, zero pipes counted
del RUN_STEPS[:]
ROLL_BACK.add("Pipes")
oc, out = _run(True, False, decisions={"presizing_warnings": "continue"})
ROLL_BACK.clear()
w = oc["writes"]
check("rollback stops the run", oc["status"] == "refused"
      and "rolled them back" in (oc["stop"] or ""))
check("rollback named", w["rolled_back"] == ["pipe sizes"]
      and w["pipe_transaction_error"])
check("fittings and stubs skipped after rollback", RUN_STEPS == [
    "RJA Tools - Size Gas Pipes"])
check("no pipe counted after rollback", w["pipes_sized"] == 0
      and w["written_ids"] == [] and w["approach"] is None)

# fitting step rolled back: stubs skipped, step named
del RUN_STEPS[:]
ROLL_BACK.add("Fittings")
oc, out = _run(True, False, decisions={"presizing_warnings": "continue"})
ROLL_BACK.clear()
w = oc["writes"]
check("fitting rollback keeps status ok", oc["status"] == "ok")
check("fitting rollback named", w["rolled_back"] == ["fittings"]
      and any("rolled back" in t for t in w["transaction_errors"]))
check("stubs skipped after fitting rollback",
      not any("Stub" in n for n in RUN_STEPS))

if FAILED:
    print("{} failed".format(len(FAILED)))
    sys.exit(1)
print("OK")
