# Plain CPython tests for agent_api.py with fake tool modules.
# Run: python _test_agent_api.py
import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import headless
import agent_api as api

FAILS = []
COUNT = [0]


def check(name, cond):
    COUNT[0] += 1
    if not cond:
        FAILS.append(name)
        print("FAIL " + name)


CALLS = {}


def fake(name, tool, extra=None):
    m = types.ModuleType(name)
    m.TOOL = tool

    def run_headless(*args, **kwargs):
        CALLS[name] = (args, kwargs)
        return headless.result(tool, "ok", data={"echo": True})
    m.run_headless = run_headless
    for k, v in (extra or {}).items():
        setattr(m, k, v)
    sys.modules[name] = m
    return m


fake("water_run", "size_water",
     {"default_settings": lambda doc=None: {"recirc_delta_t_f": 10.0}})
fake("water_perf_run", "water_performance",
     {"default_settings": lambda doc=None: {"velocity": True}})
fake("duct_perf_run", "duct_performance",
     {"default_controls": lambda scope="rtu": {"scope_seen": scope}})
fake("gas_run", "size_gas",
     {"default_settings": lambda: {"table_id": None},
      "table_options": lambda: [{"table_id": "402.4(2)", "material": "a"},
                                {"table_id": "402.4(2)", "material": "b"},
                                {"table_id": "402.4(5)", "material": "a"}]})
fake("diagnose_run", "diagnose")

# tools() matches each module's TOOL, no module key leaks
listed = api.tools()
check("tools count", len(listed) == 5)
for row in listed:
    mod = sys.modules[api._tool(row["operation"])["module"]]
    check("tool name " + row["operation"], mod.TOOL == row["operation"])
    check("no module key " + row["operation"], "module" not in row)
    for key in ("ids", "engine", "mcp", "writes", "decisions"):
        check("key %s %s" % (key, row["operation"]), key in row)
check("duct cpython", api._tool("duct_performance")["engine"] == "cpython")

# default_settings delegates
check("defaults water", api.default_settings("size_water") == {"recirc_delta_t_f": 10.0})
check("defaults duct scope", api.default_settings("duct_performance", scope="exhaust") == {"scope_seen": "exhaust"})
check("defaults gas", api.default_settings("size_gas") == {"table_id": None})
try:
    api.default_settings("size_duct")
    check("unknown op raises", False)
except ValueError:
    check("unknown op raises", True)

# write flags default off and arguments pass through
r = api.size_water("DOC", 5)
check("size_water ok", r["status"] == "ok")
args, kw = CALLS["water_run"]
check("size_water write off", kw["write"] is False)
check("size_water args", args == ("DOC", 5) and kw["start"] == "rpz")
api.water_performance("DOC", 6, start="tee", view_id=9)
args, kw = CALLS["water_perf_run"]
check("water_perf sheets off", kw["make_sheets"] is False)
check("water_perf pass", kw["start"] == "tee" and kw["view_id"] == 9)
api.diagnose("DOC", 7, include_report=True)
check("diagnose pass", CALLS["diagnose_run"][1]["include_report"] is True)

# gas table_id required and checked
r = api.size_gas("DOC", 8)
check("gas no table refused", r["status"] == "refused")
check("gas lists ids once", r["data"]["table_ids"] == ["402.4(2)", "402.4(5)"])
r = api.size_gas("DOC", 8, table_id="bogus")
check("gas bad table refused", r["status"] == "refused")
CALLS.pop("gas_run", None)
r = api.size_gas("DOC", 8, table_id="402.4(5)", settings={"heat_content_btu_per_cf": 1000.0})
args, kw = CALLS["gas_run"]
check("gas ok", r["status"] == "ok")
check("gas table merged", kw["settings"]["table_id"] == "402.4(5)" and kw["settings"]["heat_content_btu_per_cf"] == 1000.0)
check("gas write off", kw["write"] is False)

# engine check: CPython here runs; faked IronPython is refused before the call
CALLS.pop("duct_perf_run", None)
r = api.duct_performance("DOC", 10, scope="vav", fan_cfm=300)
check("duct runs on cpython", r["status"] == "ok")
check("duct pass", CALLS["duct_perf_run"][1]["scope"] == "vav" and CALLS["duct_perf_run"][1]["fan_cfm"] == 300)
check("duct sheets off", CALLS["duct_perf_run"][1]["make_sheets"] is False)
real_platform = sys.platform
CALLS.pop("duct_perf_run", None)
try:
    sys.platform = "cli"
    r = api.duct_performance("DOC", 10)
finally:
    sys.platform = real_platform
check("duct refused on ironpython", r["status"] == "refused" and "cpython" in r["errors"][0])
check("duct not called on ironpython", "duct_perf_run" not in CALLS)

# a failure inside a tool comes back as an error envelope
def boom(*a, **k):
    raise RuntimeError("revit said no")
sys.modules["diagnose_run"].run_headless = boom
r = api.diagnose("DOC", 1)
check("error envelope", r["status"] == "error" and "revit said no" in r["errors"][0])
check("envelope keys", set(r) >= {"tool", "status", "data", "warnings", "errors", "created"})
del sys.modules["water_run"]
sys.modules["water_run"] = None  # import of water_run now fails
r = api.size_water("DOC", 5)
check("import failure envelope", r["status"] == "error")

print("%d checks, %d failed" % (COUNT[0], len(FAILS)))
sys.exit(1 if FAILS else 0)
