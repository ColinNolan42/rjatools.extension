# Plain CPython tests for agent_api.py with fake tool modules.
# Run: python _test_agent_api.py
import io
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
fake("duct_depth_run", "duct_size_in_depth")

# tools() matches each module's TOOL, no module key leaks
listed = api.tools()
check("tools count", len(listed) == 6)
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

# gas: table_id from the argument or settings, refused only when both differ
CALLS.pop("gas_run", None)
r = api.size_gas("DOC", 8, table_id="402.4(5)", settings={"heat_content_btu_per_cf": 1000.0})
args, kw = CALLS["gas_run"]
check("gas ok", r["status"] == "ok")
check("gas table merged", kw["settings"]["table_id"] == "402.4(5)" and kw["settings"]["heat_content_btu_per_cf"] == 1000.0)
check("gas write off", kw["write"] is False)
CALLS.pop("gas_run", None)
api.size_gas("DOC", 8, settings={"table_id": "402.4(2)"})
check("gas table in settings", CALLS["gas_run"][1]["settings"]["table_id"] == "402.4(2)")
CALLS.pop("gas_run", None)
api.size_gas("DOC", 8)
check("gas no table left to gas_run", CALLS["gas_run"][1]["settings"].get("table_id") is None)
CALLS.pop("gas_run", None)
r = api.size_gas("DOC", 8, table_id="402.4(5)", settings={"table_id": "402.4(2)"})
check("gas table twice refused", r["status"] == "refused" and "gas_run" not in CALLS)

# duct passes straight through (duct_perf_run itself refuses IronPython)
CALLS.pop("duct_perf_run", None)
r = api.duct_performance("DOC", 10, scope="vav", fan_cfm=300)
check("duct ok", r["status"] == "ok")
check("duct pass", CALLS["duct_perf_run"][1]["scope"] == "vav" and CALLS["duct_perf_run"][1]["fan_cfm"] == 300)
check("duct sheets off", CALLS["duct_perf_run"][1]["make_sheets"] is False)

# duct space check passes rows and settings through, doc ignored
CALLS.pop("duct_depth_run", None)
rows = [{"id": "S1", "cfm": 1000}]
r = api.duct_size_in_depth("DOC", rows=rows, settings={"hanger_in": 4})
args, kw = CALLS["duct_depth_run"]
check("depth ok", r["status"] == "ok")
check("depth pass", kw["rows"] == rows and kw["settings"] == {"hanger_in": 4})
check("depth engine both", api._tool("duct_size_in_depth")["engine"] == "both")

# decisions are structured: qid plus how to answer it
for row in listed:
    for d in row["decisions"]:
        check("decision shape " + row["operation"], set(d) == {"qid", "answer_with"} and d["answer_with"] in ("decisions", "fan_cfm", "question to Colin"))

# a failure inside a tool comes back as an error envelope
def boom(*a, **k):
    raise RuntimeError("revit said no")
sys.modules["diagnose_run"].run_headless = boom
r = api.diagnose("DOC", 1)
check("error envelope", r["status"] == "error" and "revit said no" in r["errors"][0])
check("error traceback", "traceback" in r["data"] and any("boom" in line or "revit said no" in line for line in r["data"]["traceback"]))
def ask(*a, **k):
    raise headless.NeedsDecision("incomplete_system", "System is incomplete. Continue?", "stop")
sys.modules["diagnose_run"].run_headless = ask
r = api.diagnose("DOC", 1)
check("escaped question", r["status"] == "needs_decision" and r["decision"]["qid"] == "incomplete_system")
check("envelope keys", set(r) >= {"tool", "status", "data", "warnings", "errors", "created"})
del sys.modules["water_run"]
sys.modules["water_run"] = None  # import of water_run now fails
r = api.size_water("DOC", 5)
check("import failure envelope", r["status"] == "error")


# drift: the hand written TOOLS table against the real modules' source
import ast
HERE = os.path.dirname(os.path.abspath(__file__))
def consts(module):
    tree = ast.parse(io.open(os.path.join(HERE, module + ".py"), encoding="utf-8").read())
    out = {}
    for node in tree.body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            try:
                out[node.targets[0].id] = ast.literal_eval(node.value)
            except ValueError:
                out[node.targets[0].id] = node.value
    return out
gas_c = consts("gas_run")
for t in api.TOOLS:
    c = consts(t["module"])
    check("real TOOL " + t["operation"], c.get("TOOL") == t["operation"])
    qids = c.get("QIDS")
    if t["module"] == "gas_run":
        qids = (gas_c["QID_WARNINGS"],)
    check("real QIDS " + t["operation"], tuple(d["qid"] for d in t["decisions"]) == tuple(qids))
duct_c = consts("duct_perf_run")
check("real duct engine", duct_c["ENGINE"] == api._tool("duct_performance")["engine"])
depth_c = consts("duct_depth_run")
check("real depth engine", depth_c["ENGINE"] == api._tool("duct_size_in_depth")["engine"])
for sc in duct_c["SCOPES"]:
    check("scope listed " + sc, sc in api._tool("duct_performance")["ids"])
duct_src = io.open(os.path.join(HERE, "duct_perf_run.py"), encoding="utf-8").read()
check("duct refuses ironpython", "if sys.platform == 'cli':\n        return done('refused'" in duct_src)
for mod, qids in (("water_run", ("incomplete_system", "no_cold_connector")), ("water_perf_run", ("no_cold_connector",))):
    src = io.open(os.path.join(HERE, mod + ".py"), encoding="utf-8").read()
    for q in qids:
        check("qid raised in %s %s" % (mod, q), '"%s"' % q in src.split("QIDS =", 1)[1])

print("%d checks, %d failed" % (COUNT[0], len(FAILS)))
sys.exit(1 if FAILS else 0)
