# Plain CPython tests for headless.py. Run: python _test_headless.py
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import headless as h

FAILS = []
COUNT = [0]


def check(name, cond):
    COUNT[0] += 1
    if not cond:
        FAILS.append(name)
        print("FAIL " + name)


# RecordingOutput, standalone
out = h.RecordingOutput()
out.print_md("## Title")
out.print_code("x = 1")
out.print_html("<b>hi</b>")
out.print_table([[1, "a"], (2, "b")], title="T", columns=["n", "s"])
out.set_title("ignored")
out.resize_lines(5)
check("log kinds", [k for k, _ in out.log] == ["md", "code", "html", "table"])
check("table payload", out.log[3][1] == {"title": "T", "columns": ["n", "s"],
                                         "rows": [[1, "a"], [2, "b"]]})
check("text", out.text() == "## Title\nx = 1\n<b>hi</b>\nT")


# RecordingOutput, forwarded
class Window(object):
    def __init__(self):
        self.calls = []

    def print_md(self, t):
        self.calls.append(("md", t))

    def print_table(self, table_data, title="", columns=None):
        self.calls.append(("table", title))

    def set_title(self, t):
        self.calls.append(("title", t))


win = Window()
fwd = h.RecordingOutput(forward=win)
fwd.print_md("a")
fwd.print_table([[1]], title="X")
fwd.print_html("no html on window")
fwd.set_title("S")
check("forwarded", win.calls == [("md", "a"), ("table", "X"), ("title", "S")])
check("forward still logs", len(fwd.log) == 3)

# HeadlessPrompter
p = h.HeadlessPrompter({"incomplete_system": "continue", "no_cold": "stop"})
check("continue", p.confirm("incomplete_system", "go on?") is True)
check("stop", p.confirm("no_cold", "go on?") is False)
try:
    p.confirm("fan_cfm", "Use rated cfm?", suggested="continue")
    check("raises NeedsDecision", False)
except h.NeedsDecision as exc:
    check("needs qid", exc.qid == "fan_cfm" and exc.text == "Use rated cfm?")
check("asked", p.asked == ["incomplete_system", "no_cold", "fan_cfm"])
p.notify("info")
check("notify to warnings", p.warnings == ["info"])
try:
    h.HeadlessPrompter({"x": "yes"})
    check("bad decision rejected", False)
except ValueError:
    check("bad decision rejected", True)


# UiPrompter
class Forms(object):
    def __init__(self, answer):
        self.answer = answer
        self.calls = []

    def alert(self, text, title="", **kw):
        self.calls.append((text, title, kw))
        return self.answer


f = Forms(True)
ui = h.UiPrompter(f)
check("ui confirm", ui.confirm("q", "Sure?", title="T") is True)
check("ui yes/no args", f.calls[0][2] == {"yes": True, "no": True})
ui.notify("Done", title="T")
check("ui notify", f.calls[1] == ("Done", "T", {}))
check("ui no", h.UiPrompter(Forms(None)).confirm("q", "x") is False)


# jsonable
class ElementId(object):
    def __init__(self, v):
        self.Value = v


class Plan(object):
    def __init__(self):
        self.element_id = ElementId(42)
        self.size = (12, 8)
        self._private = 1


data = {1: "one", "nan": float("nan"), "ids": {3, 1, 2}, "plan": Plan(),
        "nested": [(1, 2), {"b": True, "n": None}]}
j = h.jsonable(data)
check("int key", j["1"] == "one")
check("nan to None", j["nan"] is None)
check("set sorted", j["ids"] == [1, 2, 3])
check("object vars", j["plan"] == {"element_id": 42, "size": [12, 8]})
check("nested", j["nested"] == [[1, 2], {"b": True, "n": None}])
check("dumps", isinstance(json.dumps(j), str))

# result envelope
r = h.result("size_water", "ok", data={"pipes": 3}, warnings=["w"],
             created={"sheet": ElementId(9)})
check("keys", sorted(r) == ["created", "data", "errors", "status", "tool", "warnings"])
check("created id", r["created"] == {"sheet": 9})
check("no log by default", "log" not in r)
try:
    h.result("x", "done")
    check("bad status rejected", False)
except ValueError:
    check("bad status rejected", True)
nd = h.needs_decision("size_water", h.NeedsDecision("incomplete_system", "Open ends", "stop"))
check("needs_decision", nd["status"] == "needs_decision"
      and nd["decision"] == {"qid": "incomplete_system", "text": "Open ends",
                             "suggested": "stop"})
check("envelope dumps", isinstance(json.dumps(nd), str))

print("{} checks, {} failed".format(COUNT[0], len(FAILS)))
sys.exit(1 if FAILS else 0)
