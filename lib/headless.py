# headless.py
# Pieces every tool shares when it runs with no dialog, no pick and no alert
# (pyRevit 7 agent runtime, run_query / run_modify). Pure Python, no Revit
# imports, so it is tested in plain CPython (_test_headless.py).
#
# A tool's run_headless(doc, settings, start_id, output=None, prompter=None)
# writes its report to `output` (RecordingOutput when headless, the pyRevit
# output window behind the button) and asks yes/no questions through
# `prompter` (HeadlessPrompter when headless, UiPrompter behind the button).
# It returns result(...), a JSON safe envelope.
#
# IronPython 2.7 and CPython 3.

_NET_INTS = ("Int16", "Int32", "Int64", "UInt16", "UInt32", "UInt64", "Byte")
_NET_FLOATS = ("Double", "Single", "Decimal")

STATUSES = ("ok", "cancelled", "refused", "needs_decision", "error")
DECISIONS = ("continue", "stop")

try:
    _STR_TYPES = (str, unicode)  # noqa: F821  (IronPython 2.7)
except NameError:
    _STR_TYPES = (str,)

try:
    _INT_TYPES = (int, long)  # noqa: F821
except NameError:
    _INT_TYPES = (int,)


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

class RecordingOutput(object):
    """Stands in for the pyRevit output window. Every call is kept in `log`
    as (kind, payload); `forward` (the real window) also gets it when given."""

    def __init__(self, forward=None):
        self.forward = forward
        self.log = []

    def _keep(self, kind, payload, name, args, kwargs):
        self.log.append((kind, payload))
        if self.forward is not None:
            fn = getattr(self.forward, name, None)
            if fn is not None:
                fn(*args, **kwargs)

    def print_md(self, text, *args, **kwargs):
        self._keep("md", text, "print_md", (text,) + args, kwargs)

    def print_html(self, html, *args, **kwargs):
        self._keep("html", html, "print_html", (html,) + args, kwargs)

    def print_code(self, text, *args, **kwargs):
        self._keep("code", text, "print_code", (text,) + args, kwargs)

    def print_table(self, table_data, title="", columns=None, *args, **kwargs):
        payload = {"title": title, "columns": list(columns or []),
                   "rows": [list(r) for r in (table_data or [])]}
        self._keep("table", payload, "print_table",
                   (table_data, title, columns) + args, kwargs)

    def __getattr__(self, name):
        # Window calls with nothing to record (set_title, resize, close,
        # linkify, ...) are forwarded when possible and otherwise ignored.
        if name.startswith("__"):
            raise AttributeError(name)
        forward = self.__dict__.get("forward")
        target = getattr(forward, name, None) if forward is not None else None
        if callable(target):
            return target
        def _ignored(*args, **kwargs):
            return None
        return _ignored

    def text(self):
        """The recorded markdown, code and table titles as one string."""
        parts = []
        for kind, payload in self.log:
            if kind == "table":
                parts.append(payload.get("title") or "")
            else:
                parts.append(payload if isinstance(payload, _STR_TYPES) else repr(payload))
        return "\n".join(p for p in parts if p)


# ---------------------------------------------------------------------------
# Yes/no questions and info alerts
# ---------------------------------------------------------------------------

class NeedsDecision(Exception):
    """Raised by HeadlessPrompter.confirm when no answer was given for qid.
    The tool stops before any write and returns status needs_decision."""

    def __init__(self, qid, text="", suggested="continue"):
        Exception.__init__(self, qid)
        self.qid = qid
        self.text = text
        self.suggested = suggested


class UiPrompter(object):
    """Behind the buttons: the same forms.alert calls as before."""

    def __init__(self, forms):
        self.forms = forms

    def confirm(self, qid, text, title="", suggested="continue"):
        return bool(self.forms.alert(text, title=title, yes=True, no=True))

    def notify(self, text, title="", **kwargs):
        self.forms.alert(text, title=title, **kwargs)


class HeadlessPrompter(object):
    """Headless: info alerts become warnings, a yes/no reads decisions[qid]
    ("continue" or "stop") and raises NeedsDecision when there is none."""

    def __init__(self, decisions=None):
        self.decisions = dict(decisions or {})
        bad = [k for k, v in self.decisions.items() if v not in DECISIONS]
        if bad:
            raise ValueError("decisions must be 'continue' or 'stop': {}".format(
                ", ".join(sorted(bad))))
        self.warnings = []
        self.asked = []

    def confirm(self, qid, text, title="", suggested="continue"):
        self.asked.append(qid)
        if qid not in self.decisions:
            raise NeedsDecision(qid, text, suggested)
        return self.decisions[qid] == "continue"

    def notify(self, text, title="", **kwargs):
        self.warnings.append(text)


# ---------------------------------------------------------------------------
# Result envelope
# ---------------------------------------------------------------------------

def result(tool, status, data=None, warnings=None, errors=None, created=None,
           log=None, decision=None):
    """The one envelope every headless tool returns.

    created: volatile ids and names (views, sheets, drafting views) kept apart
    so two runs compare on `data` alone. decision: {qid, text, suggested}
    when status is needs_decision."""
    if status not in STATUSES:
        raise ValueError("unknown status: {}".format(status))
    out = {
        "tool": tool,
        "status": status,
        "data": jsonable(data if data is not None else {}),
        "warnings": jsonable(list(warnings or [])),
        "errors": jsonable(list(errors or [])),
        "created": jsonable(created if created is not None else {}),
    }
    if decision is not None:
        out["decision"] = jsonable(decision)
    if log is not None:
        out["log"] = jsonable(log)
    return out


def needs_decision(tool, exc, warnings=None, log=None):
    """Envelope for a NeedsDecision raised during a run."""
    return result(tool, "needs_decision", warnings=warnings, log=log,
                  decision={"qid": exc.qid, "text": exc.text,
                            "suggested": exc.suggested})


def jsonable(obj, _depth=0):
    """Plain JSON types only: dict keys as text, tuples and sets as lists,
    Revit ElementId as its integer, anything else as its text."""
    if _depth > 50:
        return repr(obj)
    if obj is None or isinstance(obj, bool):
        return obj
    if isinstance(obj, _INT_TYPES):
        return int(obj)
    if isinstance(obj, float):
        if obj != obj or obj in (float("inf"), float("-inf")):
            return None
        return obj
    if isinstance(obj, _STR_TYPES):
        return obj
    # .NET numbers (ElementId.Value is System.Int64 in IronPython 2.7, which
    # is neither int nor long there).
    tname = type(obj).__name__
    if tname in _NET_INTS:
        return int(obj)
    if tname in _NET_FLOATS:
        return jsonable(float(obj), _depth + 1)
    if isinstance(obj, dict):
        return dict((_key(k), jsonable(v, _depth + 1)) for k, v in obj.items())
    if isinstance(obj, (list, tuple, set, frozenset)):
        items = sorted(obj, key=repr) if isinstance(obj, (set, frozenset)) else obj
        return [jsonable(v, _depth + 1) for v in items]
    for attr in ("Value", "IntegerValue"):  # Revit ElementId (2024+, older)
        value = getattr(obj, attr, None)
        if isinstance(value, _INT_TYPES) and type(obj).__name__ == "ElementId":
            return int(value)
    if hasattr(obj, "__dict__") and not callable(obj):
        return dict((_key(k), jsonable(v, _depth + 1))
                    for k, v in vars(obj).items() if not k.startswith("_"))
    try:
        return str(obj)
    except Exception:
        return repr(obj)


def _key(k):
    if isinstance(k, _STR_TYPES):
        return k
    return str(jsonable(k))
