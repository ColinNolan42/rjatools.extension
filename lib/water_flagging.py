# -*- coding: ascii -*-
# water_flagging.py
# The "pipes needing attention" list for Water Performance, laid out like the
# Duct Performance table (Colin, 2026-10-26).
#
#   * Every red, yellow or purple pipe is its own numbered row, with the drawn
#     size, the size it needs, the flow and the speed.
#   * A run of gray pipes is ONE row: the first pipe of the run, upstream-most,
#     with a count of the pipes behind it. Those depend on fixing that first
#     pipe, so they are not listed. This keeps the list short and stops every
#     elbow and transition behind it from repeating the same finding.
#   * The row number is the number printed in the circle on the plan.
#
# Pure logic, no Revit. IronPython 2.7

import water_graph
import water_performance
import water_tables

COLUMNS = [("#", 0.045), ("STATUS", 0.080), ("REASON", 0.360),
           ("SYSTEM", 0.075), ("SIZE", 0.075), ("REQUIRED SIZE", 0.105),
           ("GPM", 0.060), ("FPS", 0.055)]

CONSOLE_COLUMNS = ["#", "Status", "Reason", "System", "Size", "Required size",
                   "GPM", "FPS"]

_ORDER = (water_performance.RED, water_performance.YELLOW,
          water_performance.PURPLE, water_performance.GRAY)


class Item(object):
    """One row of the table and one circle on the plan."""

    def __init__(self, number, result, dependents, steps=()):
        self.number = number
        self.result = result
        self.dependents = dependents      # pipes behind this one, same status
        self.steps = list(steps)          # (pipe id, size) where the size changes


def _upstream_pipe(graph, pipe_id, result):
    """The nearest pipe above pipe_id on its own tree, or None."""
    node = graph.nodes.get(pipe_id)
    if node is None:
        return None
    key = water_graph.TREE_RETURN if result.is_return else node.system
    seen = set()
    current = node.parent_by_system.get(key)
    while current is not None and current not in seen:
        seen.add(current)
        other = graph.nodes.get(current)
        if other is None:
            return None
        if other.kind == water_graph.KIND_PIPE:
            return current
        current = other.parent_by_system.get(key)
    return None


def run_groups(graph, results):
    """Group every problem pipe (anything not green) into runs.

    A run is a stretch of connected pipes with the same status. Only its first
    pipe, upstream-most, is reported. The pipes behind it are dependent on
    fixing that one.

    Returns (roots, dependent_ids, depth): roots maps the first pipe of each
    run to the ids of the pipes behind it, nearest first; dependent_ids is the
    set of those pipes; depth maps a pipe to how many pipes it sits below the
    start of its run.
    """
    flagged = set(pid for pid, r in results.items() if r.label in _ORDER)
    above = {}
    for pid in flagged:
        up = _upstream_pipe(graph, pid, results[pid])
        same = (up in flagged and results[up].label == results[pid].label)
        above[pid] = up if same else None

    depth = {}

    def root_of(pid):
        hops = 0
        seen = set()
        while above.get(pid) is not None and pid not in seen:
            seen.add(pid)
            pid = above[pid]
            hops += 1
        return pid, hops

    roots = {}
    dependents = set()
    for pid in flagged:
        root, hops = root_of(pid)
        depth[pid] = hops
        roots.setdefault(root, [])
        if root != pid:
            roots[root].append(pid)
            dependents.add(pid)
    for root in roots:
        roots[root].sort(key=lambda p: (depth[p], p))
    return roots, dependents, depth, above


def size_steps(results, root_id, dependents, above):
    """Where the size the run needs changes, as [(pipe id, required size)].

    These are the pipes behind the first one that would be upsized or
    downsized differently from the pipe above them.
    """
    steps = []
    for pid in dependents:
        up = above.get(pid)
        if up is None:
            continue
        mine = results[pid].required_nominal
        theirs = results[up].required_nominal
        if mine and theirs and mine != theirs:
            steps.append((pid, mine))
    return steps


def flag_items(graph, results):
    """Numbered table items, worst first, one per run of problem pipes."""
    roots, dependents, _depth, above = run_groups(graph, results)
    flagged = [r for r in results.values()
               if r.label in _ORDER and r.element_id not in dependents]
    flagged.sort(key=lambda r: (-water_performance.PRIORITY[r.label],
                                r.element_id))
    items = []
    for number, r in enumerate(flagged, 1):
        behind = roots.get(r.element_id, [])
        steps = size_steps(results, r.element_id, behind, above)
        steps.sort(key=lambda t: (_depth[t[0]], t[0]))
        items.append(Item(number, r, behind, steps))
    return items


def short_reason(result, dependents, steps=()):
    """The Reason cell, short, in the style of the Duct Performance table."""
    label = result.label
    if label == water_performance.RED and getattr(result, "over_table",
                                                  False):
        text = "Load above the firm table, size by hand"
    elif label == water_performance.RED:
        text = "Undersized"
    elif label == water_performance.PURPLE:
        text = "Oversized"
    elif label == water_performance.YELLOW:
        text = "Velocity over limit"
    else:
        text = _gray_reason(result.reason or "")
    if getattr(result, "existing", False):
        # An existing pipe is reported, never resized.
        text = "Existing pipe, " + text[:1].lower() + text[1:]
    if dependents:
        text += ", {} more pipe(s) behind it depend on this".format(
            len(dependents))
    if steps:
        shown = "; ".join("{} in at pipe {}".format(size, pid)
                          for pid, size in steps[:3])
        if len(steps) > 3:
            shown += "; +{} more".format(len(steps) - 3)
        text += ". Size changes: " + shown
    return text


def _gray_reason(reason):
    lowered = reason.lower()
    if lowered.startswith("no fixture demand"):
        return "No demand downstream"
    if "serves no hot fixtures" in lowered:
        return "Heater has no load"
    if "could not find where" in lowered:
        return "Return end not found"
    if "above the firm table" in lowered:
        return "Load above the table"
    if "drawn size not read" in lowered:
        return "Size not read"
    return reason[:1].upper() + reason[1:] if reason else "Not judged"


def _size_text(inches):
    if inches is None:
        return "-"
    return "{} in".format(
        "{:.3f}".format(float(inches)).rstrip("0").rstrip("."))


def row_cells(item):
    """The cells of one table row, in COLUMNS order."""
    r = item.result
    system = ("Return" if r.is_return else
              ("Hot" if r.system == water_tables.SYSTEM_HOT else "Cold"))
    required = "-"
    if r.label != water_performance.GRAY and r.required_nominal:
        required = "{} in".format(r.required_nominal)
    return [
        str(item.number),
        r.label,
        short_reason(r, item.dependents, item.steps),
        system,
        _size_text(r.drawn_in),
        required,
        "{:.1f}".format(r.gpm) if r.gpm is not None else "-",
        "{:.1f}".format(r.fps) if r.fps is not None else "-",
    ]
