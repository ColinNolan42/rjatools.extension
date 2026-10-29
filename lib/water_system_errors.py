# -*- coding: ascii -*-
# water_system_errors.py
# The SYSTEM ERRORS list for Water Performance (Colin, 2026-10-26).
#
# An estimate of how complete the model is: the places where the model is not
# finished or is wrong in a way that can make a calculation wrong, gathered
# in one list under the result counts. A missing recirculation pump, a
# disconnected pipe, hot tied into cold, no circuit setters, a heater with no
# load, a fixture with no load, and so on.
#
# Pure logic, no Revit. IronPython 2.7

import water_checks
import water_graph
import water_tables

ERROR = water_checks.SEVERITY_ERROR
WARNING = water_checks.SEVERITY_WARNING


def collect(graph, sizing, recirc, checks, items=None, pressure=None,
            pumps=None):
    """Return a list of (severity, text), errors first.

    Args:
        graph: the traversed WaterGraph.
        sizing: what water_sizing_engine.size_network returned.
        recirc: what water_recirc.analyze returned, or None.
        checks: what water_checks.check_system returned.
        items: water_flagging items, for the gray runs.
        pressure: water_pressure.analyze result, or None.
        pumps: water_pump.duty_points result, or None.
    """
    found = []

    for finding in checks["findings"]:
        if finding.severity in (ERROR, WARNING):
            found.append((finding.severity,
                          finding.short or _first_sentence(finding.message)))

    found.extend(_circuit_setters(graph, recirc))

    unknown = [n for n in (graph.nodes.get(i) for i in
                           getattr(graph, "unrecognized_ids", []))
               if n is not None]
    if unknown:
        names = ", ".join("{} ({})".format(n.element_id, n.family_name)
                          for n in unknown[:5])
        if len(unknown) > 5:
            names += ", +{} more".format(len(unknown) - 5)
        found.append((WARNING, "{} element(s) on the water piping are not "
                      "recognized as a fixture, heater, pump or device, so "
                      "their load is not counted: {}.".format(
                          len(unknown), names)))

    over = [s.element_id for s in sizing["segments"] if s.exceeds_table]
    if over:
        top = water_tables.firm_sizing_rows()[-1]
        found.append((ERROR, "{} pipe(s) carry more load than the firm table "
                      "covers ({} in maximum: {:.0f} cold WSFU, {:.0f} hot "
                      "WSFU). The tool is not built for this load; size "
                      "them by hand: {}.".format(
                          len(over), top["nominal_size"],
                          float(top["max_wsfu_8fps_cold"]),
                          float(top["max_wsfu_5fps_hot"]), _ids(over))))

    upsize = [s.element_id for s in sizing["segments"]
              if getattr(s, "upsize_needed", False)]
    if upsize:
        found.append((ERROR, "{} existing pipe(s) are too small for the new "
                      "load: {}.".format(len(upsize), _ids(upsize))))

    if recirc and recirc.get("unrooted_pipe_ids"):
        found.append((ERROR, "{} return pipe(s) are not traced to the "
                      "heater: {}.".format(len(recirc["unrooted_pipe_ids"]),
                                           _ids(recirc["unrooted_pipe_ids"]))))

    if pressure:
        for warning in pressure.get("warnings", []):
            found.append((WARNING, warning.strip()))
        notes = sorted(set(n for o in pressure["outlets"] for n in o.notes
                           if "Device_Pressure_Drop_PSI" in n))
        for note in notes:
            found.append((WARNING, note.rstrip(".") + "."))

    if pumps:
        for pump in pumps:
            for circuit in pump["circuits"]:
                if any("supply path" in n for n in circuit["notes"]):
                    found.append((WARNING, "The pump head leaves out the "
                                  "supply side. The supply path to the "
                                  "heater was not found."))
                    break

    gray_runs = [i for i in (items or []) if i.result.label == "GRAY"]
    if gray_runs:
        pipes = sum(1 + len(i.dependents) for i in gray_runs)
        found.append((WARNING, "{} run(s) of piping, {} pipe(s) in all, could "
                      "not be judged. See the gray rows in the table.".format(
                          len(gray_runs), pipes)))

    unique = []
    for entry in found:
        if entry not in unique:
            unique.append(entry)
    unique.sort(key=lambda e: 0 if e[0] == ERROR else 1)
    return unique


def _circuit_setters(graph, recirc):
    if not recirc or not recirc.get("legs"):
        return []
    legs = len(recirc["legs"])
    setters = [
        n for n in graph.nodes.values()
        if n.kind == water_graph.KIND_ACCESSORY and
        "circuit setter" in "{} {}".format(
            n.family_name or "", getattr(n, "type_name", "") or "").lower()]
    if not setters:
        return [(ERROR, "No circuit setters are modeled on the return "
                 "piping. One per return leg is assumed ({} leg(s)).".format(
                     legs))]
    if len(setters) < legs:
        return [(WARNING, "{} circuit setter(s) are modeled for {} return "
                 "leg(s). One per leg is assumed.".format(len(setters), legs))]
    return []


def summary_line(entries):
    """The line printed above the list."""
    if not entries:
        return "No system errors found."
    errors = len([e for e in entries if e[0] == ERROR])
    warnings = len(entries) - errors
    return ("{} error(s) and {} warning(s) in the model can make the results "
            "above wrong.".format(errors, warnings))


def lines(entries):
    """Display lines: the summary, then one line per entry."""
    out = [summary_line(entries)]
    for severity, text in entries:
        out.append("{}: {}".format(severity, text))
    return out


def _first_sentence(message):
    text = (message or "").strip()
    cut = text.find(". ")
    return text if cut < 0 else text[:cut + 1]


def _ids(values, show=5):
    values = list(values)
    text = ", ".join(str(v) for v in values[:show])
    if len(values) > show:
        text += ", +{} more".format(len(values) - show)
    return text
