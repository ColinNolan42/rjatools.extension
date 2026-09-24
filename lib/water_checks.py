# -*- coding: ascii -*-
# water_checks.py
# Completeness checks run BEFORE sizing.
#
# A domestic water system is "complete" when the tool can account for every
# system the model is actually asking for:
#
#   RPZ          -> the traversal origin, and the start of the cold water
#   water heater -> proves hot water is served, and is what the hot tree
#                   hangs off. Hot fixtures with no heater cannot be sized.
#   circ pump    -> proves recirculation is really present, so fixtures asking
#                   for HWR are actually on a loop
#
# The point is to fail loudly on a half-modelled system rather than quietly
# sizing part of it. Mirrors the gas Diagnose report's ready_for_sizing flag.
#
# Pure logic. No Revit imports, so it is testable in plain CPython.
#
# IronPython 2.7

import shared_params
import water_graph


SEVERITY_ERROR = "ERROR"
SEVERITY_WARNING = "WARNING"
SEVERITY_INFO = "INFO"


class Finding(object):
    def __init__(self, severity, code, message):
        self.severity = severity
        self.code = code
        self.message = message

    def __repr__(self):
        return "{}: {}".format(self.severity, self.message)


def check_system(graph, return_system_type_ids=None):
    """Run every completeness check against a traversed graph.

    Returns dict with:
        "findings"          list of Finding, errors first
        "errors"            list of Finding with severity ERROR
        "warnings"          list of Finding with severity WARNING
        "ready_for_sizing"  bool, False when any ERROR was raised
        "profile"           which systems the model supports
    """
    findings = []
    returns = set(return_system_type_ids or [])

    profile = _build_profile(graph, returns)

    _check_origin(graph, findings)
    _check_fixtures(graph, findings)
    _check_hot_water(graph, profile, findings)
    _check_recirculation(graph, profile, returns, findings)
    _check_connectivity(graph, findings)

    errors = [f for f in findings if f.severity == SEVERITY_ERROR]
    warnings = [f for f in findings if f.severity == SEVERITY_WARNING]

    order = {SEVERITY_ERROR: 0, SEVERITY_WARNING: 1, SEVERITY_INFO: 2}
    findings.sort(key=lambda f: order.get(f.severity, 3))

    return {
        "findings": findings,
        "errors": errors,
        "warnings": warnings,
        "ready_for_sizing": len(errors) == 0,
        "profile": profile,
    }


# ---------------------------------------------------------------------------
# Profile: what the model actually contains
# ---------------------------------------------------------------------------

def _build_profile(graph, returns):
    fixtures = [graph.nodes[f] for f in graph.fixture_ids if f in graph.nodes]
    hot_fixtures = [f for f in fixtures if f.hw_wsfu > 0]
    hwr_fixtures = [f for f in fixtures if f.hwr_active]

    hot_system_type_ids = set()
    for node in graph.nodes.values():
        if (node.kind == water_graph.KIND_PIPE and
                node.system == shared_params.SYSTEM_DOMESTIC_HOT_WATER and
                node.system_type_id is not None):
            hot_system_type_ids.add(node.system_type_id)

    return {
        "fixture_count": len(fixtures),
        "hot_fixture_count": len(hot_fixtures),
        "hwr_fixture_count": len(hwr_fixtures),
        "heater_count": len(graph.heater_ids),
        "pump_count": len(graph.pump_ids),
        "hot_system_type_ids": hot_system_type_ids,
        "return_system_type_ids": set(returns),
        "has_cold": True,
        "has_hot": len(graph.heater_ids) > 0,
        "has_recirc": len(returns) > 0 and len(graph.pump_ids) > 0,
    }


# ---------------------------------------------------------------------------
# Individual checks
# ---------------------------------------------------------------------------

def _check_origin(graph, findings):
    origin = graph.nodes.get(graph.origin_id)
    if origin is None:
        findings.append(Finding(
            SEVERITY_ERROR, "no_origin",
            "The picked element is not in the graph. Nothing was traversed."))
        return

    cold = [c for c in origin.connectors
            if c.get("system_type") == shared_params.SYSTEM_DOMESTIC_COLD_WATER]

    if not cold:
        findings.append(Finding(
            SEVERITY_ERROR, "origin_not_on_cold_water",
            "The picked element {} ('{}') has no Domestic Cold Water "
            "connector, so it cannot be the RPZ. Pick the backflow preventer "
            "on the cold water main.".format(
                origin.element_id, origin.family_name)))
        return

    if len(cold) < 2:
        findings.append(Finding(
            SEVERITY_WARNING, "origin_not_inline",
            "The picked element {} ('{}') has only one cold water connector. "
            "An RPZ is an inline device with an inlet and an outlet, so this "
            "may be an end-of-line fitting rather than the backflow "
            "preventer.".format(origin.element_id, origin.family_name)))

    connected = [c for c in cold if c.get("is_connected")]
    if not connected:
        findings.append(Finding(
            SEVERITY_ERROR, "origin_disconnected",
            "The picked element {} ('{}') has cold water connectors but none "
            "of them are connected to anything.".format(
                origin.element_id, origin.family_name)))


def _check_fixtures(graph, findings):
    if not graph.fixture_ids:
        findings.append(Finding(
            SEVERITY_ERROR, "no_fixtures",
            "No water fixtures were reached from the RPZ. Either the piping "
            "is not connected through to them, or the fixtures do not carry "
            "the {} parameter.".format(shared_params.PARAM_IS_WATER_FIXTURE)))
        return

    zero_load = []
    for fid in graph.fixture_ids:
        node = graph.nodes.get(fid)
        if node is not None and node.total_wsfu <= 0:
            zero_load.append(node)

    if zero_load:
        names = ", ".join("{} ('{}')".format(n.element_id, n.fixture_name)
                          for n in zero_load[:8])
        findings.append(Finding(
            SEVERITY_ERROR, "fixture_zero_wsfu",
            "{} fixture(s) carry zero total WSFU, which is a data error, not "
            "a zero demand. Check the Type and the public/private setting "
            "on: {}{}".format(
                len(zero_load), names,
                "" if len(zero_load) <= 8 else ", ...")))

    # A failed Type-name read is not cosmetic: the take-off groups rows by
    # Type, so every fixture collapses into a single "UNKNOWN TYPE" row.
    unknown = [graph.nodes[f] for f in graph.fixture_ids
               if f in graph.nodes and
               graph.nodes[f].type_name in (None, "UNKNOWN TYPE")]
    if unknown:
        findings.append(Finding(
            SEVERITY_WARNING, "fixture_type_name_unreadable",
            "{} fixture(s) returned no readable Type name, so they will group "
            "together as 'UNKNOWN TYPE' in the take-off table. The per-segment "
            "pipe sizes are unaffected, they read each fixture's own fixture "
            "units.".format(len(unknown))))


def _check_hot_water(graph, profile, findings):
    heaters = profile["heater_count"]
    hot_fixtures = profile["hot_fixture_count"]

    if hot_fixtures > 0 and heaters == 0:
        findings.append(Finding(
            SEVERITY_ERROR, "hot_demand_no_heater",
            "{} fixture(s) draw hot water but NO water heater was found on "
            "the network. A heater is identified by having both a cold water "
            "inlet and a hot water outlet. Without one, the cold branch that "
            "feeds it cannot be sized and no hot piping can be "
            "sized.".format(hot_fixtures)))
        return

    if heaters == 0:
        findings.append(Finding(
            SEVERITY_INFO, "cold_only_system",
            "No water heater found, and no fixture draws hot water. Sizing "
            "this as a cold water only system."))
        return

    findings.append(Finding(
        SEVERITY_INFO, "hot_water_available",
        "Hot water available: {} heater(s) serving {} hot fixture(s).".format(
            heaters, hot_fixtures)))

    unserved = []
    for hid in graph.heater_ids:
        node = graph.nodes.get(hid)
        if node is not None and not node.served_fixture_ids:
            unserved.append(node)

    if unserved:
        names = ", ".join("{} ('{}')".format(n.element_id, n.family_name)
                          for n in unserved)
        findings.append(Finding(
            SEVERITY_WARNING, "heater_serves_nothing",
            "{} water heater(s) reach no fixtures on their hot outlet, so "
            "the cold branch feeding them will size at zero. Check the hot "
            "piping is connected: {}".format(len(unserved), names)))


def _check_recirculation(graph, profile, returns, findings):
    pumps = profile["pump_count"]
    hwr_fixtures = profile["hwr_fixture_count"]
    hot_types = profile["hot_system_type_ids"]

    pump_names = ", ".join(
        "{} ('{}')".format(graph.nodes[p].element_id, graph.nodes[p].family_name)
        for p in graph.pump_ids if p in graph.nodes)

    if pumps:
        findings.append(Finding(
            SEVERITY_INFO, "recirc_pump_found",
            "Recirculation available: {} pump candidate(s) found on the hot "
            "water side: {}. Confirm these are circulation pumps, they are "
            "identified by category and connectors, not by "
            "name.".format(pumps, pump_names)))

    if hwr_fixtures and not pumps:
        findings.append(Finding(
            SEVERITY_WARNING, "hwr_wanted_no_pump",
            "{} fixture(s) have HWR_ACTIVE set, but no recirculation pump was "
            "found on the hot water network. Either the pump is missing from "
            "the model, or it is not connected, or its connectors could not "
            "be read.".format(hwr_fixtures)))

    if pumps and not hwr_fixtures:
        findings.append(Finding(
            SEVERITY_WARNING, "pump_no_hwr_fixtures",
            "A recirculation pump was found but no fixture has HWR_ACTIVE "
            "set, so nothing is asking to be on the loop. Check 'Add HWR' on "
            "the fixtures that need it."))

    if hwr_fixtures and not returns:
        extra = ""
        if len(hot_types) > 1:
            extra = (" This model has {} hot water system types, so one of "
                     "them is probably the return.".format(len(hot_types)))
        findings.append(Finding(
            SEVERITY_WARNING, "hwr_wanted_no_return_system",
            "{} fixture(s) have HWR_ACTIVE set, but no piping system type was "
            "marked as the RETURN in the dialog, so every hot pipe was sized "
            "as supply.{}".format(hwr_fixtures, extra)))

    if returns and not pumps:
        findings.append(Finding(
            SEVERITY_WARNING, "return_system_no_pump",
            "A return system was marked in the dialog but no recirculation "
            "pump was found on it. A recirculation loop needs a pump, "
            "gravity and thermosiphon circulation are not permitted."))

    if profile["has_recirc"]:
        findings.append(Finding(
            SEVERITY_INFO, "recirc_complete",
            "Recirculation is complete: return piping identified and a pump "
            "found. Return piping is reported but NOT sized, it is sized on "
            "circulation flow rather than fixture units."))


def _check_connectivity(graph, findings):
    if graph.open_ends:
        findings.append(Finding(
            SEVERITY_WARNING, "open_connectors",
            "{} open connector(s) on the network. Anything past an open "
            "connector was never reached and is therefore not counted in any "
            "pipe size.".format(len(graph.open_ends))))

    if graph.unreadable:
        findings.append(Finding(
            SEVERITY_WARNING, "unreadable_connectors",
            "{} element(s) had no readable connectors, so the traversal could "
            "not pass through them. A circulation pump is a known offender "
            "here.".format(len(graph.unreadable))))

    if graph.loops:
        findings.append(Finding(
            SEVERITY_INFO, "loops_present",
            "{} loop connection(s) found. A recirculation loop is expected to "
            "show up here. Any other loop is a modelling error.".format(
                len(graph.loops))))


# ---------------------------------------------------------------------------
# Report block
# ---------------------------------------------------------------------------

def format_checks(result):
    """Render the check result for the pyRevit window."""
    profile = result["profile"]
    lines = ["SYSTEM COMPLETENESS CHECK", "-" * 78]

    lines.append("  Cold water:      YES (from the picked RPZ)")
    lines.append("  Hot water:       {}".format(
        "YES, {} heater(s)".format(profile["heater_count"])
        if profile["has_hot"] else "NO heater found"))
    lines.append("  Recirculation:   {}".format(
        "YES, {} pump(s) and a return system".format(profile["pump_count"])
        if profile["has_recirc"]
        else "not established, see below" if (profile["pump_count"] or
                                              profile["hwr_fixture_count"])
        else "none in this model"))
    lines.append("  Fixtures:        {} total, {} with hot water, {} on "
                 "recirculation".format(profile["fixture_count"],
                                        profile["hot_fixture_count"],
                                        profile["hwr_fixture_count"]))
    lines.append("-" * 78)

    if not result["findings"]:
        lines.append("  Nothing to report.")
    for finding in result["findings"]:
        prefix = "  [{}] ".format(finding.severity)
        for i, chunk in enumerate(_wrap(finding.message, 70)):
            lines.append(prefix + chunk if i == 0
                         else " " * len(prefix) + chunk)

    lines.append("-" * 78)
    lines.append("  READY FOR SIZING: {}".format(
        "YES" if result["ready_for_sizing"] else
        "NO, {} error(s) must be fixed first".format(len(result["errors"]))))
    lines.append("-" * 78)
    return lines


def _wrap(text, width):
    words = str(text).split()
    if not words:
        return [""]
    lines = []
    current = words[0]
    for word in words[1:]:
        if len(current) + 1 + len(word) <= width:
            current = current + " " + word
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines
