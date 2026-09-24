# -*- coding: ascii -*-
# water_sizing_engine.py
# Domestic water pipe sizing from water supply fixture units.
#
# Takes a traversed WaterGraph and picks a nominal size for every pipe using
# the firm WSFU-per-size table (see water_tables.py). Pure decision logic plus
# one optional Revit write phase, so the size selection can be unit tested
# without Revit.
#
# Basis of design: 2024 IPC. See water_tables.basis_of_design_lines().
#
# Accounting rule (Colin, 2026-09-24): cold water piping is sized on TOTAL
# fixture units, hot water piping on HOT fixture units. water_graph assigns the
# per-pipe demand; this module only converts demand to a size.
#
# IronPython 2.7

import shared_params
import water_tables
import water_graph


# Minimum pipe size serving two or more fixtures, per the firm design standard
# ("pipe to a single fixture per the RJA plumbing fixture schedule, pipe to two
# or more fixtures not smaller than 3/4 inch"). Applied unless disabled.
MIN_MULTI_FIXTURE_SIZE_IN = 0.75


class SegmentResult(object):
    """The sizing outcome for one pipe."""

    def __init__(self, element_id):
        self.element_id      = element_id
        self.system          = None
        self.demand_wsfu     = 0.0
        self.demand_gpm      = None
        self.existing_size_in = None
        self.nominal_size    = None
        self.size_inches     = None
        self.rule            = ""
        self.flag            = ""
        self.fixture_count   = 0
        self.length_feet     = 0.0


def size_network(graph, apply_minimums=True, return_system_type_ids=None):
    """Choose a nominal size for every pipe in the graph.

    Args:
        graph: a traversed WaterGraph.
        apply_minimums: enforce the firm minimum sizes (single fixture takes
            its connector size, two or more fixtures take 3/4 inch minimum).
        return_system_type_ids: PipingSystemType element ids the user marked as
            recirculation / return. Those pipes are reported but NOT sized,
            because return piping is sized on circulation flow, not on fixture
            units. Revit classifies a recirc system as Domestic Hot Water, the
            same as the supply, so this cannot be detected automatically.

    Returns:
        dict with "segments" (list of SegmentResult, sorted by demand
        descending), "flags" (list of strings) and "totals".
    """
    segments = []
    flags = []
    returns = set(return_system_type_ids or [])

    for node in graph.nodes.values():
        if node.kind != water_graph.KIND_PIPE:
            continue

        result = SegmentResult(node.element_id)
        result.system = node.system
        result.demand_wsfu = node.demand_wsfu
        result.existing_size_in = node.diameter_inches
        result.length_feet = node.length_feet

        if node.system_type_id is not None and node.system_type_id in returns:
            result.rule = "NOT SIZED"
            result.flag = ("recirculation / return system '{}' - sized on "
                           "circulation flow, not fixture units, and not part "
                           "of this build".format(node.system_type_name))
            segments.append(result)
            continue

        system_key = (water_tables.SYSTEM_HOT
                      if node.system == shared_params.SYSTEM_DOMESTIC_HOT_WATER
                      else water_tables.SYSTEM_COLD)

        # Demand in gpm, for the report only. The size comes from WSFU.
        demand = water_tables.wsfu_to_gpm(node.demand_wsfu)
        result.demand_gpm = demand["gpm"]

        selection = water_tables.select_size(node.demand_wsfu, system_key)
        result.nominal_size = selection["nominal_size"]
        result.size_inches = selection["size_inches"]

        if selection["status"] == "exceeds_table":
            result.rule = "NOT SIZED"
            result.flag = ("demand {:.1f} wsfu exceeds the firm table maximum "
                           "of {:.0f} wsfu".format(
                               float(node.demand_wsfu),
                               float(selection["max_table_wsfu"])))
            flags.append("Pipe {}: {}".format(node.element_id, result.flag))
            segments.append(result)
            continue

        if selection["status"] == "no_demand":
            result.rule = "NOT SIZED"
            result.flag = ("no fixture demand downstream - pipe is past the "
                           "last fixture, or its fixtures carry no WSFU")
            flags.append("Pipe {}: {}".format(node.element_id, result.flag))
            segments.append(result)
            continue

        result.rule = "{} wsfu, {} table limit {:.0f}".format(
            _fmt(node.demand_wsfu), system_key, float(selection["limit_wsfu"]))

        fixture_ids = _served_fixture_ids(graph, node)
        result.fixture_count = len(fixture_ids)

        if apply_minimums:
            result = _apply_minimums(graph, node, result, fixture_ids)

        segments.append(result)

    segments.sort(key=lambda s: s.demand_wsfu, reverse=True)

    sized = [s for s in segments if s.nominal_size is not None]
    totals = {
        "pipe_count":   len(segments),
        "sized_count":  len(sized),
        "unsized_count": len(segments) - len(sized),
        "total_length_ft": sum(s.length_feet for s in segments),
    }
    return {"segments": segments, "flags": flags, "totals": totals}


def _apply_minimums(graph, node, result, fixture_ids):
    """Raise the selected size to the firm minimum where one applies.

    Two rules, per the firm design standard:
      - A pipe serving exactly one fixture is not smaller than that fixture's
        own connector size (from the family's supply size parameter).
      - A pipe serving two or more fixtures is not smaller than 3/4 inch.
    A pipe is never reduced by this step.
    """
    if result.size_inches is None:
        return result

    minimum_in = None
    reason = ""

    if len(fixture_ids) == 1:
        fixture = graph.nodes.get(fixture_ids[0])
        connector_in = _fixture_supply_size_in(fixture, node.system)
        if connector_in:
            minimum_in = connector_in
            reason = "fixture connector size"
    elif len(fixture_ids) > 1:
        minimum_in = MIN_MULTI_FIXTURE_SIZE_IN
        reason = "two or more fixtures, 3/4 in minimum"

    if minimum_in is None or result.size_inches >= minimum_in:
        return result

    for row in water_tables.firm_sizing_rows():
        if row["size_inches"] >= minimum_in:
            result.nominal_size = row["nominal_size"]
            result.size_inches = row["size_inches"]
            result.rule += "; raised to {} ({})".format(
                row["nominal_size"], reason)
            return result

    return result


def _fixture_supply_size_in(fixture, system):
    """The fixture's own supply connector size in inches, or None."""
    if fixture is None:
        return None
    for connector in fixture.connectors:
        if connector.get("system_type") != system:
            continue
        diameter = connector.get("diameter_in")
        if diameter:
            return diameter
    return None


def _served_fixture_ids(graph, node):
    """Which fixtures this pipe serves, matching water_graph's demand rule."""
    served = []
    seen = set()
    for nid in graph.descendants(node.element_id, node.system):
        child = graph.nodes[nid]
        if child.kind == water_graph.KIND_FIXTURE:
            if nid not in seen:
                seen.add(nid)
                served.append(nid)
        elif (child.kind == water_graph.KIND_HEATER and
              node.system == shared_params.SYSTEM_DOMESTIC_COLD_WATER):
            for fid in child.served_fixture_ids:
                if fid not in seen:
                    seen.add(fid)
                    served.append(fid)
    return served


def write_sizes(doc, graph, sizing_result, transaction_factory):
    """Write the selected sizes back to the Revit pipes.

    Kept separate from size_network so sizing can be run and reported without
    touching the model.

    Args:
        doc: the Revit Document.
        graph: the traversed WaterGraph.
        sizing_result: what size_network returned.
        transaction_factory: callable taking (doc, name) and returning a
            Revit Transaction. Injected so this module needs no Revit import.

    Returns:
        dict with "written", "failed" and "errors".
    """
    import revit_helpers

    written = 0
    failed = 0
    errors = []

    transaction = transaction_factory(doc, "RJA Tools - Size Water Pipes")
    transaction.Start()
    try:
        for segment in sizing_result["segments"]:
            if segment.size_inches is None:
                continue
            node = graph.nodes.get(segment.element_id)
            if node is None or node.element is None:
                failed += 1
                errors.append("Pipe {}: element not available.".format(
                    segment.element_id))
                continue
            ok, approach = revit_helpers.set_pipe_diameter(
                node.element, segment.size_inches)
            if ok:
                written += 1
            else:
                failed += 1
                errors.append("Pipe {}: could not set diameter ({}).".format(
                    segment.element_id, approach))
        transaction.Commit()
    except Exception as exc:
        transaction.RollBack()
        raise exc

    return {"written": written, "failed": failed, "errors": errors}


def _fmt(value):
    """Format a WSFU value without a trailing .0 for whole numbers.

    IronPython 2.7 raises ValueError on '{:.0f}'.format(an_int), so the cast
    to float is required, not cosmetic.
    """
    number = float(value)
    if abs(number - round(number)) < 0.01:
        return "{:.0f}".format(number)
    return "{:.2f}".format(number)
