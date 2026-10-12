# -*- coding: ascii -*-
# water_recirc.py
# Hot water recirculation: circuit setter flows and the flow in every return
# pipe.
#
# METHOD (Colin, 2026-10-12, from the firm calc "Template - Circ Size
# Calc.xlsx", tab HOT WATER RECIRC HEAD LOSS):
#
#   BTU/h = effective length / 100 x heat loss per 100 ft
#   gpm   = BTU/h / (500 x temperature drop)      never below 0.5 gpm
#
# Two methods, which can both appear on one system:
#   LOOP    the return that starts after the last hot water fixture and runs
#           back to the water heater. Effective length is that return run.
#   BRANCH  a return picked up at one fixture. Effective length is the hot
#           supply branch from the main to that fixture, which is short, so a
#           branch lands on the 0.5 gpm minimum.
#
# One circuit setter is assumed per return leg (a leg is each place a return
# line starts). Walking back from the farthest leg toward the water heater the
# flows add at every junction, the same way fixture units add on the supply.
# The pipe sizes themselves are chosen in water_sizing_engine.
#
# Pure logic over a WaterGraph that has been through
# water_graph.apply_return_system_types(). No Revit imports.
#
# IronPython 2.7

import shared_params
import water_graph


METHOD_LOOP = "LOOP"
METHOD_BRANCH = "BRANCH"

# BTU/h = 500 x gpm x temperature drop, for water. Taken from the firm tab
# ("BTU=500*GPM*deltaT"), not derived here.
BTUH_PER_GPM_PER_DEGF = 500.0


def heat_loss_gpm(length_ft, heat_loss_btu_per_100ft, delta_t_f):
    """Circulation flow that makes up the heat loss of a length of pipe."""
    btuh = (float(length_ft) / 100.0) * float(heat_loss_btu_per_100ft)
    return btuh / (BTUH_PER_GPM_PER_DEGF * float(delta_t_f))


def analyze(graph,
            heat_loss_btu_per_100ft=shared_params.DEFAULT_RECIRC_HEAT_LOSS_BTU_PER_100FT,
            delta_t_f=shared_params.DEFAULT_RECIRC_DELTA_T_F,
            min_gpm=shared_params.MIN_CIRCUIT_SETTER_GPM):
    """Work out every circuit setter flow and the flow in every return pipe.

    Args:
        graph: a WaterGraph after water_graph.apply_return_system_types().
        heat_loss_btu_per_100ft: pipe heat loss, BTU/h per 100 ft.
        delta_t_f: temperature drop round the loop, degrees F.
        min_gpm: the lowest flow a circuit setter is set to.

    Returns:
        dict with:
          "legs"        list of dicts, one per circuit setter, farthest first:
                        number, method, start_pipe_id, start_element_id,
                        root_id, effective_length_ft, calc_gpm, gpm
          "pipe_gpm"    return pipe id -> accumulated gpm
          "root_gpm"    return line end (heater) id -> total gpm, the pump flow
          "total_gpm"   sum over every return line
          "unrooted_pipe_ids"  return pipes whose end was not found
          "heat_loss_btu_per_100ft", "delta_t_f", "min_gpm"  as used
    """
    if float(heat_loss_btu_per_100ft) <= 0 or float(delta_t_f) <= 0:
        raise ValueError(
            "Heat loss and temperature drop shall both be greater than zero "
            "(got {} and {}).".format(heat_loss_btu_per_100ft, delta_t_f))

    result = {
        "legs": [],
        "pipe_gpm": {},
        "root_gpm": {},
        "total_gpm": 0.0,
        "unrooted_pipe_ids": list(graph.return_unrooted),
        "heat_loss_btu_per_100ft": float(heat_loss_btu_per_100ft),
        "delta_t_f": float(delta_t_f),
        "min_gpm": float(min_gpm),
    }

    tree = water_graph.TREE_RETURN
    pipes = [node for node in graph.nodes.values()
             if water_graph.is_return_pipe(graph, node) and
             tree in node.parent_by_system]
    if not pipes:
        return result

    # --- where each return line starts -------------------------------------
    upstream = {}       # pipe id -> the return pipes that drain through it
    for pipe in pipes:
        upstream[pipe.element_id] = [
            nid for nid in graph.descendants(pipe.element_id, tree)
            if water_graph.is_return_pipe(graph, graph.nodes[nid])]

    starts = sorted(pid for pid, above in upstream.items() if not above)

    # --- one leg per start --------------------------------------------------
    adjacency = graph.hot_joints
    legs = []
    for start_id in starts:
        run_ft, last_pipe_id, root_id = _run_to_end(graph, start_id)
        legs.append({
            "start_pipe_id": start_id,
            "start_element_id": _start_element(graph, adjacency, start_id),
            "root_id": root_id,
            "return_run_ft": run_ft,
            "last_pipe_id": last_pipe_id,
        })

    # The LOOP is the longest return run on each line into the heater: the one
    # that starts after the last fixture. Every other leg joins it on the way
    # back and is a BRANCH.
    longest = {}
    for leg in legs:
        best = longest.get(leg["last_pipe_id"])
        if best is None or leg["return_run_ft"] > best["return_run_ft"]:
            longest[leg["last_pipe_id"]] = leg

    for leg in legs:
        if longest[leg["last_pipe_id"]] is leg:
            leg["method"] = METHOD_LOOP
            leg["effective_length_ft"] = leg["return_run_ft"]
        else:
            leg["method"] = METHOD_BRANCH
            leg["effective_length_ft"] = _supply_branch_length(
                graph, leg["start_element_id"])
        leg["calc_gpm"] = heat_loss_gpm(
            leg["effective_length_ft"], heat_loss_btu_per_100ft, delta_t_f)
        leg["gpm"] = max(float(min_gpm), leg["calc_gpm"])

    legs.sort(key=lambda l: (l["method"] != METHOD_LOOP,
                             -l["return_run_ft"], l["start_pipe_id"]))
    for number, leg in enumerate(legs):
        leg["number"] = number + 1

    # --- add the flows up toward the heater ---------------------------------
    gpm_by_start = dict((leg["start_pipe_id"], leg["gpm"]) for leg in legs)
    for pipe in pipes:
        pid = pipe.element_id
        flow = gpm_by_start.get(pid, 0.0)
        for nid in upstream[pid]:
            flow += gpm_by_start.get(nid, 0.0)
        result["pipe_gpm"][pid] = flow

    for leg in legs:
        result["root_gpm"][leg["root_id"]] = (
            result["root_gpm"].get(leg["root_id"], 0.0) + leg["gpm"])
    result["total_gpm"] = sum(result["root_gpm"].values())
    result["legs"] = legs
    return result


def _run_to_end(graph, start_pipe_id):
    """Return pipe length from a leg's start to where the return ends.

    Returns (length in feet, the last return pipe before the end, end id).
    """
    tree = water_graph.TREE_RETURN
    length = 0.0
    last_pipe_id = start_pipe_id
    current_id = start_pipe_id
    seen = set()
    while current_id is not None and current_id not in seen:
        seen.add(current_id)
        node = graph.nodes[current_id]
        if water_graph.is_return_pipe(graph, node):
            length += float(node.length_feet or 0.0)
            last_pipe_id = current_id
        parent_id = node.parent_by_system.get(tree)
        if parent_id is None:
            return length, last_pipe_id, current_id
        current_id = parent_id
    return length, last_pipe_id, current_id


def _start_element(graph, adjacency, start_pipe_id):
    """The supply-side element a return leg is picked up from.

    Either the fixture the return connects to, or the fitting at the end of
    the hot main. Found by stepping from the start pipe, through any valve or
    fitting on it, to the first element that sits on the hot supply tree.
    """
    hot = shared_params.SYSTEM_DOMESTIC_HOT_WATER
    tree = water_graph.TREE_RETURN

    candidates = [start_pipe_id]
    for nid in graph.descendants(start_pipe_id, tree):
        if graph.nodes[nid].kind != water_graph.KIND_PIPE:
            candidates.append(nid)

    for candidate_id in candidates:
        if (candidate_id != start_pipe_id and
                hot in graph.nodes[candidate_id].parent_by_system):
            return candidate_id
        toward_heater = graph.nodes[candidate_id].parent_by_system.get(tree)
        for neighbour_id in adjacency.get(candidate_id, []):
            if neighbour_id in candidates or neighbour_id == toward_heater:
                continue
            neighbour = graph.nodes[neighbour_id]
            if water_graph.is_return_pipe(graph, neighbour):
                continue
            if hot in neighbour.parent_by_system:
                return neighbour_id
    return None


def _supply_branch_length(graph, start_element_id):
    """Hot supply length from the main to the fixture a branch return serves.

    Walks up the supply tree from the fixture and adds pipe length until it
    reaches a pipe that serves two or more fixtures, which is the main.
    """
    if start_element_id is None:
        return 0.0
    hot = shared_params.SYSTEM_DOMESTIC_HOT_WATER

    length = 0.0
    current_id = start_element_id
    seen = set()
    while current_id not in seen:
        seen.add(current_id)
        parent_id = graph.nodes[current_id].parent_by_system.get(hot)
        if parent_id is None:
            break
        parent = graph.nodes[parent_id]
        if parent.kind == water_graph.KIND_HEATER:
            break
        if parent.kind == water_graph.KIND_PIPE:
            fixtures = [nid for nid in graph.descendants(parent_id, hot)
                        if graph.nodes[nid].kind == water_graph.KIND_FIXTURE]
            if len(fixtures) >= 2:
                break
            length += float(parent.length_feet or 0.0)
        current_id = parent_id
    return length
