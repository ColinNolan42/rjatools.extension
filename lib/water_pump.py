# -*- coding: ascii -*-
# water_pump.py
# Hot water recirculation pump duty point: flow and head.
#
#   Flow  the circuit setter flows added up at each water heater, exactly as
#         water_recirc.analyze works them out (the firm circ size calc).
#   Head  the friction round the worst circuit (the index circuit): out of the
#         heater along the hot supply to where the return leg starts, then
#         back along the return to the heater. Each pipe is priced at the
#         circulation flow it carries with no fixture drawing: a supply pipe
#         carries the flow of every leg beyond it, a return pipe its
#         accumulated return flow. Fittings and valves by Table E103.3(6),
#         the circuit setter as a design pressure drop, devices by
#         Device_Pressure_Drop_PSI. Friction and fittings use the same
#         Darcy-Weisbach basis as the pressure loss check (water_pressure).
#   Then a safety factor on the head.
#
# A duty point, not a pump selection: the model is picked from a pump curve.
#
# Pure logic, no Revit. IronPython 2.7

import shared_params
import water_graph
import water_performance
import water_pressure

HOT = shared_params.SYSTEM_DOMESTIC_HOT_WATER
FT_PER_PSI = 2.31

# Firm defaults (Colin, 2026-10-20). Both are dialog inputs and are printed
# in the basis.
DEFAULT_CIRCUIT_SETTER_PSI = 1.0
DEFAULT_SAFETY_FACTOR = 0.10


def duty_points(graph, recirc, circuit_setter_psi=DEFAULT_CIRCUIT_SETTER_PSI,
                safety_factor=DEFAULT_SAFETY_FACTOR):
    """Pump flow and head at each water heater with a return.

    Args:
        graph: WaterGraph after apply_return_system_types.
        recirc: what water_recirc.analyze returned.
        circuit_setter_psi: design drop across one circuit setter.
        safety_factor: fraction added to the index circuit head.

    Returns:
        list of dicts, one per return end (heater): root_id, gpm, head_ft,
        head_psi (before the safety factor), index_leg (number), circuits
        (list of per leg dicts: number, method, supply_psi, return_psi,
        fitting_psi, setter_psi, device_psi, total_psi, notes).
    """
    if recirc is None or not recirc.get("legs"):
        return []

    legs = recirc["legs"]
    supply_paths = {}
    for leg in legs:
        supply_paths[leg["number"]] = _supply_path(graph,
                                                   leg["start_element_id"])

    # No-draw flow in each supply pipe: every leg whose supply path uses it.
    supply_gpm = {}
    for leg in legs:
        for nid in supply_paths[leg["number"]] or []:
            node = graph.nodes.get(nid)
            if node is not None and node.kind == water_graph.KIND_PIPE:
                supply_gpm[nid] = supply_gpm.get(nid, 0.0) + leg["gpm"]

    circuits = {}
    for leg in legs:
        circuits[leg["number"]] = _circuit(
            graph, leg, supply_paths[leg["number"]], supply_gpm,
            recirc["pipe_gpm"], circuit_setter_psi)

    points = []
    for root_id, gpm in sorted(recirc["root_gpm"].items()):
        mine = [circuits[l["number"]] for l in legs if l["root_id"] == root_id]
        if not mine:
            continue
        index = max(mine, key=lambda c: c["total_psi"])
        head_psi = index["total_psi"]
        points.append({
            "root_id": root_id,
            "gpm": float(gpm),
            "head_psi": head_psi,
            "head_ft": head_psi * FT_PER_PSI * (1.0 + float(safety_factor)),
            "index_leg": index["number"],
            "circuits": sorted(mine, key=lambda c: -c["total_psi"]),
        })
    return points


def _supply_path(graph, start_element_id):
    """Ids from the heater to the leg's start element, in flow order."""
    if start_element_id is None:
        return None
    ids = []
    current = start_element_id
    seen = set()
    while current is not None and current not in seen:
        seen.add(current)
        node = graph.nodes.get(current)
        if node is None:
            return None
        ids.append(current)
        if node.kind == water_graph.KIND_HEATER:
            ids.reverse()
            return ids
        current = node.parent_by_system.get(HOT)
    return None


def _return_path(graph, start_pipe_id):
    """Ids from the leg's first return pipe to the return end, flow order."""
    ids = []
    current = start_pipe_id
    seen = set()
    tree = water_graph.TREE_RETURN
    while current is not None and current not in seen:
        seen.add(current)
        ids.append(current)
        node = graph.nodes.get(current)
        if node is None:
            break
        current = node.parent_by_system.get(tree)
    return ids


def _circuit(graph, leg, supply_path, supply_gpm, return_gpm,
             circuit_setter_psi):
    out = {"number": leg["number"], "method": leg["method"],
           "supply_psi": 0.0, "return_psi": 0.0, "fitting_psi": 0.0,
           "setter_psi": float(circuit_setter_psi), "device_psi": 0.0,
           "total_psi": 0.0, "length_ft": 0.0, "notes": []}

    if supply_path is None:
        out["notes"].append("supply path to the heater not found, supply "
                            "side not counted")
        supply_path = []
    ret_path = _return_path(graph, leg["start_pipe_id"])

    for path, flows, key in ((supply_path, supply_gpm, "supply_psi"),
                             (ret_path, return_gpm, "return_psi")):
        rates = {}
        for nid in path:
            node = graph.nodes.get(nid)
            if node is None or node.kind != water_graph.KIND_PIPE:
                continue
            rate = water_pressure.friction_psi_per_100ft(
                flows.get(nid, 0.0),
                water_performance.inside_diameter_in(node.diameter_inches))
            if rate is None:
                out["notes"].append("pipe {} size not tabulated, not "
                                    "priced".format(nid))
                rate = 0.0
            rates[nid] = rate
            out[key] += node.length_feet * rate / 100.0
            out["length_ft"] += node.length_feet

        for i, nid in enumerate(path):
            node = graph.nodes.get(nid)
            if node is None or node.kind in (water_graph.KIND_PIPE,
                                             water_graph.KIND_PUMP,
                                             water_graph.KIND_HEATER):
                if node is not None and node.kind == water_graph.KIND_HEATER:
                    drop = getattr(node, "device_drop_psi", None)
                    if drop and key == "supply_psi":
                        out["device_psi"] += float(drop)
                continue
            prev_id = path[i - 1] if i > 0 else None
            next_id = path[i + 1] if i + 1 < len(path) else None
            drop = getattr(node, "device_drop_psi", None)
            if drop:
                out["device_psi"] += float(drop)
                continue
            if node.kind == water_graph.KIND_FITTING:
                kind, size, _note = water_pressure.classify_fitting(
                    node, prev_id, next_id)
            elif node.kind == water_graph.KIND_ACCESSORY:
                kind = water_pressure.valve_kind(node)
                size = water_pressure._first_size(node)
                if kind is None:
                    continue
            elif node.kind == water_graph.KIND_FIXTURE:
                kind, size = "tee_run", water_pressure._first_size(node)
            else:
                continue
            rate = rates.get(next_id, rates.get(prev_id, 0.0))
            length, _why = water_pressure.equivalent_length_ft(kind, size)
            out["fitting_psi"] += length * rate / 100.0
            out["length_ft"] += length

    out["total_psi"] = (out["supply_psi"] + out["return_psi"] +
                        out["fitting_psi"] + out["setter_psi"] +
                        out["device_psi"])
    return out


TABLE_COLUMNS = [("HEATER", 0.090), ("PUMP GPM", 0.075),
                 ("HEAD FT", 0.070), ("INDEX LEG", 0.080),
                 ("SUPPLY PSI", 0.080), ("RETURN PSI", 0.080),
                 ("FITTINGS PSI", 0.090), ("SETTER PSI", 0.080),
                 ("DEVICES PSI", 0.085)]


def table_rows(points):
    rows = []
    for p in points:
        index = [c for c in p["circuits"] if c["number"] == p["index_leg"]][0]
        rows.append([str(p["root_id"]), "{:.1f}".format(p["gpm"]),
                     "{:.1f}".format(p["head_ft"]), str(p["index_leg"]),
                     "{:.2f}".format(index["supply_psi"]),
                     "{:.2f}".format(index["return_psi"]),
                     "{:.2f}".format(index["fitting_psi"]),
                     "{:.2f}".format(index["setter_psi"]),
                     "{:.2f}".format(index["device_psi"])])
    return rows


def basis_lines(circuit_setter_psi, safety_factor):
    return [
        "Pump duty point: flow = circuit setter flows added at the heater "
        "(firm circ size calc). Head = friction round the index circuit, "
        "heater out along the hot supply to the return leg and back, each "
        "pipe at the circulation flow it carries with no fixture drawing.",
        "  Friction and fittings on the pressure loss basis (Darcy-Weisbach, "
        "Table E103.3(6)). Circuit setter {:.2f} psi. Devices by Device_"
        "Pressure_Drop_PSI. Safety factor {:.0f} percent on head. "
        "{:.2f} ft per psi.".format(float(circuit_setter_psi),
                                     float(safety_factor) * 100.0,
                                     FT_PER_PSI),
        "  A duty point, not a selection: choose the pump from its curve.",
    ]
