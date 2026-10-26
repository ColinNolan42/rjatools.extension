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
            "heater_id": _heater_for(graph, root_id),
            "safety_factor": float(safety_factor),
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
           "total_psi": 0.0, "length_ft": 0.0, "notes": [],
           "supply_ft": 0.0, "return_ft": 0.0, "device_qty": 0,
           "fittings": {}}

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
            out[key.replace("_psi", "_ft")] += node.length_feet

        for i, nid in enumerate(path):
            node = graph.nodes.get(nid)
            if node is None or node.kind in (water_graph.KIND_PIPE,
                                             water_graph.KIND_PUMP,
                                             water_graph.KIND_HEATER):
                if node is not None and node.kind == water_graph.KIND_HEATER:
                    drop = getattr(node, "device_drop_psi", None)
                    if drop and key == "supply_psi":
                        out["device_psi"] += float(drop)
                        out["device_qty"] += 1
                continue
            prev_id = path[i - 1] if i > 0 else None
            next_id = path[i + 1] if i + 1 < len(path) else None
            drop = getattr(node, "device_drop_psi", None)
            if drop:
                out["device_psi"] += float(drop)
                out["device_qty"] += 1
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
            entry = out["fittings"].setdefault(
                kind, {"qty": 0, "length_ft": 0.0, "psi": 0.0})
            entry["qty"] += 1
            entry["length_ft"] += length
            entry["psi"] += length * rate / 100.0

    out["total_psi"] = (out["supply_psi"] + out["return_psi"] +
                        out["fitting_psi"] + out["setter_psi"] +
                        out["device_psi"])
    return out


def _heater_for(graph, root_id):
    """The water heater a return line ends at.

    The end of the return tree is usually the transition or fitting on the
    heater's cold inlet, not the heater itself, so look a few joints out.
    """
    if graph.nodes.get(root_id) is None:
        return root_id
    seen = set([root_id])
    frontier = [root_id]
    for _depth in range(4):
        following = []
        for nid in frontier:
            current = graph.nodes.get(nid)
            if current is None:
                continue
            if current.kind == water_graph.KIND_HEATER:
                return nid
            for c in current.connectors or []:
                other = c.get("connected_element_id")
                if other is not None and other not in seen:
                    seen.add(other)
                    following.append(other)
        frontier = following
    for nid in seen:
        other = graph.nodes.get(nid)
        if other is not None and other.kind == water_graph.KIND_HEATER:
            return nid
    return root_id


TABLE_COLUMNS = [("HEATER", 0.090), ("PUMP GPM", 0.075),
                 ("HEAD FT", 0.070), ("INDEX LEG", 0.080)]

# The system head chart: every item the pump has to overcome.
HEAD_COLUMNS = [("SYSTEM HEAD ITEM", 0.200), ("QTY", 0.050),
                ("EQ. LENGTH FT", 0.110), ("PSI", 0.070)]

SIGMA = u"\u03a3"

_FITTING_ORDER = ("elbow_90", "elbow_45", "tee_run", "tee_branch",
                  "coupling")
_FITTING_NAMES = {
    "elbow_90": "90 degree elbows", "elbow_45": "45 degree elbows",
    "tee_run": "Tees, run", "tee_branch": "Tees, branch",
    "coupling": "Transitions and couplings",
}


def table_rows(points):
    rows = []
    for p in points:
        rows.append([str(p.get("heater_id", p["root_id"])),
                     "{:.1f}".format(p["gpm"]),
                     "{:.1f}".format(p["head_ft"]), str(p["index_leg"])])
    return rows


def index_circuit(point):
    """The circuit that sets the head."""
    for circuit in point["circuits"]:
        if circuit["number"] == point["index_leg"]:
            return circuit
    return point["circuits"][0] if point["circuits"] else None


def head_rows(point):
    """The system head chart for one heater.

    Returns (rows, equation). Rows follow HEAD_COLUMNS; the equation is the
    line printed under the chart.
    """
    circuit = index_circuit(point)
    if circuit is None:
        return [], ""
    rows = [
        ["Supply pipe friction", "", "{:.1f}".format(circuit["supply_ft"]),
         "{:.2f}".format(circuit["supply_psi"])],
        ["Return pipe friction", "", "{:.1f}".format(circuit["return_ft"]),
         "{:.2f}".format(circuit["return_psi"])],
    ]
    kinds = list(_FITTING_ORDER) + sorted(
        k for k in circuit["fittings"] if k not in _FITTING_ORDER)
    for kind in kinds:
        entry = circuit["fittings"].get(kind)
        if not entry:
            continue
        name = _FITTING_NAMES.get(
            kind, water_pressure.FITTING_LABELS.get(kind, kind).capitalize())
        rows.append([name, str(entry["qty"]),
                     "{:.1f}".format(entry["length_ft"]),
                     "{:.2f}".format(entry["psi"])])
    rows.append(["Circuit setter", "1", "",
                 "{:.2f}".format(circuit["setter_psi"])])
    if circuit["device_psi"]:
        rows.append(["Devices", str(circuit["device_qty"]), "",
                     "{:.2f}".format(circuit["device_psi"])])
    rows.append(["TOTAL", "", "", "{:.2f}".format(circuit["total_psi"])])
    equation = ("Head = {:.2f} ft/psi x {:.2f} psi x (1 + {:.2f}) = "
                "{:.1f} ft".format(FT_PER_PSI, circuit["total_psi"],
                                   point.get("safety_factor", 0.0),
                                   point["head_ft"]))
    return rows, equation


def flow_line(point):
    """The pump flow as the sum of its circuit setter flows."""
    return (u"Pump duty point: flow = " + SIGMA + u" circuit setter flow = "
            u"{:.1f} gpm".format(float(point["gpm"])))


def basis_lines(circuit_setter_psi, safety_factor):
    return [
        u"Flow = " + SIGMA + u" circuit setter flow at the heater.",
        "Head = {:.2f} ft/psi x (pipe friction + fittings + {:.2f} psi "
        "circuit setter + devices) x (1 + {:.2f}).".format(
            FT_PER_PSI, float(circuit_setter_psi), float(safety_factor)),
        "A duty point, not a selection. Choose the pump from its curve.",
    ]
