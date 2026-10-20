# -*- coding: ascii -*-
# water_pressure.py
# Domestic water pressure loss, 2024 IPC Appendix E, Section E103.3
# (segmented loss method), checked against Table 604.3.
#
# Pure logic, no Revit, so it can be unit tested in CPython.
#
# Basis (Colin, 2026-10-20):
#   - Pressure is entered at the picked start point: the pressure the RPZ
#     delivers, or the pressure in the existing main at the tee. Street,
#     tap, meter and RPZ losses are upstream of it and not modeled.
#   - Pipe friction: Darcy-Weisbach, Swamee-Jain friction factor, drawn
#     copper (Colin's choice over Hazen-Williams C=150). Against the code's
#     Type L chart, Figure E103.3(3), it reads 4 to 10 percent LOW at every
#     point checked; Hazen-Williams C=150 reads 2 to 10 percent low. Pinned by
#     _test_water_pressure.TestFriction.
#   - Fittings and valves: equivalent length from Table E103.3(6), added to
#     the pipe they sit on and priced at that pipe's friction rate (code text
#     E103.3 item 3.6). Fitting kind is read from the connector geometry,
#     never from a family name.
#   - Static head 0.433 psi per ft of rise from the start point to the
#     fixture outlet. A fixture below the start point gains it back
#     (Table E103.3(1) note a).
#   - Flow in each pipe: the probable demand the sizing run worked out
#     (Table E103.3(3)), at the DRAWN size.
#
# IronPython 2.7

import math

import shared_params
import water_graph
import water_performance

COLD = shared_params.SYSTEM_DOMESTIC_COLD_WATER
HOT = shared_params.SYSTEM_DOMESTIC_HOT_WATER

PSI_PER_FT_HEAD = 0.433          # E103.3 item 3.4
MAX_STATIC_PSI = 80.0            # Section 604.8
ROUGHNESS_FT = 0.000005          # drawn copper tubing, 1.5 micron
KINEMATIC_VISCOSITY_FT2_S = 1.217e-5   # water at 60 F
WATER_TEMP_F = 60.0              # all water, hot included (conservative)
G_FT_S2 = 32.174
DEFAULT_START_PSI = 38.6         # RJA template: 60 street - 1.4 tap - 5 meter - 15 RPZ

# 2024 IPC Table E103.3(6), page E-14. Equivalent length of tube, ft.
# None is a dash in the table. Note a: soldered fittings; double for
# threaded; based on Hazen-Williams C = 150; rounded to the nearest half ft.
TABLE_E103_3_6 = {
    0.375: {"elbow_90": 0.5, "elbow_45": None, "tee_branch": 1.5, "tee_run": None, "coupling": None, "ball": None, "gate": None, "butterfly": None, "check": 1.5},
    0.5:   {"elbow_90": 1.0, "elbow_45": 0.5, "tee_branch": 2.0, "tee_run": None, "coupling": None, "ball": None, "gate": None, "butterfly": None, "check": 2.0},
    0.625: {"elbow_90": 1.5, "elbow_45": 0.5, "tee_branch": 2.0, "tee_run": None, "coupling": None, "ball": None, "gate": None, "butterfly": None, "check": 2.5},
    0.75:  {"elbow_90": 2.0, "elbow_45": 0.5, "tee_branch": 3.0, "tee_run": None, "coupling": None, "ball": None, "gate": None, "butterfly": None, "check": 3.0},
    1.0:   {"elbow_90": 2.5, "elbow_45": 1.0, "tee_branch": 4.5, "tee_run": None, "coupling": None, "ball": 0.5, "gate": None, "butterfly": None, "check": 4.5},
    1.25:  {"elbow_90": 3.0, "elbow_45": 1.0, "tee_branch": 5.5, "tee_run": 0.5, "coupling": 0.5, "ball": 0.5, "gate": None, "butterfly": None, "check": 5.5},
    1.5:   {"elbow_90": 4.0, "elbow_45": 1.5, "tee_branch": 7.0, "tee_run": 0.5, "coupling": 0.5, "ball": 0.5, "gate": None, "butterfly": None, "check": 6.5},
    2.0:   {"elbow_90": 5.5, "elbow_45": 2.0, "tee_branch": 9.0, "tee_run": 0.5, "coupling": 0.5, "ball": 0.5, "gate": 0.5, "butterfly": 7.5, "check": 9.0},
    2.5:   {"elbow_90": 7.0, "elbow_45": 2.5, "tee_branch": 12.0, "tee_run": 0.5, "coupling": 0.5, "ball": None, "gate": 1.0, "butterfly": 10.0, "check": 11.5},
    3.0:   {"elbow_90": 9.0, "elbow_45": 3.5, "tee_branch": 15.0, "tee_run": 1.0, "coupling": 1.0, "ball": None, "gate": 1.5, "butterfly": 15.5, "check": 14.5},
    3.5:   {"elbow_90": 9.0, "elbow_45": 3.5, "tee_branch": 14.0, "tee_run": 1.0, "coupling": 1.0, "ball": None, "gate": 2.0, "butterfly": None, "check": 12.5},
    4.0:   {"elbow_90": 12.5, "elbow_45": 5.0, "tee_branch": 21.0, "tee_run": 1.0, "coupling": 1.0, "ball": None, "gate": 2.0, "butterfly": 16.0, "check": 18.5},
    5.0:   {"elbow_90": 16.0, "elbow_45": 6.0, "tee_branch": 27.0, "tee_run": 1.5, "coupling": 1.5, "ball": None, "gate": 3.0, "butterfly": 11.5, "check": 23.5},
    6.0:   {"elbow_90": 19.0, "elbow_45": 7.0, "tee_branch": 34.0, "tee_run": 2.0, "coupling": 2.0, "ball": None, "gate": 3.5, "butterfly": 13.5, "check": 26.5},
    8.0:   {"elbow_90": 29.0, "elbow_45": 11.0, "tee_branch": 50.0, "tee_run": 3.0, "coupling": 3.0, "ball": None, "gate": 5.0, "butterfly": 12.5, "check": 39.0},
}

FITTING_LABELS = {
    "elbow_90": "90 elbow", "elbow_45": "45 elbow", "tee_branch": "tee branch",
    "tee_run": "tee run", "coupling": "coupling", "ball": "ball valve",
    "gate": "gate valve", "butterfly": "butterfly valve",
    "check": "check valve",
}

# Valve kind from the accessory's family or type name. Words, not exact
# family names, so any manufacturer's valve family is recognised.
_VALVE_WORDS = (("butterfly", "butterfly"), ("ball", "ball"),
                ("gate", "gate"), ("check", "check"))

# A pressure reducing valve holds its set pressure downstream, so it is not a
# fixed loss. Outlets behind one are reported, not judged.
_PRV_WORDS = ("pressure reducing", "pressure regulating", "prv",
              "reducing valve")

# 2024 IPC Table 604.3, pages 6-27 and 6-28: flow pressure (psi) at the
# fixture supply outlet. Matched to the RJA water fixture Types by words in
# the Type name, first match wins.
#   Flushometer valve water closet: siphonic row (35 psi). The blowout row is
#     45 psi; a blowout bowl shall be checked by hand.
#   Showers and bathtubs: the balanced-pressure / thermostatic rows (20 psi),
#     since commercial showers and tubs carry those valves (Section 424).
#   Flush tank urinal: no Table 604.3 row, takes the tank water closet value
#     (Colin, 2026-10-20).
FIXTURE_PRESSURE_RULES = [
    (("water closet", "flushometer valve"), 35.0,
     "Water closet, siphonic, flushometer valve"),
    (("water closet", "flushometer tank"), 20.0,
     "Water closet, flushometer tank"),
    (("water closet",), 20.0, "Water closet, tank"),
    (("urinal", "flushometer"), 25.0, "Urinal, valve"),
    (("urinal",), 20.0, "Urinal, flush tank (as Water closet, tank, Colin)"),
    (("lavatory",), 8.0, "Lavatory"),
    (("kitchen sink",), 8.0, "Sink, residential"),
    (("service sink",), 8.0, "Sink, service"),
    (("mop",), 8.0, "Sink, service"),
    (("shower",), 20.0, "Shower, balanced-pressure / thermostatic valve"),
    (("bathtub",), 20.0, "Bathtub, balanced-pressure / thermostatic valve"),
    (("dishwash",), 8.0, "Dishwasher, residential"),
    (("drinking fountain",), 8.0, "Drinking fountain"),
    (("hose bibb",), 8.0, "Sill cock, hose bibb"),
    (("sill cock",), 8.0, "Sill cock, hose bibb"),
]


# ---------------------------------------------------------------------------
# Friction
# ---------------------------------------------------------------------------

def friction_psi_per_100ft(gpm, inside_diameter_in):
    """Darcy-Weisbach loss in psi per 100 ft, or None when not computable."""
    if gpm is None or inside_diameter_in is None:
        return None
    gpm = float(gpm)
    if gpm <= 0.0:
        return 0.0
    d_ft = float(inside_diameter_in) / 12.0
    velocity = gpm / (water_performance.GPM_PER_FPS_IN2 *
                      float(inside_diameter_in) ** 2)
    reynolds = velocity * d_ft / KINEMATIC_VISCOSITY_FT2_S
    if reynolds < 2300.0:
        factor = 64.0 / reynolds
    else:
        factor = 0.25 / (math.log10(ROUGHNESS_FT / (3.7 * d_ft) +
                                    5.74 / reynolds ** 0.9)) ** 2
    head_ft_per_ft = factor / d_ft * velocity ** 2 / (2.0 * G_FT_S2)
    return head_ft_per_ft * 100.0 * PSI_PER_FT_HEAD


def equivalent_length_ft(kind, nominal_in):
    """Table E103.3(6) value for a fitting or valve kind, or (0, note).

    The size is the next tabulated size at or above the nominal. A dash in
    the table reads as 0, which is what the code means by it.
    """
    if nominal_in is None:
        return 0.0, "size not read"
    sizes = sorted(TABLE_E103_3_6.keys())
    key = None
    for size in sizes:
        if size >= float(nominal_in) - 0.001:
            key = size
            break
    if key is None:
        return 0.0, "larger than Table E103.3(6)"
    value = TABLE_E103_3_6[key].get(kind)
    return (float(value) if value is not None else 0.0), ""


# ---------------------------------------------------------------------------
# Fixture minimum pressure
# ---------------------------------------------------------------------------

def required_pressure(node):
    """(psi, Table 604.3 row) for a fixture node, or (None, reason)."""
    text = "{} {}".format(getattr(node, "type_name", "") or "",
                          getattr(node, "fixture_name", "") or "").lower()
    for words, psi, row in FIXTURE_PRESSURE_RULES:
        if all(word in text for word in words):
            return psi, row
    return None, "no Table 604.3 row matched the Type name"


# ---------------------------------------------------------------------------
# Fitting kind from geometry
# ---------------------------------------------------------------------------

def classify_fitting(node, from_id, to_id):
    """Fitting kind for flow passing from from_id to to_id through node.

    Read from the connector positions: two connectors in line are a
    coupling (tee run on a tee), a bend is an elbow (45 when it turns 60
    degrees or less), and leaving or entering a tee off the line is the
    branch. When the geometry cannot be read the worst case is taken (tee
    branch, or 90 elbow) and noted.

    Returns (kind, size_in, note).
    """
    count =len(node.connectors or []) or getattr(node, "connector_count", 0)
    c_in = _connector_to(node, from_id)
    c_out = _connector_to(node, to_id)
    size = _first_size(c_out, c_in, node)
    center = getattr(node, "location_xyz", None)

    if c_in is None or c_out is None or center is None or \
            c_in.get("origin_xyz") is None or c_out.get("origin_xyz") is None:
        kind = "tee_branch" if count >= 3 else "elbow_90"
        return kind, size, "geometry not read, worst case taken"

    v_in = _vector(center, c_in["origin_xyz"])
    v_out = _vector(center, c_out["origin_xyz"])
    angle = _angle_deg(v_in, v_out)
    if angle is None:
        kind = "tee_branch" if count >= 3 else "elbow_90"
        return kind, size, "geometry not read, worst case taken"

    in_line = angle >= 170.0
    if count >= 3:
        return ("tee_run" if in_line else "tee_branch"), size, ""
    if in_line:
        return "coupling", size, ""
    turn = 180.0 - angle
    return ("elbow_45" if turn <= 60.0 else "elbow_90"), size, ""


def valve_kind(node):
    """Valve kind from the accessory's family/type name, or None."""
    text = "{} {}".format(getattr(node, "family_name", "") or "",
                          getattr(node, "type_name", "") or "").lower()
    for word, kind in _VALVE_WORDS:
        if word in text:
            return kind
    return None


def is_prv(node):
    """True for a pressure reducing valve, from its family/type name."""
    text = "{} {}".format(getattr(node, "family_name", "") or "",
                          getattr(node, "type_name", "") or "").lower()
    return any(word in text for word in _PRV_WORDS)


def _connector_to(node, other_id):
    for c in node.connectors or []:
        if c.get("connected_element_id") == other_id:
            return c
    return None


def _first_size(*sources):
    for src in sources:
        if src is None:
            continue
        if isinstance(src, dict):
            if src.get("diameter_in"):
                return float(src["diameter_in"])
        else:
            for c in getattr(src, "connectors", None) or []:
                if c.get("diameter_in"):
                    return float(c["diameter_in"])
    return None


def _vector(a, b):
    return (b[0] - a[0], b[1] - a[1], b[2] - a[2])


def _angle_deg(u, v):
    nu = math.sqrt(sum(x * x for x in u))
    nv = math.sqrt(sum(x * x for x in v))
    if nu < 1e-6 or nv < 1e-6:
        return None
    cos = sum(a * b for a, b in zip(u, v)) / (nu * nv)
    cos = max(-1.0, min(1.0, cos))
    return math.degrees(math.acos(cos))


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

def path_to_fixture(graph, fixture_id, system):
    """Element ids from the origin to the fixture, in flow order.

    Cold: up the cold tree. Hot: up the hot tree to the heater, then up the
    cold tree from the heater to the origin. Returns None when the walk does
    not reach the origin.
    """
    ids = [fixture_id]
    current = fixture_id
    tree = system
    guard = 0
    while current != graph.origin_id:
        guard += 1
        if guard > 100000:
            return None
        node = graph.nodes.get(current)
        if node is None:
            return None
        if tree == HOT and node.kind == water_graph.KIND_HEATER:
            tree = COLD
        parent = node.parent_by_system.get(tree)
        if parent is None:
            return None
        ids.append(parent)
        current = parent
    ids.reverse()
    return ids


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

class OutletResult(object):
    """Pressure at one fixture outlet (one system)."""

    def __init__(self, fixture_id, system):
        self.fixture_id = fixture_id
        self.system = system
        self.fixture_name = ""
        self.row = ""
        self.required_psi = None
        self.start_psi = 0.0
        self.static_psi = 0.0
        self.pipe_psi = 0.0
        self.fitting_psi = 0.0
        self.device_psi = 0.0
        self.developed_ft = 0.0
        self.residual_psi = None
        self.margin_psi = None
        self.passes = None
        self.notes = []
        self.reached = True          # a path to the start point was found
        self.elevation_read = True
        self.unpriced = []           # items whose loss could not be priced
        self.behind_prv = False


def analyze(graph, sizing_result, start_psi=DEFAULT_START_PSI):
    """Pressure left at every fixture outlet.

    Args:
        graph: WaterGraph after apply_return_system_types.
        sizing_result: what water_sizing_engine.size_network returned; its
            segment demand_gpm is the flow in each pipe.
        start_psi: pressure at the picked start point.

    Returns dict:
        "outlets"   list of OutletResult, worst margin first
        "critical"  the OutletResult with the least margin, or None
        "max_static_psi"  highest static pressure at any outlet
        "warnings"  list of strings
    """
    gpm_by_pipe = {}
    for seg in sizing_result.get("segments", []):
        if not seg.is_return:
            gpm_by_pipe[seg.element_id] = seg.demand_gpm

    origin = graph.nodes.get(graph.origin_id)
    origin_z = _z(origin)
    outlets = []
    warnings = []
    max_static = None

    for fid in graph.fixture_ids:
        fixture = graph.nodes.get(fid)
        if fixture is None:
            continue
        systems = []
        if fixture.has_cw and float(fixture.cw_wsfu or 0.0) > 0.0:
            systems.append(COLD)
        if fixture.has_hw and float(fixture.hw_wsfu or 0.0) > 0.0:
            systems.append(HOT)
        for system in systems:
            result = _outlet(graph, fixture, system, start_psi, origin_z,
                             gpm_by_pipe)
            outlets.append(result)
            if (not (result.reached and result.elevation_read) or
                    result.behind_prv):
                continue
            static_here = float(start_psi) - result.static_psi
            if max_static is None or static_here > max_static:
                max_static = static_here

    judged = [o for o in outlets if o.margin_psi is not None]
    judged.sort(key=lambda o: o.margin_psi)
    unjudged = [o for o in outlets if o.margin_psi is None]
    critical = judged[0] if judged else None

    if max_static is not None and max_static > MAX_STATIC_PSI:
        warnings.append(
            "Static pressure reaches {:.1f} psi at the lowest outlet, above "
            "the {:.0f} psi limit of Section 604.8. A pressure reducing valve "
            "shall be provided.".format(max_static, MAX_STATIC_PSI))
    if origin_z is None:
        warnings.append("Start point elevation not read, static head not "
                        "counted.")
    unreached = sorted(set(o.fixture_name for o in outlets if not o.reached))
    if unreached:
        warnings.append("No path to the start point for: {}. Not judged."
                        .format(", ".join(unreached)))
    unpriced = [o for o in outlets if o.reached and o.unpriced]
    if unpriced:
        warnings.append("{} outlet(s) not judged because a pipe or fitting on "
                        "the way could not be priced (see notes).".format(
                            len(unpriced)))
    if any(o.behind_prv for o in outlets):
        warnings.append("A pressure reducing valve is on the way to some "
                        "outlets. Its set pressure is not modeled, so those "
                        "outlets are not judged.")
    return {"outlets": judged + unjudged, "critical": critical,
            "max_static_psi": max_static, "warnings": warnings}


def _outlet(graph, fixture, system, start_psi, origin_z, gpm_by_pipe):
    result = OutletResult(fixture.element_id, system)
    result.fixture_name = fixture.fixture_name or str(fixture.element_id)
    result.start_psi = float(start_psi)
    result.required_psi, result.row = required_pressure(fixture)

    path = path_to_fixture(graph, fixture.element_id, system)
    if path is None:
        result.reached = False
        result.notes.append("path to the start point not found")
        return result

    fixture_z = _connector_z(fixture, system)
    if fixture_z is None:
        fixture_z = _z(fixture)
    if origin_z is not None and fixture_z is not None:
        result.static_psi = (fixture_z - origin_z) * PSI_PER_FT_HEAD
    else:
        result.elevation_read = False
        result.unpriced.append("elevation")
        result.notes.append("elevation not read, static head not counted")

    # Friction rate of each pipe on the path, for pricing the fittings that
    # sit on it.
    rates = {}
    for nid in path:
        node = graph.nodes.get(nid)
        if node is None or node.kind != water_graph.KIND_PIPE:
            continue
        gpm = gpm_by_pipe.get(nid)
        rate = friction_psi_per_100ft(
            gpm, water_performance.inside_diameter_in(node.diameter_inches))
        if rate is None:
            result.unpriced.append(nid)
            result.notes.append("pipe {} not priced ({})".format(
                nid, "no flow" if gpm is None else "size not tabulated"))
            rate = 0.0
        rates[nid] = rate
        result.pipe_psi += node.length_feet * rate / 100.0
        result.developed_ft += node.length_feet

    for index, nid in enumerate(path):
        node = graph.nodes.get(nid)
        if node is None or node.kind == water_graph.KIND_PIPE:
            continue
        if nid == graph.origin_id or nid == fixture.element_id:
            continue
        prev_id = path[index - 1] if index > 0 else None
        next_id = path[index + 1] if index + 1 < len(path) else None
        rate = _adjacent_rate(rates, next_id, prev_id)

        drop = getattr(node, "device_drop_psi", None)
        if drop:
            result.device_psi += float(drop)
            continue

        kind, size, note = None, None, ""
        if node.kind == water_graph.KIND_FITTING:
            kind, size, note = classify_fitting(node, prev_id, next_id)
        elif node.kind == water_graph.KIND_ACCESSORY:
            if is_prv(node):
                result.behind_prv = True
                result.notes.append("pressure reducing valve {} on the way, "
                                    "set pressure not modeled".format(nid))
                continue
            kind = valve_kind(node)
            size = _first_size(node)
            if kind is None:
                result.notes.append("{} '{}' not priced (no Device_Pressure_"
                                    "Drop_PSI, not a known valve)".format(
                                        nid, node.family_name))
                continue
        elif node.kind == water_graph.KIND_FIXTURE:
            # A fixture the hot main threads through acts as a tee
            # (Colin, 2026-10-12).
            kind, size = "tee_run", _first_size(node)
        elif node.kind == water_graph.KIND_HEATER:
            result.notes.append("water heater {} has no Device_Pressure_"
                                "Drop_PSI, counted as 0".format(nid))
            continue
        else:
            continue
        if note:
            result.notes.append("{} {}: {}".format(
                FITTING_LABELS.get(kind, kind), nid, note))
        length, why = equivalent_length_ft(kind, size)
        if why:
            result.unpriced.append(nid)
            result.notes.append("{} {}: {}".format(
                FITTING_LABELS.get(kind, kind), nid, why))
        result.fitting_psi += length * rate / 100.0
        result.developed_ft += length

    result.residual_psi = (result.start_psi - result.static_psi -
                           result.pipe_psi - result.fitting_psi -
                           result.device_psi)
    # Judged only when every loss on the way was priced: an unpriced pipe
    # reads as zero loss and could turn a fail into a pass.
    if (result.required_psi is not None and not result.unpriced and
            not result.behind_prv):
        result.margin_psi = result.residual_psi - result.required_psi
        result.passes = result.margin_psi >= 0.0
    return result


def _adjacent_rate(rates, next_id, prev_id):
    if next_id in rates:
        return rates[next_id]
    if prev_id in rates:
        return rates[prev_id]
    return 0.0


def _z(node):
    if node is None or not getattr(node, "location_xyz", None):
        return None
    return float(node.location_xyz[2])


def _connector_z(node, system):
    for c in node.connectors or []:
        if c.get("system_type") == system and c.get("origin_xyz"):
            return float(c["origin_xyz"][2])
    return None


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

TABLE_COLUMNS = [("FIXTURE", 0.110), ("SYS", 0.040), ("REQ PSI", 0.060),
                 ("STATIC", 0.055), ("PIPE", 0.050), ("FITTINGS", 0.065),
                 ("DEVICES", 0.060), ("LEFT", 0.050), ("MARGIN", 0.060),
                 ("DEV LENGTH FT", 0.090), ("RESULT", 0.060)]


def table_rows(analysis):
    """Rows for the pyRevit and sheet tables, worst first."""
    rows = []
    for o in analysis["outlets"]:
        rows.append([
            o.fixture_name,
            "HW" if o.system == HOT else "CW",
            _f(o.required_psi),
            _f(o.static_psi),
            _f(o.pipe_psi),
            _f(o.fitting_psi),
            _f(o.device_psi),
            _f(o.residual_psi),
            _f(o.margin_psi),
            "{:.0f}".format(o.developed_ft),
            "-" if o.passes is None else ("PASS" if o.passes else "FAIL"),
        ])
    return rows


def basis_lines(start_psi):
    return [
        "Pressure loss basis: 2024 IPC Appendix E, Section E103.3 "
        "(segmented loss method), checked against Table 604.3.",
        "  Start pressure {:.1f} psi at the picked point (after the RPZ, or "
        "in the existing main at the tee).".format(float(start_psi)),
        "  Pipe friction: Darcy-Weisbach, Swamee-Jain friction factor, drawn "
        "copper roughness {} ft, water at {:.0f} F for hot and cold "
        "(conservative for hot water), ASTM B88 Type L inside "
        "diameters.".format(ROUGHNESS_FT, WATER_TEMP_F),
        "  Fittings and valves: equivalent length, Table E103.3(6), at the "
        "friction rate of the pipe they sit on.",
        "  Devices: Device_Pressure_Drop_PSI on the element when set.",
        "  Static head {:.3f} psi per ft of rise above the start point. "
        "Section 604.8: {:.0f} psi maximum static.".format(
            PSI_PER_FT_HEAD, MAX_STATIC_PSI),
        "  Flow: probable demand per pipe, Table E103.3(3), at the drawn "
        "size.",
    ]


def _f(value):
    return "-" if value is None else "{:.1f}".format(float(value))
