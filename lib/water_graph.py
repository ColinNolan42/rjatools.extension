# -*- coding: ascii -*-
# water_graph.py
# System-aware connector traversal for domestic water piping.
#
# Walks the domestic cold water network from the RPZ (backflow preventer),
# finds the water heaters on it, walks each heater's hot water tree, and
# assigns a WSFU demand to every pipe.
#
# WHY THIS IS NOT pipe_graph.py (gas):
#   1. Gas follows EVERY connector regardless of system. A water fixture
#      carries cold, hot and hot-return connectors, so an unfiltered walk
#      leaks one system into another. This module filters on the connector's
#      own PipeSystemType.
#   2. Gas treats "3 or more connectors" as a tee and "1 connector" as a
#      fixture. Water roles come from parameters and connector systems.
#   3. Gas detects elbows by the string "Elbow" in the family name. This
#      codebase's rule is no family or Type names in code, and water has no
#      elbow-equivalent length at this stage, so that is not carried over.
#   4. Gas carries one load number per node. Water carries cold, hot and total.
#
# pipe_graph.py is deliberately left untouched so the gas tools cannot regress.
#
# MODEL: every element is a vertex. The BFS from the origin produces a spanning
# tree with parent pointers, so "downstream of X" is simply the subtree rooted
# at X. Pipes are vertices too, because pipes are what gets sized.
#
# IronPython 2.7  -  Revit API via Autodesk.Revit.DB

import shared_params
import revit_helpers


# Node kinds
KIND_ORIGIN     = "origin"       # the picked RPZ
KIND_PIPE       = "pipe"
KIND_FITTING    = "fitting"
KIND_ACCESSORY  = "accessory"
KIND_FIXTURE    = "fixture"      # IS_WATER_FIXTURE = Yes
KIND_HEATER     = "heater"       # cold in + hot out, not a fixture
KIND_PUMP       = "pump"         # hot-only equipment, the recirculation pump
KIND_EQUIPMENT  = "equipment"
KIND_UNKNOWN    = "unknown"


class WaterNode(object):
    """One element in the water network."""

    def __init__(self, element_id, element, kind):
        self.element_id   = element_id
        self.element      = element
        self.kind         = kind
        self.family_name  = _get_family_name(element)
        self.location_xyz = revit_helpers.get_element_location(element)

        self.connectors      = []
        self.connector_count = 0

        # Tree structure, PER SYSTEM. A fixture with both a cold and a hot
        # connector legitimately belongs to the cold tree AND to a heater's hot
        # tree, so a single parent/children pair cannot describe it. Keyed by
        # the system name, e.g. "DomesticColdWater".
        self.parent_by_system   = {}    # system -> parent element id
        self.children_by_system = {}    # system -> [child element id, ...]
        self.system             = None  # the system this node was FIRST reached on;
                                        # for a pipe this is its own system

        # Pipe geometry (only meaningful when kind == KIND_PIPE)
        self.length_feet     = 0.0
        self.diameter_inches = 0.0
        self.system_type_id   = None   # PipingSystemType ELEMENT id
        self.system_type_name = None   # its name, for display only

        # Fixture load (only meaningful when kind == KIND_FIXTURE)
        self.is_water_fixture = False
        self.fixture_name     = ""
        self.type_name        = _get_type_name(element)   # groups the WSFU table
        self.is_public        = False
        self.cw_wsfu          = 0.0
        self.hw_wsfu          = 0.0
        self.total_wsfu       = 0.0
        self.hwr_active       = False
        self.has_cw           = True
        self.has_hw           = True

        # Heater bookkeeping (only meaningful when kind == KIND_HEATER)
        self.served_fixture_ids = []

        # Sizing results, filled by the sizing engine
        self.demand_wsfu   = 0.0
        self.assigned_size = None


class WaterGraph(object):
    """The traversed water network."""

    def __init__(self):
        self.nodes        = {}     # element_id -> WaterNode
        self.origin_id    = None
        self.log_lines    = []
        self.open_ends    = []     # (element_id, system) with an unconnected connector
        self.loops        = []     # element ids reached more than once
        self.heater_ids   = []
        self.pump_ids     = []
        self.fixture_ids  = []
        self.unreadable   = []     # elements whose connectors could not be read
        self.system_types = {}     # PipingSystemType element id -> name

    def add(self, node):
        self.nodes[node.element_id] = node

    def log(self, message):
        self.log_lines.append(message)

    def descendants(self, node_id, system):
        """Every element id below node_id on `system`'s tree, excluding itself."""
        if node_id not in self.nodes:
            return []
        found = []
        seen = set()
        stack = list(self.nodes[node_id].children_by_system.get(system, []))
        while stack:
            nid = stack.pop()
            if nid in seen:
                continue
            seen.add(nid)
            found.append(nid)
            if nid in self.nodes:
                stack.extend(self.nodes[nid].children_by_system.get(system, []))
        return found


# =============================================================================
# ENTRY POINT
# =============================================================================

def build_water_network(origin_element, doc):
    """Traverse the domestic water network from the RPZ.

    Walks the cold tree from origin_element, finds the water heaters on it,
    then walks each heater's hot tree. Assigns WSFU demand to every pipe.

    Args:
        origin_element: the user-picked RPZ / backflow preventer.
        doc: the active Revit Document.

    Returns:
        WaterGraph
    """
    graph = WaterGraph()

    origin_id = revit_helpers.eid_int(origin_element.Id)
    graph.origin_id = origin_id

    origin = WaterNode(origin_id, origin_element, KIND_ORIGIN)
    origin.connectors = revit_helpers.get_connectors(origin_element)
    origin.connector_count = len(origin.connectors)
    origin.system = shared_params.SYSTEM_DOMESTIC_COLD_WATER
    graph.add(origin)
    graph.log("ORIGIN: element {} ({}), {} connectors".format(
        origin_id, origin.family_name, origin.connector_count))

    # --- Phase 1: the cold tree ---
    # Each system gets its OWN visited set. Sharing one set across systems
    # would make a fixture that is already on the cold tree look like a loop
    # when the hot walk reaches it, so hot pipes would carry no demand and
    # heaters would appear to serve nothing.
    cold = shared_params.SYSTEM_DOMESTIC_COLD_WATER
    hot = shared_params.SYSTEM_DOMESTIC_HOT_WATER

    visited_cold = set([origin_id])
    _walk(graph, doc, origin_id, cold, visited_cold)
    graph.log("COLD TREE: {} elements, {} fixtures, {} heaters".format(
        len(graph.nodes), len(graph.fixture_ids), len(graph.heater_ids)))

    # --- Phase 2: a hot tree per heater ---
    # Needed even for a cold-only run: a cold branch that feeds only the water
    # heater has no fixtures downstream of it on the cold tree, so without
    # knowing what the heater serves that branch would size at zero.
    # One shared hot visited set across heaters, so two heaters cross-connected
    # on the same hot main do not each claim the whole main.
    visited_hot = set()
    for heater_id in list(graph.heater_ids):
        if heater_id in visited_hot:
            graph.log("HEATER {}: already reached on another heater's hot tree, "
                      "not walked again.".format(heater_id))
            continue
        visited_hot.add(heater_id)
        _walk(graph, doc, heater_id, hot, visited_hot)
        served = [nid for nid in graph.descendants(heater_id, hot)
                  if graph.nodes[nid].kind == KIND_FIXTURE]
        graph.nodes[heater_id].served_fixture_ids = served
        graph.log("HEATER {}: serves {} fixtures on its hot tree.".format(
            heater_id, len(served)))

    # --- Phase 3: demand ---
    _assign_demand(graph)

    graph.log("TRAVERSAL COMPLETE: {} elements, {} open ends, {} loop hits".format(
        len(graph.nodes), len(graph.open_ends), len(graph.loops)))
    return graph


# =============================================================================
# TRAVERSAL
# =============================================================================

def _walk(graph, doc, start_id, system, visited):
    """Breadth-first walk from start_id following only `system` connectors."""
    queue = [start_id]

    while queue:
        current_id = queue.pop(0)
        current_node = graph.nodes.get(current_id)
        if current_node is None:
            continue

        connectors = current_node.connectors
        if not connectors:
            graph.unreadable.append(current_id)
            graph.log("NO CONNECTORS on element {} ({})".format(
                current_id, current_node.family_name))
            continue

        for c in connectors:
            # System filter. A connector whose system could not be read is
            # skipped rather than guessed at, and is reported.
            c_system = c.get("system_type")
            if c_system is None:
                graph.log("Element {}: connector {} has no readable system "
                          "type, skipped.".format(current_id, c.get("connector_index")))
                continue
            if c_system != system:
                continue

            # A fixture connector the Type does not actually use (its stub is
            # hidden) still exists in the family. Ignore it, and never call it
            # an open end.
            if _connector_is_disabled(current_node, c_system):
                continue

            if not c.get("is_connected"):
                graph.open_ends.append((current_id, system))
                graph.log("OPEN END: element {} has an unconnected {} "
                          "connector.".format(current_id, system))
                continue

            next_id = c.get("connected_element_id")
            if next_id is None:
                # Connected, but only to an MEPSystem object. Not a physical
                # neighbour, so there is nothing to walk into.
                graph.log("Element {}: {} connector references a system object "
                          "only, no physical neighbour.".format(current_id, system))
                continue

            if next_id in visited:
                if next_id != current_node.parent_by_system.get(system):
                    graph.loops.append(next_id)
                    graph.log("LOOP: element {} reaches already-visited element "
                              "{} on {}.".format(current_id, next_id, system))
                continue

            next_element = _get_element(doc, next_id)
            if next_element is None:
                graph.log("WARNING: could not retrieve element {} connected to "
                          "{}.".format(next_id, current_id))
                continue

            visited.add(next_id)

            # The element may already exist as a node from another system's
            # walk (a fixture with both cold and hot connectors). Reuse it and
            # just attach it to this system's tree, rather than rebuilding it.
            node = graph.nodes.get(next_id)
            if node is None:
                node = _make_node(graph, next_element, next_id, system)
                node.system = system
                graph.add(node)

            node.parent_by_system[system] = current_id
            current_node.children_by_system.setdefault(system, []).append(next_id)
            queue.append(next_id)


def _make_node(graph, element, element_id, system):
    """Build a WaterNode and classify its role from parameters and connectors."""
    connectors = revit_helpers.get_connectors(element)

    if _is_pipe(element):
        node = WaterNode(element_id, element, KIND_PIPE)
        node.connectors = connectors
        node.connector_count = len(connectors)
        node.length_feet = revit_helpers.get_pipe_length_feet(element) or 0.0
        node.diameter_inches = revit_helpers.get_pipe_diameter_inches(element) or 0.0
        _read_pipe_system_type(graph, element, node)
        graph.log("PIPE {}: {:.2f} ft, {:.3f} in, system '{}'".format(
            element_id, node.length_feet, node.diameter_inches,
            node.system_type_name))
        return node

    is_fixture = revit_helpers.get_parameter_value(
        element, shared_params.PARAM_IS_WATER_FIXTURE)

    if is_fixture:
        node = WaterNode(element_id, element, KIND_FIXTURE)
        node.connectors = connectors
        node.connector_count = len(connectors)
        node.is_water_fixture = True
        node.fixture_name = revit_helpers.get_parameter_value(
            element, shared_params.PARAM_FIXTURE_NAME) or "UNNAMED"
        node.cw_wsfu = _as_float(revit_helpers.get_parameter_value(
            element, shared_params.PARAM_CW_FIXTURE_UNITS))
        node.hw_wsfu = _as_float(revit_helpers.get_parameter_value(
            element, shared_params.PARAM_HW_FIXTURE_UNITS))
        node.total_wsfu = _as_float(revit_helpers.get_parameter_value(
            element, shared_params.PARAM_TOTAL_FIXTURE_UNITS))
        node.hwr_active = bool(revit_helpers.get_parameter_value(
            element, shared_params.PARAM_HWR_ACTIVE))
        node.is_public = bool(revit_helpers.get_parameter_value(
            element, shared_params.PARAM_IS_PUBLIC_OCCUPANCY))
        has_cw = revit_helpers.get_type_parameter_value(
            element, shared_params.PARAM_HAS_CW)
        has_hw = revit_helpers.get_type_parameter_value(
            element, shared_params.PARAM_HAS_HW)
        node.has_cw = True if has_cw is None else bool(has_cw)
        node.has_hw = True if has_hw is None else bool(has_hw)
        graph.fixture_ids.append(element_id)
        graph.log("FIXTURE {}: '{}' cold={} hot={} total={} hwr={}".format(
            element_id, node.fixture_name, node.cw_wsfu, node.hw_wsfu,
            node.total_wsfu, node.hwr_active))
        return node

    if _looks_like_heater(connectors):
        node = WaterNode(element_id, element, KIND_HEATER)
        node.connectors = connectors
        node.connector_count = len(connectors)
        graph.heater_ids.append(element_id)
        graph.log("WATER HEATER {}: '{}' (cold in + hot out)".format(
            element_id, node.family_name))
        return node

    if _looks_like_pump(element, connectors):
        node = WaterNode(element_id, element, KIND_PUMP)
        node.connectors = connectors
        node.connector_count = len(connectors)
        graph.pump_ids.append(element_id)
        graph.log("PUMP CANDIDATE {}: '{}' (hot-only equipment, {} "
                  "connectors)".format(element_id, node.family_name,
                                       len(connectors)))
        return node

    kind = _classify_by_category(element, len(connectors))
    node = WaterNode(element_id, element, kind)
    node.connectors = connectors
    node.connector_count = len(connectors)
    graph.log("{} {}: '{}', {} connectors".format(
        kind.upper(), element_id, node.family_name, len(connectors)))
    return node


def _read_pipe_system_type(graph, pipe, node):
    """Record the pipe's PipingSystemType ELEMENT, not just the enum.

    The enum cannot separate hot water supply from hot water recirculation:
    both report DomesticHotWater. The project's PipingSystemType element does
    separate them, so it is recorded here for the caller to map. Names are
    reported, never matched against in code.
    """
    try:
        type_id = pipe.MEPSystem.GetTypeId()
        system_type = pipe.Document.GetElement(type_id)
        if system_type is not None:
            node.system_type_id = revit_helpers.eid_int(system_type.Id)
            node.system_type_name = system_type.Name
            graph.system_types[node.system_type_id] = node.system_type_name
            return
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter
        param = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
        if param is not None:
            node.system_type_name = param.AsValueString()
    except Exception:
        node.system_type_name = None


def _connector_is_disabled(node, system):
    """True if this fixture Type does not actually use that system.

    The family keeps all three connectors on every Type and hides the unused
    stubs, so an unused connector is present but meaningless. Reporting it as
    an open end would be a false error.
    """
    if node.kind != KIND_FIXTURE:
        return False
    if system == shared_params.SYSTEM_DOMESTIC_COLD_WATER:
        return not node.has_cw
    if system == shared_params.SYSTEM_DOMESTIC_HOT_WATER:
        return not node.has_hw
    return False


# =============================================================================
# DEMAND
# =============================================================================

def _assign_demand(graph):
    """Set demand_wsfu on every pipe.

    Accounting rule (Colin, 2026-09-24): cold water piping is sized on TOTAL
    fixture units, hot water piping on HOT fixture units, and the cold branch
    feeding a water heater is sized on the HOT total of what that heater
    serves.

    So for a cold pipe, each fixture downstream is counted exactly ONCE:
      - reached directly on the cold tree  -> its TOTAL value
      - reached only through a water heater -> its HOT value, because all that
        branch carries is the heater's make-up
      - reached both ways -> TOTAL, which already covers the hot portion

    Counting TOTAL on the cold side and HOT again through the heater would
    double-count: Table E103.3(2) note a already makes the separate hot and
    cold loads three-fourths of the total.
    """
    for node in graph.nodes.values():
        if node.kind != KIND_PIPE:
            continue

        downstream = graph.descendants(node.element_id, node.system)

        if node.system == shared_params.SYSTEM_DOMESTIC_HOT_WATER:
            total = 0.0
            for nid in downstream:
                child = graph.nodes[nid]
                if child.kind == KIND_FIXTURE:
                    total += child.hw_wsfu
            node.demand_wsfu = total
            continue

        # Cold pipe. Split the downstream fixtures into those reached directly
        # and those reached only through a heater, so each is counted once and
        # a heater-only branch carries hot load rather than total load.
        direct = set()
        via_heater = set()
        for nid in downstream:
            child = graph.nodes[nid]
            if child.kind == KIND_FIXTURE:
                direct.add(nid)
            elif child.kind == KIND_HEATER:
                for fid in child.served_fixture_ids:
                    via_heater.add(fid)

        total = 0.0
        for fid in direct:
            total += graph.nodes[fid].total_wsfu
        for fid in via_heater:
            if fid not in direct:
                total += graph.nodes[fid].hw_wsfu
        node.demand_wsfu = total


# =============================================================================
# CLASSIFICATION HELPERS
# =============================================================================

def _looks_like_heater(connectors):
    """A water heater has a cold water connector and a hot water connector.

    Identified by connector systems only, never by family or Type name. A
    heater may carry extra connectors (relief, recirculation return, and some
    that throw on every property read), so this only checks that both systems
    are present.
    """
    has_cold = False
    has_hot = False
    for c in connectors:
        system = c.get("system_type")
        if system == shared_params.SYSTEM_DOMESTIC_COLD_WATER:
            has_cold = True
        elif system == shared_params.SYSTEM_DOMESTIC_HOT_WATER:
            has_hot = True
    return has_cold and has_hot


def _looks_like_pump(element, connectors):
    """A recirculation pump: Mechanical Equipment, hot water only.

    Identified by category and connector systems, never by family or Type
    name. The discriminators are that a pump sits on the hot side with NO cold
    water connector (which is what separates it from the water heater), and
    that it is Mechanical Equipment (which is what separates it from an inline
    balancing valve or circuit setter, which are Pipe Accessories).

    This is a CANDIDATE, not a certainty: any other hot-water mechanical
    equipment with two connectors matches too. The checks report the family
    name so a human can confirm, rather than asserting what it is.

    UNVERIFIED: the one circulation pump seen live (a Grundfos-style
    ecocircXL) threw on every connector property read in a probe that guarded
    only the outer loop. revit_helpers.get_connectors guards each property
    separately, so some fields may still come back, but whether system_type is
    readable on that family has NOT been confirmed. If it is not, the pump
    cannot be classified here and will be reported as equipment with
    unreadable connectors instead of being silently missed.
    """
    try:
        category = element.Category.Name
    except Exception:
        return False
    if "Equipment" not in category:
        return False

    hot = 0
    for c in connectors:
        system = c.get("system_type")
        if system == shared_params.SYSTEM_DOMESTIC_COLD_WATER:
            return False
        if system == shared_params.SYSTEM_DOMESTIC_HOT_WATER:
            hot += 1
    return hot >= 2


def _classify_by_category(element, connector_count):
    """Fall back to the Revit category for anything not a pipe/fixture/heater."""
    try:
        name = element.Category.Name
    except Exception:
        return KIND_UNKNOWN
    if "Fitting" in name:
        return KIND_FITTING
    if "Accessor" in name:
        return KIND_ACCESSORY
    if "Equipment" in name:
        return KIND_EQUIPMENT
    return KIND_UNKNOWN


def _as_float(value):
    if value is None:
        return 0.0
    try:
        return float(value)
    except Exception:
        return 0.0


def _get_element(doc, element_id_int):
    """Fetch an element by integer id, version-safely."""
    from Autodesk.Revit.DB import ElementId
    try:
        return doc.GetElement(ElementId(element_id_int))
    except Exception:
        return None


def _is_pipe(element):
    try:
        return element.GetType().Name == "Pipe"
    except Exception:
        return False


def _get_family_name(element):
    """Family name, via the shared helper.

    Never read .Name directly here: it throws under pyRevit's IronPython even
    though it works in C#. revit_helpers handles the fallbacks.
    """
    name = revit_helpers.get_element_family_name(element)
    if name:
        return name
    try:
        return element.Name
    except Exception:
        return "UNKNOWN"


def _get_type_name(element):
    """The family Type name, which is what the WSFU schedule groups by.

    Getting this wrong is not cosmetic: the take-off groups rows by Type, so a
    failed lookup collapses every fixture into one row.
    """
    return revit_helpers.get_element_type_name(element) or "UNKNOWN TYPE"
