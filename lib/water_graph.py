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
import tie_in
import water_tables


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

# Tree key for the hot water RETURN piping. The return shares the hot water
# system CLASSIFICATION, so it cannot have a tree of its own until the return
# System Types are known. apply_return_system_types() builds it, rooted where
# the return ends (the water heater) with children pointing back toward the
# fixtures, so "everything below a return pipe" is everything that drains
# through it.
TREE_RETURN = "HotWaterReturn"

# What the return may pass through on its way back. A fixture or a heater ends
# a return line, it never sits in the middle of one.
_RETURN_PASSABLE_KINDS = (KIND_FITTING, KIND_ACCESSORY, KIND_PUMP,
                          KIND_EQUIPMENT, KIND_UNKNOWN)


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
        self.is_existing      = False  # created before the project's last phase

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
        self.is_flush_valve   = False   # flushometer valve water closet or urinal
        self.fixed_gpm        = 0.0     # fixed cold gpm load, a hose bibb (Table 604.3)

        # Heater bookkeeping (only meaningful when kind == KIND_HEATER)
        self.served_fixture_ids = []

        # Hot pipe between water heaters in parallel and the tee where their
        # outlets have all combined: the heater group tuple, else False (set
        # by _assign_parallel_hot_demand)
        self.parallel_leg = False

        # Sizing results, filled by the sizing engine
        self.demand_wsfu   = 0.0
        self.demand_fixed_gpm = 0.0   # fixed gpm loads downstream (cold pipes)
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
        self.loop_edges   = []     # (from id, to id, system) joints that closed a loop
        self.return_system_type_ids = set()   # set by apply_return_system_types
        self.return_roots    = []  # where each return line ends (heater, or the cold tie-in)
        self.return_unrooted = []  # return pipe ids whose end could not be found
        self.hot_joints      = {}  # every hot water joint, kept from before the split
        self.heater_groups   = {}  # heater id -> tuple of heaters in parallel

        # Existing building: the run starts at the first NEW cold water pipe
        # instead of an RPZ (Colin, 2026-10-13). The connectors listed here
        # face the existing piping and are never walked through.
        self.origin_is_new_pipe         = False   # any existing building tie-in
        self.origin_is_tee              = False   # ... picked as the tee
        self.origin_upstream_connectors = set()   # connector_index values
        self.origin_upstream_reason     = ""
        self.origin_phase_id            = None    # created phase of the picked pipe
        self.existing_connections       = []      # (new element id, existing element id)
        self.existing_passthrough       = []      # existing devices inside the new run
        self.last_phase_id              = None    # the project's last phase (new work)
        self.upstream_ids               = set()   # service piping behind the RPZ inlet
        self.cross_ties                 = []      # (from id, pipe id, pipe system, walked system)
        self.parallel_backflow          = []      # other backflow preventers in parallel with the picked RPZ
        self.unrecognized_ids           = []      # on the piping, not a fixture, heater, pump or device

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

def build_water_network(origin_element, doc, upstream=None):
    """Traverse the domestic water network from the RPZ, or from the first
    new cold water pipe on an existing building.

    Walks the cold tree from origin_element, finds the water heaters on it,
    then walks each heater's hot tree. Assigns WSFU demand to every pipe.

    Args:
        origin_element: the user-picked RPZ / backflow preventer, or the first
            new construction cold water pipe.
        doc: the active Revit Document.
        upstream: for a picked pipe, what tie_in.resolve_new_pipe_upstream()
            returned for its cold water connectors.
            Worked out here when not given.

    Returns:
        WaterGraph

    Raises:
        ValueError: a pipe was picked and its upstream end cannot be told.
    """
    graph = WaterGraph()
    graph.last_phase_id = revit_helpers.get_last_phase_id(doc)

    origin_id = revit_helpers.eid_int(origin_element.Id)
    graph.origin_id = origin_id

    if _is_pipe(origin_element):
        # The picked pipe is new construction and has to be sized like any
        # other, so it goes in as a PIPE. Everything it feeds hangs below it,
        # and the end facing the existing piping is closed to the walk.
        if upstream is None:
            upstream = tie_in.resolve_new_pipe_upstream(
                origin_element, doc,
                shared_params.SYSTEM_DOMESTIC_COLD_WATER)
        if upstream.get("status") != "ok":
            raise ValueError(
                "Could not tell which end of the picked pipe connects to "
                "existing: {}.".format(upstream.get("reason", "unknown")))
        origin = _make_node(graph, origin_element, origin_id,
                            shared_params.SYSTEM_DOMESTIC_COLD_WATER)
        origin.system = shared_params.SYSTEM_DOMESTIC_COLD_WATER
        graph.add(origin)
        graph.origin_is_new_pipe = True
        graph.origin_phase_id = revit_helpers.get_created_phase_id(
            origin_element)
        graph.origin_upstream_reason = upstream.get("reason", "")
        graph.origin_upstream_connectors = set([upstream["connector_index"]])
        graph.log("ORIGIN: new construction pipe {}. Upstream end: {}".format(
            origin_id, graph.origin_upstream_reason or "NOT RESOLVED"))
    else:
        origin = WaterNode(origin_id, origin_element, KIND_ORIGIN)
        origin.connectors = revit_helpers.get_connectors(origin_element)
        origin.connector_count = len(origin.connectors)
        origin.system = shared_params.SYSTEM_DOMESTIC_COLD_WATER
        graph.add(origin)
        graph.log("ORIGIN: element {} ({}), {} connectors".format(
            origin_id, origin.family_name, origin.connector_count))
        if upstream is not None and upstream.get("tee"):
            # A tee on the existing main (tie_in.resolve_tee_branch). Only its
            # branch is walked, and the walk stops at anything from an earlier
            # phase, exactly as for a picked new pipe.
            if upstream.get("status") != "ok":
                raise ValueError(
                    "Could not find the new branch of the picked tee: "
                    "{}.".format(upstream.get("reason", "unknown")))
            graph.origin_is_new_pipe = True
            graph.origin_is_tee = True
            graph.origin_phase_id = upstream.get("new_phase_id")
            graph.origin_upstream_reason = upstream.get("reason", "")
            graph.origin_upstream_connectors = set(
                upstream["upstream_connectors"])
            graph.log("ORIGIN: tee {} on the existing main. {}".format(
                origin_id, graph.origin_upstream_reason))

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

    # --- Service side of the RPZ ---
    if origin.kind == KIND_ORIGIN and not graph.origin_is_new_pipe:
        _mark_upstream_service(graph, cold)

    # --- Phase 3: demand ---
    _assign_demand(graph)

    graph.log("TRAVERSAL COMPLETE: {} elements, {} open ends, {} loop hits".format(
        len(graph.nodes), len(graph.open_ends), len(graph.loops)))
    return graph


def _mark_upstream_service(graph, cold):
    """Mark the piping behind the RPZ inlet so it is never sized.

    The RPZ is the start of the scope; the service line upstream may run at a
    smaller size on purpose (a higher pressure, with a PRV beyond it), so it is
    none of this run's business (Colin, 2026-10-26). Both RPZ connectors are
    cold water, so the walk crosses to the service side. A branch of the origin
    with no fixture and no water heater anywhere below it is that side. Its
    open end is not reported as an open end.
    """
    origin = graph.nodes.get(graph.origin_id)
    if origin is None:
        return
    for child_id in origin.children_by_system.get(cold, []):
        below = [child_id] + list(graph.descendants(child_id, cold))
        if any(graph.nodes[nid].kind in (KIND_FIXTURE, KIND_HEATER)
               for nid in below if nid in graph.nodes):
            continue
        graph.upstream_ids.update(below)
        graph.log("UPSTREAM SERVICE: {} element(s) behind the RPZ inlet "
                  "starting at {} are outside the scope and not sized.".format(
                      len(below), child_id))
    if graph.upstream_ids:
        graph.open_ends = [(eid, system) for eid, system in graph.open_ends
                           if eid not in graph.upstream_ids]


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
            if (current_node.kind == KIND_PUMP and c_system != system and
                    _pump_connector_continues(doc, c, system)):
                # A circulation pump does not report its own connectors on the
                # system it sits on (see _looks_like_pump), so the walk used to
                # stop dead at it and the piping beyond was never reached.
                c_system = system
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

            # The end of the picked new pipe that faces the existing piping.
            # Not an open end, and never walked: the existing main is assumed
            # adequate and is none of this run's business.
            if (current_id == graph.origin_id and
                    c.get("connector_index") in
                    graph.origin_upstream_connectors):
                continue

            if not c.get("is_connected"):
                if _unused_return_outlet(current_node, c, system):
                    # HW Out of a fixture that is not on recirculation. Left
                    # open by design, so it is not an open end.
                    continue
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
                    # Keep the joint itself. The spanning tree drops it, but
                    # it is a real connection, and the return piping can only
                    # be laid out if every joint is known.
                    graph.loop_edges.append((current_id, next_id, system))
                    graph.log("LOOP: element {} reaches already-visited element "
                              "{} on {}.".format(current_id, next_id, system))
                continue

            next_element = _get_element(doc, next_id)
            if next_element is None:
                graph.log("WARNING: could not retrieve element {} connected to "
                          "{}.".format(next_id, current_id))
                continue

            # A run that started at a new pipe sizes new construction only.
            # Anything from another phase is existing: the connection to it
            # is assumed adequate for the new load, and the walk stops there
            # (Colin, 2026-10-13).
            if graph.origin_is_new_pipe and graph.origin_phase_id is not None:
                phase = revit_helpers.get_created_phase_id(next_element)
                if (phase is not None and phase != graph.origin_phase_id and
                        _existing_inside_new_run(graph, doc, next_element,
                                                 current_id)):
                    # An existing valve or fitting with new construction on
                    # every other side is part of the new run: walk through
                    # it and say so, rather than silently cutting the run.
                    graph.existing_passthrough.append(next_id)
                    graph.log("EXISTING IN NEW RUN: element {} is existing "
                              "phase but only new construction is beyond "
                              "it. Walked through.".format(next_id))
                elif phase is not None and phase != graph.origin_phase_id:
                    visited.add(next_id)
                    graph.existing_connections.append((current_id, next_id))
                    graph.log("EXISTING: element {} connects to existing "
                              "phase element {} on {}. Not followed.".format(
                                  current_id, next_id, system))
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

            if node.kind == KIND_PIPE:
                # A pipe of the other water system reached through a fitting
                # is a hot to cold tie. The only legal one is the return
                # tying into the heater cold supply, which water_checks
                # clears by System Type once the return is known.
                own = None
                for connector in node.connectors:
                    own = connector.get("system_type")
                    if own:
                        break
                if own in (shared_params.SYSTEM_DOMESTIC_COLD_WATER,
                           shared_params.SYSTEM_DOMESTIC_HOT_WATER)                         and own != system:
                    graph.cross_ties.append((current_id, next_id, own, system))
                    graph.log("HOT/COLD TIE: pipe {} is {} but was reached "
                              "from {} on {}.".format(
                                  next_id, own, current_id, system))

            node.parent_by_system[system] = current_id
            current_node.children_by_system.setdefault(system, []).append(next_id)
            queue.append(next_id)


def _pump_connector_continues(doc, connector, system):
    """True if a pump connector that does not report `system` leads on to it.

    Only an unreadable or OtherPipe connector is considered, and only when the
    element on the other side carries a connector that IS on `system`. That
    keeps the walk from leaking into another system through a piece of
    equipment that merely looks like a pump.
    """
    if connector.get("system_type") not in (None,
                                            shared_params.SYSTEM_OTHER_PIPE):
        return False
    if not connector.get("is_connected"):
        return False
    next_id = connector.get("connected_element_id")
    if next_id is None:
        return False
    neighbour = _get_element(doc, next_id)
    if neighbour is None:
        return False
    for c in revit_helpers.get_connectors(neighbour):
        if c.get("system_type") == system:
            return True
    return False


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
        if graph.last_phase_id is not None:
            phase = revit_helpers.get_created_phase_id(element)
            node.is_existing = (phase is not None and
                                phase != graph.last_phase_id)
        graph.log("PIPE {}: {:.2f} ft, {:.3f} in, system '{}'{}".format(
            element_id, node.length_feet, node.diameter_inches,
            node.system_type_name,
            ", EXISTING phase" if node.is_existing else ""))
        return node

    is_fixture = revit_helpers.get_parameter_value(
        element, shared_params.PARAM_IS_WATER_FIXTURE)

    if is_fixture:
        node = WaterNode(element_id, element, KIND_FIXTURE)
        node.connectors = connectors
        node.connector_count = len(connectors)
        node.is_water_fixture = True
        node.fixture_name = revit_helpers.get_parameter_value(
            element, shared_params.PARAM_FIXTURE_NAME) or \
            node.type_name or "UNNAMED"
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
        node.is_flush_valve = _is_flush_valve(element, node)
        graph.fixture_ids.append(element_id)
        graph.log("FIXTURE {}: '{}' cold={} hot={} total={} hwr={}".format(
            element_id, node.fixture_name, node.cw_wsfu, node.hw_wsfu,
            node.total_wsfu, node.hwr_active))
        return node

    if _name_is_hose_bib(_get_family_name(element), _get_type_name(element)):
        # A hose bibb carries no WSFU in Table E103.3(2); it is a fixed
        # 5 gpm load from Table 604.3 (Colin, 2026-10-26). Any other family
        # that is not marked as a fixture is flagged, not guessed at.
        node = WaterNode(element_id, element, KIND_FIXTURE)
        node.connectors = connectors
        node.connector_count = len(connectors)
        node.is_water_fixture = True
        node.fixture_name = "Hose bibb"
        node.fixed_gpm = water_tables.hose_bibb_gpm()
        node.has_hw = False
        graph.fixture_ids.append(element_id)
        graph.log("HOSE BIBB {}: fixed {} gpm".format(
            element_id, node.fixed_gpm))
        return node

    if _looks_like_heater(connectors):
        node = WaterNode(element_id, element, KIND_HEATER)
        node.connectors = connectors
        node.connector_count = len(connectors)
        graph.heater_ids.append(element_id)
        graph.log("WATER HEATER {}: '{}' (cold in + hot out)".format(
            element_id, node.family_name))
        return node

    if _looks_like_pump(element, connectors, system):
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
    if kind in (KIND_EQUIPMENT, KIND_UNKNOWN):
        graph.unrecognized_ids.append(element_id)
    graph.log("{} {}: '{}', {} connectors".format(
        kind.upper(), element_id, node.family_name, len(connectors)))
    return node


def _name_is_hose_bib(*names):
    """True when a family or Type name says hose bibb or sill cock."""
    for name in names:
        text = "".join(ch for ch in (name or "").lower() if ch.isalnum())
        if "hosebib" in text or "sillcock" in text:
            return True
    return False


def _read_pipe_system_type(graph, pipe, node):
    """Record the pipe's PipingSystemType ELEMENT, not just the enum.

    This is the distinction that matters. System CLASSIFICATION is
    DomesticHotWater for hot supply and hot water recirculation alike, so
    classification alone can never separate them. The System TYPE can: a
    project carries "Domestic Hot Water" and "Domestic Hot Water
    Recirculation" as two different PipingSystemType elements, with different
    ids and different names. That is what is recorded here, and what
    detect_return_system_types() reasons over.

    The id is taken FIRST and separately from the name, because reading a name
    can throw under IronPython and an exception while getting the name used to
    discard the id along with it, leaving the recirculation system
    indistinguishable for the rest of the run.
    """
    system_type = None
    try:
        system_type = pipe.Document.GetElement(pipe.MEPSystem.GetTypeId())
    except Exception:
        system_type = None

    if system_type is not None:
        try:
            node.system_type_id = revit_helpers.eid_int(system_type.Id)
        except Exception:
            node.system_type_id = None
        node.system_type_name = revit_helpers.get_element_type_name(system_type)

    if not node.system_type_name:
        # The pipe's own parameter gives the same name without touching the
        # type element, so it still works when the element read failed.
        try:
            from Autodesk.Revit.DB import BuiltInParameter
            param = pipe.get_Parameter(
                BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
            if param is not None:
                node.system_type_name = param.AsValueString()
        except Exception:
            pass

    if node.system_type_id is not None:
        graph.system_types[node.system_type_id] = (
            node.system_type_name or "(unnamed system type)")


# Words a recirculation system type is named with. Only ever used to
# CORROBORATE the topology below, and always reported to the user as a reason
# they can overrule, so a project that names its systems differently loses a
# hint and never gets a wrong answer.
_RETURN_NAME_HINTS = ("recirc", "re-circ", "return", "hwr", "hwc")


def detect_return_system_types(graph):
    """Work out which hot water System Types are the RETURN, from the model.

    Hot supply and hot water recirculation share one system CLASSIFICATION
    (DomesticHotWater), which is why classification cannot answer this. They
    do NOT share a System Type: they are two separate PipingSystemType
    elements. This reads those types off the pipes actually traversed and
    decides which is the return, rather than asking the user to.

    Evidence, strongest first. Everything found is reported, not just the
    winner, so the user can see why:

      1. A recirculation pump sits on it. The pump is found by topology, so
         this holds whatever anyone named anything.
      2. It returns to a water heater that another hot type feeds. A supply
         type leaves the heater; a return type arrives at one.
      3. Its name says so.
      4. It is the minority type: on a real job the return carries far fewer
         pipes than the supply (11 against 66 on Grantham 4).

    Signal 1 or 2 alone is enough to decide. Otherwise 3 decides, and 4 alone
    is treated as a suggestion rather than a decision, because a small supply
    branch on a job with no return would otherwise be mislabelled.

    Returns:
        dict with "detected" (set of system type ids), "unconfirmed" (the
        subset of detected picked on pipe count alone), "candidates" (list
        of dicts with id, name, pipe_count, reasons, detected, decisive) and
        "certain" (True when topology decided it).
    """
    hot = {}
    for node in graph.nodes.values():
        if node.kind != KIND_PIPE:
            continue
        if node.system != shared_params.SYSTEM_DOMESTIC_HOT_WATER:
            continue
        if node.system_type_id is None:
            continue
        entry = hot.setdefault(node.system_type_id, {
            "id": node.system_type_id,
            "name": node.system_type_name or "(unnamed system type)",
            "pipe_count": 0,
            "reasons": [],
            "detected": False,
            "decisive": False,
        })
        entry["pipe_count"] += 1

    if not hot:
        return {"detected": set(), "unconfirmed": set(), "candidates": [],
                "certain": False}

    # --- 1. the pump's own system type -------------------------------------
    for pump_id in graph.pump_ids:
        for type_id in _nearby_pipe_system_types(graph, pump_id):
            if type_id in hot:
                entry = hot[type_id]
                if not entry["decisive"]:
                    entry["reasons"].append(
                        "a recirculation pump sits on it")
                entry["detected"] = True
                entry["decisive"] = True

    # --- 2. arrives at a heater that a different hot type leaves -----------
    for heater_id in graph.heater_ids:
        arriving = _nearby_pipe_system_types(graph, heater_id)
        if len(arriving) > 1:
            # More than one hot type touches this heater. The supply is the
            # one the heater feeds outward, which is the type carried by the
            # heater's children; anything else at the heater comes back to it.
            # Looked up THROUGH fittings and valves: a heater outlet often
            # connects to a union or transition first, which has no System
            # Type, and stopping there made the supply look like it arrived.
            # Read off the heater's own hot OUTLET connectors, so a second
            # heater whose outlet tees into the first heater's riser (reached
            # by the walk from the supply side) still reads as supply.
            outgoing = _outlet_pipe_types(graph, heater_id)
            for type_id in arriving:
                if type_id in hot and type_id not in outgoing:
                    entry = hot[type_id]
                    if not entry["decisive"]:
                        entry["reasons"].append(
                            "it returns to a water heater rather than "
                            "leaving one")
                        entry["detected"] = True
                        entry["decisive"] = True

    # --- 3. the name --------------------------------------------------------
    for entry in hot.values():
        lowered = (entry["name"] or "").lower()
        for hint in _RETURN_NAME_HINTS:
            if hint in lowered:
                entry["reasons"].append(
                    "its System Type is named '{}'".format(entry["name"]))
                if not any(e["decisive"] for e in hot.values()):
                    entry["detected"] = True
                break

    # --- 4. the minority ----------------------------------------------------
    if len(hot) > 1:
        ordered = sorted(hot.values(),
                         key=lambda e: (e["pipe_count"], e["id"]))
        smallest, largest = ordered[0], ordered[-1]
        if smallest["pipe_count"] * 2 <= largest["pipe_count"]:
            smallest["reasons"].append(
                "it carries far fewer pipes than '{}' ({} against {})".format(
                    largest["name"], smallest["pipe_count"],
                    largest["pipe_count"]))
            if not any(e["detected"] for e in hot.values()):
                # Suggestion only. Nothing stronger spoke, so say so.
                smallest["detected"] = True
                smallest["unconfirmed"] = True
                smallest["reasons"].append(
                    "NOTHING CONFIRMED THIS, CHECK IT")

    candidates = sorted(hot.values(), key=lambda e: (-e["pipe_count"],
                                                     e["name"]))
    return {
        "detected": set(e["id"] for e in candidates if e["detected"]),
        # Picked on pipe count alone. Reported, but never acted on: the run
        # does not split the network or size anything on a guess.
        "unconfirmed": set(e["id"] for e in candidates
                           if e.get("unconfirmed")),
        "candidates": candidates,
        "certain": any(e["decisive"] for e in candidates),
    }


def _outlet_pipe_types(graph, heater_id, max_hops=4):
    """System Type ids of the first pipes on a heater's hot OUTLET connectors.

    Starts at the element each hot connector flowing Out is joined to (a
    connector whose direction could not be read also counts, since a hot
    return port on a heater flows In), then steps through fittings and
    valves, never back through the heater, to the first pipe on each path.
    """
    hot = shared_params.SYSTEM_DOMESTIC_HOT_WATER
    heater = graph.nodes.get(heater_id)
    if heater is None:
        return set()

    found = set()
    seen = set([heater_id])
    frontier = []
    for c in heater.connectors:
        if c.get("system_type") != hot or c.get("direction") == "In":
            continue
        if c.get("connected_element_id") is not None:
            frontier.append((c["connected_element_id"], 1))
    if not frontier:
        # No connector data: fall back to what hangs below the heater.
        frontier = [(child_id, 1) for child_id in
                    heater.children_by_system.get(hot, [])]
    while frontier:
        current_id, depth = frontier.pop()
        if current_id in seen:
            continue
        seen.add(current_id)
        node = graph.nodes.get(current_id)
        if node is None:
            continue
        if node.kind == KIND_PIPE:
            if node.system_type_id is not None:
                found.add(node.system_type_id)
            continue
        if depth >= max_hops or node.kind in (KIND_HEATER, KIND_FIXTURE,
                                              KIND_ORIGIN):
            continue
        neighbours = list(node.children_by_system.get(hot, []))
        if hot in node.parent_by_system:
            neighbours.append(node.parent_by_system[hot])
        for next_id in neighbours:
            frontier.append((next_id, depth + 1))
    return found


def _nearby_pipe_system_types(graph, node_id, max_hops=3):
    """System Type ids of the pipes within a few hops of an element.

    A pump or heater connects through fittings, so its own neighbours are
    often not pipes. Only a pipe carries a System Type, so the walk steps
    past the fittings to reach one.
    """
    found = set()
    seen = set([node_id])
    frontier = [(node_id, 0)]
    while frontier:
        current_id, depth = frontier.pop()
        if depth >= max_hops:
            continue
        node = graph.nodes.get(current_id)
        if node is None:
            continue
        neighbours = []
        for system, parent_id in node.parent_by_system.items():
            neighbours.append(parent_id)
        for system, child_ids in node.children_by_system.items():
            neighbours.extend(child_ids)
        for neighbour_id in neighbours:
            if neighbour_id in seen:
                continue
            seen.add(neighbour_id)
            neighbour = graph.nodes.get(neighbour_id)
            if neighbour is None:
                continue
            if (neighbour.kind == KIND_PIPE and
                    neighbour.system_type_id is not None):
                found.add(neighbour.system_type_id)
                continue    # a pipe ends the search down this leg
            if neighbour.kind in (KIND_HEATER, KIND_FIXTURE, KIND_ORIGIN):
                # Never look THROUGH a heater or a fixture. A pump piped
                # close to the heater would otherwise see the supply main on
                # the far side of the heater and call the supply the return.
                continue
            frontier.append((neighbour_id, depth + 1))
    return found


# =============================================================================
# RETURN PIPING
# =============================================================================

def hot_adjacency(graph):
    """Every hot water joint, both ways: element id -> [neighbour ids].

    The hot walk leaves a spanning tree, and a spanning tree of a
    recirculation loop is missing the joint that closed it. This puts the tree
    links and the loop-closing joints back together.
    """
    hot = shared_params.SYSTEM_DOMESTIC_HOT_WATER
    adjacency = {}

    def _join(a, b):
        if a not in graph.nodes or b not in graph.nodes:
            return
        for one, other in ((a, b), (b, a)):
            neighbours = adjacency.setdefault(one, [])
            if other not in neighbours:
                neighbours.append(other)

    for node_id in sorted(graph.nodes.keys()):
        node = graph.nodes[node_id]
        for child_id in node.children_by_system.get(hot, []):
            _join(node_id, child_id)
    for a, b, system in graph.loop_edges:
        if system == hot:
            _join(a, b)
    return adjacency


def is_return_pipe(graph, node):
    """A pipe whose System Type was identified as the hot water return."""
    return (node is not None and node.kind == KIND_PIPE and
            node.system_type_id is not None and
            node.system_type_id in graph.return_system_type_ids)


def apply_return_system_types(graph, return_system_type_ids):
    """Split the hot water network into supply and return, then redo demand.

    The hot walk follows every hot water connector, and the return shares the
    hot water classification, so on a recirculating system it walks round the
    loop from both ends. Where the two fronts meet is decided by element
    count, not by flow, so fixtures near the end of the main can hang off the
    return side of the tree and drop out of the supply main's fixture units.

    Once the return System Types are known this fixes it:
      1. the hot SUPPLY tree is rebuilt from each heater without ever stepping
         into a return pipe, so every hot fixture is counted on the supply;
      2. the RETURN piping gets its own tree (TREE_RETURN), rooted where it
         ends, so circulation flow can be accumulated toward the heater;
      3. demand is reassigned on the corrected trees.

    Does nothing when no return System Type was identified.

    Returns:
        the hot adjacency used, or None when there was nothing to do.
    """
    returns = set(return_system_type_ids or [])
    graph.return_roots = []
    graph.return_unrooted = []

    # A return needs a supply to return from. If every hot pipe carries a
    # return System Type the detection is wrong, and acting on it would leave
    # no supply tree at all, so nothing is split.
    hot_types = set(
        node.system_type_id for node in graph.nodes.values()
        if node.kind == KIND_PIPE and
        node.system == shared_params.SYSTEM_DOMESTIC_HOT_WATER and
        node.system_type_id is not None)
    if returns and not (hot_types - returns):
        graph.log("RETURN SPLIT REFUSED: every hot water System Type was "
                  "identified as the return, which cannot be right.")
        returns = set()

    graph.return_system_type_ids = returns
    if not returns:
        return None

    hot = shared_params.SYSTEM_DOMESTIC_HOT_WATER
    # Taken once. After the split the hot tree no longer holds the return
    # joints, so a second call must reuse what the first one read.
    adjacency = graph.hot_joints or hot_adjacency(graph)
    graph.hot_joints = adjacency

    for node in graph.nodes.values():
        node.parent_by_system.pop(hot, None)
        node.children_by_system.pop(hot, None)
        node.parent_by_system.pop(TREE_RETURN, None)
        node.children_by_system.pop(TREE_RETURN, None)

    # --- 1. supply -----------------------------------------------------------
    visited = set()
    for heater_id in graph.heater_ids:
        if heater_id in visited or heater_id not in graph.nodes:
            continue
        visited.add(heater_id)
        queue = [heater_id]
        while queue:
            current_id = queue.pop(0)
            for next_id in adjacency.get(current_id, []):
                if next_id in visited:
                    continue
                if is_return_pipe(graph, graph.nodes[next_id]):
                    continue
                visited.add(next_id)
                graph.nodes[next_id].parent_by_system[hot] = current_id
                graph.nodes[current_id].children_by_system.setdefault(
                    hot, []).append(next_id)
                queue.append(next_id)

    # Every heater, including one reached on the tree of another heater, so
    # none keeps a fixture list from the walk that went round the loop.
    for heater_id in graph.heater_ids:
        if heater_id not in graph.nodes:
            continue
        served = [nid for nid in graph.descendants(heater_id, hot)
                  if graph.nodes[nid].kind == KIND_FIXTURE]
        graph.nodes[heater_id].served_fixture_ids = served
        graph.log("HEATER {}: serves {} fixtures on its hot SUPPLY tree, "
                  "return piping excluded.".format(heater_id, len(served)))

    # --- 2. return -----------------------------------------------------------
    reached = set()
    for heater_id in graph.heater_ids:
        if heater_id in graph.nodes:
            _walk_return(graph, adjacency, heater_id, reached)

    # A return that ties into the cold water inlet, instead of a connector on
    # the heater, never touches the heater on the hot side. It ends at the
    # fitting it shares with the cold piping.
    cold = shared_params.SYSTEM_DOMESTIC_COLD_WATER
    for node_id in sorted(graph.nodes.keys()):
        node = graph.nodes[node_id]
        if node_id in reached or not is_return_pipe(graph, node):
            continue
        component = _return_component(graph, adjacency, node_id)
        tie_in = None
        for member_id in component:
            member = graph.nodes[member_id]
            if member.kind == KIND_PIPE:
                continue
            if (cold in member.parent_by_system or
                    member.children_by_system.get(cold)):
                tie_in = member_id
                break
        if tie_in is None:
            for member_id in component:
                reached.add(member_id)
                if is_return_pipe(graph, graph.nodes[member_id]):
                    graph.return_unrooted.append(member_id)
            graph.log("RETURN: could not find where the return line holding "
                      "pipe {} ends. It reaches neither a water heater nor "
                      "the cold water piping.".format(node_id))
            continue
        _walk_return(graph, adjacency, tie_in, reached)

    # --- 3. demand -----------------------------------------------------------
    _assign_demand(graph)
    graph.log("RETURN SPLIT: {} return line end(s), {} return pipe(s) with no "
              "end found.".format(len(graph.return_roots),
                                  len(graph.return_unrooted)))
    return adjacency


def _return_can_enter(graph, node):
    return is_return_pipe(graph, node) or node.kind in _RETURN_PASSABLE_KINDS


def _walk_return(graph, adjacency, root_id, reached):
    """Lay the return tree out from where it ends, back toward the fixtures."""
    reached.add(root_id)
    queue = [root_id]
    found_pipe = False
    while queue:
        current_id = queue.pop(0)
        for next_id in adjacency.get(current_id, []):
            if next_id in reached:
                continue
            node = graph.nodes[next_id]
            if not _return_can_enter(graph, node):
                continue
            reached.add(next_id)
            node.parent_by_system[TREE_RETURN] = current_id
            graph.nodes[current_id].children_by_system.setdefault(
                TREE_RETURN, []).append(next_id)
            if node.kind == KIND_PIPE:
                found_pipe = True
            queue.append(next_id)
    if found_pipe:
        graph.return_roots.append(root_id)


def _return_component(graph, adjacency, start_id):
    """The return pipes and in-line devices joined to start_id."""
    seen = set([start_id])
    queue = [start_id]
    while queue:
        current_id = queue.pop(0)
        for next_id in adjacency.get(current_id, []):
            if next_id in seen:
                continue
            if not _return_can_enter(graph, graph.nodes[next_id]):
                continue
            seen.add(next_id)
            queue.append(next_id)
    return sorted(seen)


def _existing_inside_new_run(graph, doc, element, came_from_id):
    """True for an existing phase in-line device with new work beyond it.

    Only a non-pipe element qualifies (a valve or fitting): an existing pipe
    is the existing main. Every connector other than the one the walk arrived
    through shall be connected, and to an element of the run's phase.
    """
    if _is_pipe(element):
        return False
    try:
        connectors = revit_helpers.get_connectors(element)
    except Exception:
        return False    # unreadable: stop at it, as for any existing element
    beyond = 0
    for c in connectors:
        neighbour_id = c.get("connected_element_id")
        if neighbour_id == came_from_id:
            continue
        if not c.get("is_connected") or neighbour_id is None:
            return False
        neighbour = _get_element(doc, neighbour_id)
        if neighbour is None:
            return False
        if revit_helpers.get_created_phase_id(neighbour) != graph.origin_phase_id:
            return False
        beyond += 1
    return beyond > 0


def _unused_return_outlet(node, connector, system):
    """True for the open HW Out of a fixture whose HWR_ACTIVE is No.

    The family keeps an HW Out on every Type, but only a fixture on
    recirculation pipes it, so an open one elsewhere is expected.
    """
    return (node.kind == KIND_FIXTURE and
            system == shared_params.SYSTEM_DOMESTIC_HOT_WATER and
            connector.get("direction") == "Out" and
            not node.hwr_active)


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

    Water heaters in parallel (Colin, 2026-10-21): heaters whose hot outlets
    join the same hot supply piping split that piping's hot load equally. A
    cold pipe carries the group's hot load times the share of the group's
    heaters downstream of it: the cold feed to one of two heaters carries
    half, the common cold header ahead of both carries all of it.
    """
    groups = heater_groups(graph)
    group_fixtures = {}
    for heater_id, group in groups.items():
        fixtures = group_fixtures.setdefault(group, set())
        fixtures.update(graph.nodes[heater_id].served_fixture_ids)

    for node in graph.nodes.values():
        if node.kind != KIND_PIPE:
            continue
        node.parallel_leg = False   # demand is redone after the return split

        downstream = graph.descendants(node.element_id, node.system)

        node.demand_fixed_gpm = 0.0
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
        heaters_by_group = {}
        for nid in downstream:
            child = graph.nodes[nid]
            if child.kind == KIND_FIXTURE:
                direct.add(nid)
            elif child.kind == KIND_HEATER:
                group = groups.get(nid, (nid,))
                heaters_by_group.setdefault(group, set()).add(nid)

        # Share of each fixture's hot load this pipe carries through heaters.
        via_heater = {}
        for group, heaters in heaters_by_group.items():
            share = float(len(heaters)) / float(len(group))
            fixtures = group_fixtures.get(
                group, graph.nodes[list(heaters)[0]].served_fixture_ids)
            for fid in fixtures:
                via_heater[fid] = min(1.0, via_heater.get(fid, 0.0) + share)

        total = 0.0
        for fid in direct:
            total += graph.nodes[fid].total_wsfu
            node.demand_fixed_gpm += float(graph.nodes[fid].fixed_gpm or 0.0)
        for fid, share in via_heater.items():
            if fid not in direct:
                total += graph.nodes[fid].hw_wsfu * share
        node.demand_wsfu = total

    _assign_parallel_hot_demand(graph, groups)
    _assign_parallel_backflow_demand(graph)


def _is_backflow_preventer(node):
    name = (node.family_name or "").lower()
    return node.kind in (KIND_ACCESSORY, KIND_EQUIPMENT) and (
        "rpz" in name or "backflow" in name)


def _assign_parallel_backflow_demand(graph):
    """Size the leg behind a second backflow preventer for the full demand.

    Two RPZs in parallel off one service each feed the same main, so either
    one can carry the whole load. From the picked RPZ the other one's outlet
    leg hangs off the main as a dead end with nothing downstream, which would
    leave it at zero demand and size a pipe that really carries flow at the
    minimum. That leg takes the demand of the picked RPZ's own outlet pipe, so
    the result is the same from either pick (E18, Colin 2026-10-26).
    """
    cold = shared_params.SYSTEM_DOMESTIC_COLD_WATER
    graph.parallel_backflow = []
    origin = graph.nodes.get(graph.origin_id)
    if origin is None or origin.kind != KIND_ORIGIN:
        return
    full = 0.0
    full_fixed = 0.0
    for child_id in origin.children_by_system.get(cold, []):
        child = graph.nodes.get(child_id)
        if child is not None and child.kind == KIND_PIPE:
            full = max(full, float(child.demand_wsfu or 0.0))
            full_fixed = max(full_fixed,
                             float(child.demand_fixed_gpm or 0.0))
    if full <= 0.0 and full_fixed <= 0.0:
        return
    # Where a leg starts: the other backflow preventer itself when the walk
    # came up through its outlet, or the dead end of the leg when the walk
    # reached the preventer from its inlet side and the outlet leg closed a
    # loop joint back to it.
    starts = []
    for nid, node in graph.nodes.items():
        if (nid != graph.origin_id and nid not in graph.upstream_ids and
                _is_backflow_preventer(node)):
            starts.append((nid, nid))
    for a, b, system in graph.loop_edges:
        if system != cold:
            continue
        for end, preventer in ((a, b), (b, a)):
            other = graph.nodes.get(preventer)
            if (other is not None and preventer != graph.origin_id and
                    end in graph.nodes and end not in graph.upstream_ids and
                    _is_backflow_preventer(other)):
                starts.append((end, preventer))

    done = set()
    for start_id, preventer_id in starts:
        leg = []
        seen = set()
        current = start_id
        while current is not None and current not in seen:
            seen.add(current)
            node = graph.nodes.get(current)
            if node is None or current in graph.upstream_ids:
                break
            if node.kind == KIND_PIPE:
                if (float(node.demand_wsfu or 0.0) > 0.0 or
                        float(node.demand_fixed_gpm or 0.0) > 0.0):
                    break
                leg.append(current)
            current = node.parent_by_system.get(cold)
        if not leg:
            continue
        for pid in leg:
            graph.nodes[pid].demand_wsfu = full
            graph.nodes[pid].demand_fixed_gpm = full_fixed
        if preventer_id not in done:
            done.add(preventer_id)
            graph.parallel_backflow.append(preventer_id)
        graph.log("PARALLEL BACKFLOW {}: {} pipe(s) behind it sized for the "
                  "full demand.".format(preventer_id, len(leg)))


def _assign_parallel_hot_demand(graph, groups):
    """Hot demand on the piping of water heaters in parallel.

    Colin, 2026-10-21: the hot water out of each of two heaters is sized on
    half the hot load, and the loads combine at the tee where the outlets
    join. The hot walk roots the tree at one heater, so without this that
    heater's outlet carries the whole load and the other's carries none.

    Each heater in a group puts in an equal share of the group's hot load.
    The hot supply piping is a tree, so cutting a pipe leaves two sides, and
    the pipe carries what the heaters on one side put in less what the
    fixtures on that side take out. A heater's own outlet carries 1/n, the
    piping past the tee where every outlet has joined carries all of it, and
    a pipe with only some of the heaters on one side gets parallel_leg set to
    the group.
    """
    hot = shared_params.SYSTEM_DOMESTIC_HOT_WATER
    adjacency = None
    done = set()
    for group in set(groups.values()):
        if len(group) < 2:
            continue
        if adjacency is None:
            adjacency = _hot_supply_adjacency(graph)
        component = _side(adjacency, group[0], None)
        members = set(group)
        load = sum(graph.nodes[nid].hw_wsfu for nid in component
                   if graph.nodes[nid].kind == KIND_FIXTURE)
        share = load / float(len(group))
        for nid in sorted(component):
            node = graph.nodes[nid]
            if (nid in done or node.kind != KIND_PIPE or node.system != hot or
                    is_return_pipe(graph, node)):
                continue
            done.add(nid)
            # A pipe has two ends, so two sides that agree. Every side is
            # checked and the largest kept, in case anything hangs mid pipe.
            demand = 0.0
            leg = False
            for neighbour in adjacency.get(nid, []):
                side = _side(adjacency, neighbour, nid)
                heaters = len(members & side)
                taken = sum(graph.nodes[s].hw_wsfu for s in side
                            if graph.nodes[s].kind == KIND_FIXTURE)
                demand = max(demand, abs(heaters * share - taken))
                leg = leg or 0 < heaters < len(group)
            node.demand_wsfu = demand
            node.parallel_leg = group if leg else False
        graph.log("HEATERS IN PARALLEL {}: {:.1f} hot WSFU, {:.2f} out of "
                  "each.".format(", ".join(str(h) for h in group), load, share))


def _hot_supply_adjacency(graph):
    """Hot water joints with the return piping left out."""
    adjacency = {}
    for nid, neighbours in (graph.hot_joints or hot_adjacency(graph)).items():
        if is_return_pipe(graph, graph.nodes.get(nid)):
            continue
        adjacency[nid] = [n for n in neighbours
                          if not is_return_pipe(graph, graph.nodes.get(n))]
    return adjacency


def _side(adjacency, start_id, cut_id):
    """Everything joined to start_id without passing through cut_id."""
    seen = set([start_id])
    stack = [start_id]
    while stack:
        current = stack.pop()
        for neighbour in adjacency.get(current, []):
            if neighbour != cut_id and neighbour not in seen:
                seen.add(neighbour)
                stack.append(neighbour)
    return seen


def heater_groups(graph):
    """Water heaters in parallel: heater id -> sorted tuple of its group.

    Two heaters are in parallel when both are fed from the cold tree and
    their hot outlets join the same hot supply piping (return piping is not
    followed, so heaters that only share a return stay separate). A heater on
    its own is a group of one. Stored on graph.heater_groups for the report.
    """
    cold = shared_params.SYSTEM_DOMESTIC_COLD_WATER
    adjacency = graph.hot_joints or hot_adjacency(graph)
    cold_fed = set(h for h in graph.heater_ids
                   if h in graph.nodes and
                   cold in graph.nodes[h].parent_by_system)

    groups = {}
    for start in graph.heater_ids:
        if start in groups or start not in graph.nodes:
            continue
        if start not in cold_fed:
            groups[start] = (start,)
            continue
        found = set([start])
        seen = set([start])
        queue = [start]
        while queue:
            current = queue.pop()
            for neighbour in adjacency.get(current, []):
                if neighbour in seen or neighbour not in graph.nodes:
                    continue
                seen.add(neighbour)
                other = graph.nodes[neighbour]
                if other.kind == KIND_HEATER:
                    if neighbour in cold_fed:
                        found.add(neighbour)
                    continue        # never walk through a heater
                if is_return_pipe(graph, other):
                    continue
                queue.append(neighbour)
        group = tuple(sorted(found))
        for heater_id in group:
            groups[heater_id] = group
    graph.heater_groups = groups
    return groups


# =============================================================================
# CLASSIFICATION HELPERS
# =============================================================================

def _is_flush_valve(element, node):
    """A flushometer VALVE fixture (water closet or urinal).

    Flush_Type = Valve when the family carries it; otherwise the Type or
    fixture name says "flushometer valve" or "flush valve". A flushometer
    TANK is a tank fixture, not a flush valve.
    """
    flush_type = revit_helpers.get_type_parameter_value(
        element, shared_params.PARAM_FLUSH_TYPE)
    if flush_type is None:
        flush_type = revit_helpers.get_parameter_value(
            element, shared_params.PARAM_FLUSH_TYPE)
    if isinstance(flush_type, str) and flush_type.strip():
        return flush_type.strip().lower() == "valve"
    return name_is_flush_valve(node.type_name, node.fixture_name)


def name_is_flush_valve(*names):
    text = " ".join(str(n or "") for n in names).lower()
    return "flushometer valve" in text or "flush valve" in text


def has_flush_valves(graph):
    """True when the system serves at least one flush valve fixture on cold
    water. Drives the Table E103.3(3) column for the whole cold system
    (Colin, 2026-10-21)."""
    for fid in graph.fixture_ids:
        node = graph.nodes.get(fid)
        if node is None:
            continue
        if getattr(node, "is_flush_valve", False) and \
                float(node.cw_wsfu or 0.0) > 0.0:
            return True
    return False


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


def _looks_like_pump(element, connectors, reached_on_system):
    """A recirculation pump: Mechanical Equipment sitting on the hot side.

    Identified by category, by the system the traversal REACHED it on, and by
    the absence of a cold water connector. Never by family or Type name.

    The key point, learned the hard way on (2024) Grantham 4 MP: a real
    circulation pump does NOT report its own connectors as DomesticHotWater.
    The HWCP ecocircXL there reports OtherPipe on two connectors and throws on
    the third, so an earlier rule that demanded two DomesticHotWater
    connectors never matched it and the tool wrongly reported "no pump found".

    What is trustworthy is the system the traversal arrived on: the pipe that
    connects to the pump IS classified, so if we walked to this element along
    the hot water network, it sits on the hot water network whatever its own
    connectors claim.

    Still a CANDIDATE, not a certainty: any other hot-side mechanical
    equipment with two or more connectors matches too. The checks report the
    family name so a human can confirm, rather than asserting what it is.
    """
    if reached_on_system != shared_params.SYSTEM_DOMESTIC_HOT_WATER:
        return False

    try:
        category = element.Category.Name
    except Exception:
        return False
    if "Equipment" not in category:
        return False

    # A cold water connector means a water heater, not a pump. The heater
    # check runs first anyway, but this keeps the rule true on its own.
    for c in connectors:
        if c.get("system_type") == shared_params.SYSTEM_DOMESTIC_COLD_WATER:
            return False

    # An inline pump needs an inlet and an outlet. Connectors whose system
    # could not be read still count, because unreadable is not the same as
    # absent, and on this family several of them are unreadable.
    return len(connectors) >= 2


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
