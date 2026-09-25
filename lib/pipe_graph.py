# pipe_graph.py
# Connector traversal engine. Walks the gas piping network from the meter
# through every pipe, fitting, and fixture to build a network graph.
#
# Implements:
#   - build_network(origin_element, doc) -> NetworkGraph
#
# IronPython 2.7  -  Revit API via Autodesk.Revit.DB

from Autodesk.Revit.DB import ElementId

import shared_params
import revit_helpers


# =============================================================================
# NETWORK GRAPH DATA STRUCTURES
# =============================================================================

class NetworkNode(object):
    """A node in the piping network  -  meter, fitting, or fixture."""

    def __init__(self, element_id, element, node_type):
        """
        Args:
            element_id: int  -  Revit element ID
            element: Revit Element object
            node_type: str  -  "meter", "tee", "fitting", "fixture", "unknown"
        """
        self.element_id         = element_id
        self.element            = element
        self.node_type          = node_type
        self.family_name        = _get_family_name(element)
        self.location_xyz       = revit_helpers.get_element_location(element)
        self.connector_count    = 0
        self.connectors         = []

        # Fixture-specific  -  populated during traversal if IS_GAS_FIXTURE = True
        self.gas_load_mbh       = 0.0
        self.fixture_name       = ""
        self.is_gas_fixture     = False

        # Elbow detection  -  True if family name contains "Elbow"
        # Used by _find_longest_run to add 5ft equivalent length per elbow
        self.is_elbow           = "Elbow" in self.family_name

        # Cumulative load at this node  -  sum of all downstream fixture loads
        self.cumulative_load_mbh = 0.0

        # Pressure regulating valve  -  auto-detected by family name
        # (shared_params.PRV_FAMILY_KEYWORDS), no custom parameters.
        # is_midstream_prv is True only when the pipe run downstream of it is
        # longer than shared_params.PRV_MIDSTREAM_MIN_DOWNSTREAM_FT. A PRV
        # closer to its equipment than that is the equipment's own regulator
        # and is ignored for sizing. downstream_pipe_ft is that run length.
        self.is_prv             = False
        self.is_midstream_prv   = False
        self.downstream_pipe_ft = None

        # Pressure zone  -  number of PRVs between the meter and this node
        # (a PRV itself is still in the zone it regulates FROM). run_key is
        # the element ID of the meter or PRV whose downstream system this
        # node belongs to. Set by _assign_pressure_zones().
        self.zone               = 0
        self.run_key            = None


class NetworkEdge(object):
    """An edge in the piping network  -  a single pipe segment."""

    def __init__(self, element_id, pipe, from_node_id, to_node_id):
        """
        Args:
            element_id: int  -  Revit element ID of the pipe
            pipe: Revit Pipe element
            from_node_id: int  -  upstream node element ID
            to_node_id: int  -  downstream node element ID
        """
        self.element_id         = element_id
        self.pipe               = pipe
        self.from_node_id       = from_node_id
        self.to_node_id         = to_node_id
        self.length_feet        = revit_helpers.get_pipe_length_feet(pipe) or 0.0
        self.diameter_inches    = revit_helpers.get_pipe_diameter_inches(pipe) or 0.0

        # Start and end XYZ captured for one-line diagram layout (Phase 2)
        connectors = revit_helpers.get_connectors(pipe)
        self.start_xyz = connectors[0]["origin_xyz"] if len(connectors) > 0 else None
        self.end_xyz   = connectors[1]["origin_xyz"] if len(connectors) > 1 else None

        # Cumulative load carried by this segment  -  set after traversal
        self.cumulative_load_mbh = 0.0

        # Pressure zone of this segment (0 = meter side of every PRV, 1 =
        # after one PRV) and the meter/PRV element ID whose longest run
        # governs its sizing. Set by _assign_pressure_zones().
        self.zone               = 0
        self.run_key            = None


class NetworkGraph(object):
    """The complete piping network graph."""

    def __init__(self):
        self.nodes              = {}    # element_id (int) -> NetworkNode
        self.edges              = {}    # element_id (int) -> NetworkEdge
        self.origin_id          = None  # meter element ID
        self.traversal_log      = []    # step-by-step traversal decisions
        self.disconnected       = []    # element IDs the traversal could not reach
        self.longest_run        = None  # populated by _find_longest_run()
        self.node_children      = {}    # node_id -> [child_node_id, ...]
                                        # Records direct fitting-to-fitting connections
                                        # where no pipe exists between them in the model
        self.prv_ids            = []    # element IDs of every PRV reached
        self.midstream_prv_ids  = []    # the PRVs that count as a step down
                                        # (downstream run over the threshold);
                                        # the rest are equipment PRVs, ignored
        self.nested_prv_ids     = []    # mid-stream PRVs downstream of another
                                        # (more than one step down - unsupported)
        self.zone_runs          = {}    # meter/PRV element ID -> longest-run dict
                                        # for the system that node feeds. Only
                                        # filled when the graph has a mid-stream
                                        # PRV.

    def add_node(self, node):
        self.nodes[node.element_id] = node

    def add_edge(self, edge):
        self.edges[edge.element_id] = edge

    def log(self, message):
        self.traversal_log.append(message)


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def build_network(origin_element, doc):
    """Walk the connector graph from the meter and build the network.

    Starting from origin_element (the gas meter), follows every connected
    pipe and fitting until all reachable elements are visited. Identifies
    fixtures by IS_GAS_FIXTURE parameter. Calculates cumulative loads and
    longest run.

    Args:
        origin_element: The user-selected gas meter Revit element.
        doc: The active Revit Document (needed for GetElement calls).

    Returns:
        NetworkGraph  -  fully populated graph ready for report generation.
    """
    graph = NetworkGraph()

    # --- Build origin node (meter) ---
    origin_id = revit_helpers.eid_int(origin_element.Id)
    graph.origin_id = origin_id

    origin_node = NetworkNode(origin_id, origin_element, "meter")
    origin_node.connectors = revit_helpers.get_connectors(origin_element)
    origin_node.connector_count = len(origin_node.connectors)
    graph.add_node(origin_node)
    graph.log("START: Meter element {} ({})".format(origin_id, origin_node.family_name))

    # --- Traversal ---
    # visited tracks element IDs we have already processed to prevent loops
    visited = set()
    visited.add(origin_id)

    # Queue entries are tuples: (current_element, parent_node_id)
    # Start by walking only the Out connector from the meter
    queue = []

    for connector in origin_node.connectors:
        if connector["direction"] == "Out" and connector["is_connected"]:
            connected_id = connector["connected_element_id"]
            if connected_id is not None:
                connected_element = _get_element(doc, connected_id)
                if connected_element is not None:
                    queue.append((connected_element, origin_id))
                    graph.log(
                        "METER Out connector -> queuing element {} to traverse".format(
                            connected_id))
        elif connector["direction"] == "In":
            graph.log(
                "METER In connector (street side)  -  skipping direction.")

    # Breadth-first traversal
    while queue:
        current_element, parent_node_id = queue.pop(0)
        current_id = revit_helpers.eid_int(current_element.Id)

        if current_id in visited:
            graph.log("SKIP: Element {} already visited.".format(current_id))
            continue

        visited.add(current_id)

        element_type = current_element.GetType().Name
        graph.log("VISIT: Element {} type={}".format(current_id, element_type))

        # --- Pipe ---
        if _is_pipe(current_element):
            _process_pipe(graph, doc, current_element, parent_node_id,
                          visited, queue)

        # --- Family instance (fitting, fixture, equipment) ---
        elif _is_family_instance(current_element):
            _process_family_instance(graph, doc, current_element, parent_node_id,
                                     visited, queue)

        else:
            graph.log(
                "UNKNOWN element type {} on element {}  -  skipping.".format(
                    element_type, current_id))

    # --- Post-traversal calculations ---
    _calculate_cumulative_loads(graph)
    _find_longest_run(graph)
    _classify_midstream_prvs(graph)
    _assign_pressure_zones(graph)
    _find_zone_runs(graph)

    graph.log("TRAVERSAL COMPLETE: {} nodes, {} edges.".format(
        len(graph.nodes), len(graph.edges)))

    return graph


# =============================================================================
# ELEMENT PROCESSORS
# =============================================================================

def _process_pipe(graph, doc, pipe, parent_node_id, visited, queue):
    """Process a pipe element  -  create an edge and queue the far end."""
    pipe_id = revit_helpers.eid_int(pipe.Id)

    # Find the two connectors on the pipe
    connectors = revit_helpers.get_connectors(pipe)

    # Identify which connector leads back to parent and which goes forward
    far_connected_id = None
    for c in connectors:
        if c["is_connected"] and c["connected_element_id"] != parent_node_id:
            far_connected_id = c["connected_element_id"]
            break

    graph.log(
        "PIPE {}: length={:.2f}ft, diameter={:.3f}in, far_end={}".format(
            pipe_id,
            revit_helpers.get_pipe_length_feet(pipe) or 0.0,
            revit_helpers.get_pipe_diameter_inches(pipe) or 0.0,
            far_connected_id
        ))

    # Open-ended pipe  -  far end has no connected element
    if far_connected_id is None:
        graph.log(
            "WARNING: Pipe {} has an unconnected (open) far end.".format(pipe_id))
        graph.disconnected.append(pipe_id)
        edge = NetworkEdge(pipe_id, pipe, parent_node_id, None)
        graph.add_edge(edge)
        return

    edge = NetworkEdge(pipe_id, pipe, parent_node_id, far_connected_id)
    graph.add_edge(edge)

    # Queue the far end element
    if far_connected_id not in visited:
        far_element = _get_element(doc, far_connected_id)
        if far_element is not None:
            queue.append((far_element, pipe_id))
        else:
            graph.log(
                "WARNING: Could not retrieve far end element {} from pipe {}.".format(
                    far_connected_id, pipe_id))
            graph.disconnected.append(far_connected_id)


def _process_family_instance(graph, doc, element, parent_node_id, visited, queue):
    """Process a family instance  -  fitting, fixture, or equipment."""
    eid = revit_helpers.eid_int(element.Id)
    family_name = _get_family_name(element)
    connectors = revit_helpers.get_connectors(element)
    connector_count = len(connectors)

    # Determine node type
    if connector_count >= 3:
        node_type = "tee"
    elif connector_count == 1:
        node_type = "fixture"  # will be confirmed by IS_GAS_FIXTURE below
    else:
        node_type = "fitting"

    node = NetworkNode(eid, element, node_type)
    node.connectors = connectors
    node.connector_count = connector_count

    # --- Check if this is a gas fixture ---
    is_gas_fixture = revit_helpers.get_parameter_value(
        element, shared_params.PARAM_IS_GAS_FIXTURE)
    gas_load_mbh = revit_helpers.get_parameter_value(
        element, shared_params.PARAM_GAS_LOAD_MBH)
    fixture_name = revit_helpers.get_parameter_value(
        element, shared_params.PARAM_FIXTURE_NAME)

    if is_gas_fixture:
        node.node_type          = "fixture"
        node.is_gas_fixture     = True
        node.gas_load_mbh       = float(gas_load_mbh) if gas_load_mbh is not None else 0.0
        node.fixture_name       = fixture_name if fixture_name else "UNNAMED"
        node.cumulative_load_mbh = node.gas_load_mbh
        graph.log(
            "FIXTURE {}: name='{}', load={} MBH".format(
                eid, node.fixture_name, node.gas_load_mbh))
    else:
        graph.log(
            "FITTING/TEE {}: family='{}', connectors={}, type={}".format(
                eid, family_name, connector_count, node_type))

    # --- Check if this is a pressure regulating valve ---
    # Auto-detected by family name; whether it is a mid-stream step down is
    # decided after traversal by _classify_midstream_prvs().
    lowered_name = (family_name or "").lower()
    if (not is_gas_fixture and
            any(kw in lowered_name for kw in shared_params.PRV_FAMILY_KEYWORDS)):
        node.is_prv = True
        graph.prv_ids.append(eid)
        graph.log("PRV {}: family='{}'".format(eid, family_name))

    graph.add_node(node)

    # --- Record direct node-to-node connection if parent is also a node ---
    # This happens when a tee connects straight to a fitting with no pipe
    # between them. Without recording this, the load algorithm can't cross
    # the gap because it only walks edges (pipes).
    if parent_node_id in graph.nodes:
        if parent_node_id not in graph.node_children:
            graph.node_children[parent_node_id] = []
        graph.node_children[parent_node_id].append(eid)
        graph.log(
            "DIRECT CONNECTION: node {} -> node {} (no pipe between them)".format(
                parent_node_id, eid))

    # --- Queue all connected elements we haven't visited ---
    for c in connectors:
        if c["is_connected"] and c["connected_element_id"] is not None:
            next_id = c["connected_element_id"]
            if next_id not in visited:
                next_element = _get_element(doc, next_id)
                if next_element is not None:
                    queue.append((next_element, eid))
                else:
                    graph.log(
                        "WARNING: Could not retrieve element {} connected to {}.".format(
                            next_id, eid))
                    graph.disconnected.append(next_id)


# =============================================================================
# CUMULATIVE LOAD CALCULATION
# =============================================================================

def _calculate_cumulative_loads(graph):
    """Calculate cumulative MBH loads at every node and edge.

    Works by summing loads from all fixture descendants at each node.
    Uses a bottom-up traversal from fixtures back toward the meter.
    """
    if graph.origin_id is None:
        return

    # Build adjacency: for each node, which edges connect downstream
    # We use a recursive DFS from the origin
    visited = set()

    def _sum_load(node_id):
        if node_id in visited:
            return 0.0
        visited.add(node_id)

        node = graph.nodes.get(node_id)
        if node is None:
            return 0.0

        # If this is a fixture, its load is its own gas_load_mbh
        if node.is_gas_fixture:
            return node.gas_load_mbh

        # Sum loads from downstream pipe edges
        total = 0.0
        for edge in graph.edges.values():
            if edge.from_node_id == node_id and edge.to_node_id is not None:
                downstream_load = _sum_load(edge.to_node_id)
                edge.cumulative_load_mbh = downstream_load
                total += downstream_load

        # Sum loads from direct node-to-node connections (no pipe between them)
        for child_id in graph.node_children.get(node_id, []):
            total += _sum_load(child_id)

        node.cumulative_load_mbh = total
        return total

    origin_load = _sum_load(graph.origin_id)
    origin_node = graph.nodes.get(graph.origin_id)
    if origin_node:
        origin_node.cumulative_load_mbh = origin_load


# =============================================================================
# LONGEST RUN CALCULATION
# =============================================================================

def _find_longest_run(graph):
    """Find the longest developed length from meter to any fixture.

    Per IFGC A103.1 - this single length is used to size ALL segments
    in Phase 2.

    Developed length for each path =
        sum of pipe lengths along that path
        + (number of elbows along that path x 5ft equivalent length)

    The path with the greatest developed length is the longest run.
    Elbows are counted per path - an elbow only contributes to paths
    that physically pass through it.
    """
    if graph.origin_id is None:
        return

    graph.longest_run = _longest_run_from(graph, graph.origin_id, False)


def _longest_run_from(graph, start_id, stop_at_prv):
    """Longest developed length from start_id to any terminal, as the dict
    _find_longest_run() has always produced.

    Args:
        start_id:     element ID of the node to walk out from (the meter, or
                      a PRV when sizing the system downstream of it).
        stop_at_prv:  True to treat any MID-STREAM PRV other than start_id as
                      a terminal (the walk ends there, it does not continue
                      through the regulator). False walks straight through
                      every PRV, which is the whole-system length
                      _find_longest_run() reports.
    """
    ELBOW_EQUIV_FT = 5.0

    longest = {
        "total_length_feet":        0.0,
        "pipe_length_feet":         0.0,
        "elbow_count":              0,
        "elbow_equiv_length_feet":  0.0,
        "path_element_ids":         [],
        "farthest_fixture_id":      None,
        "farthest_fixture_name":    ""
    }

    def _dfs(node_id, pipe_length, elbow_count, current_path):
        """Walk one path. pipe_length and elbow_count are local to this path."""
        node = graph.nodes.get(node_id)
        if node is None:
            return

        # If this node is an elbow, add its equivalent length to THIS path only
        local_elbow_count = elbow_count
        if node.is_elbow:
            local_elbow_count += 1

        total_length = pipe_length + (local_elbow_count * ELBOW_EQUIV_FT)

        ends_at_prv = stop_at_prv and node.is_midstream_prv and node_id != start_id
        if node.is_gas_fixture or ends_at_prv:
            end_name = node.fixture_name if node.is_gas_fixture else "PRV {}".format(node_id)
            graph.log(
                "PATH to {} '{}': pipe={:.2f}ft, elbows={}, "
                "equiv={:.0f}ft, total={:.2f}ft".format(
                    "fixture" if node.is_gas_fixture else "regulator",
                    end_name,
                    pipe_length,
                    local_elbow_count,
                    local_elbow_count * ELBOW_EQUIV_FT,
                    total_length
                ))
            if total_length > longest["total_length_feet"]:
                longest["total_length_feet"]        = total_length
                longest["pipe_length_feet"]         = pipe_length
                longest["elbow_count"]              = local_elbow_count
                longest["elbow_equiv_length_feet"]  = local_elbow_count * ELBOW_EQUIV_FT
                longest["path_element_ids"]         = list(current_path)
                longest["farthest_fixture_id"]      = node_id
                longest["farthest_fixture_name"]    = end_name
            return

        for edge in graph.edges.values():
            if edge.from_node_id == node_id and edge.to_node_id is not None:
                _dfs(
                    edge.to_node_id,
                    pipe_length + edge.length_feet,
                    local_elbow_count,
                    current_path + [edge.element_id, edge.to_node_id]
                )

        # Traverse direct node-to-node connections (zero pipe length added)
        for child_id in graph.node_children.get(node_id, []):
            _dfs(child_id, pipe_length, local_elbow_count,
                 current_path + [child_id])

    _dfs(start_id, 0.0, 0, [start_id])
    return longest


# =============================================================================
# PRESSURE ZONES (gas regulators)
# =============================================================================

def _classify_midstream_prvs(graph):
    """Decide which PRVs are a mid-stream step down.

    A PRV is mid-stream when the actual pipe length from it to its farthest
    fixture is longer than shared_params.PRV_MIDSTREAM_MIN_DOWNSTREAM_FT.
    A regulator right at its equipment is that equipment's own and is ignored
    for sizing (it stays in graph.prv_ids for reporting). Pipe length only,
    no elbow equivalents: a regulator a few feet from the equipment with a
    couple of elbows is still "at the equipment".
    """
    graph.midstream_prv_ids = []
    for pid in graph.prv_ids:
        node = graph.nodes.get(pid)
        if node is None:
            continue
        run = _longest_run_from(graph, pid, False)
        node.downstream_pipe_ft = run["pipe_length_feet"]
        node.is_midstream_prv = (
            node.downstream_pipe_ft > shared_params.PRV_MIDSTREAM_MIN_DOWNSTREAM_FT)
        if node.is_midstream_prv:
            graph.midstream_prv_ids.append(pid)
        graph.log(
            "PRV {}: {:.1f} ft of pipe downstream -> {}".format(
                pid, node.downstream_pipe_ft,
                "MID-STREAM step down" if node.is_midstream_prv
                else "equipment PRV, ignored for sizing"))


def _assign_pressure_zones(graph):
    """Tag every node and edge with its pressure zone and governing run.

    The meter feeds zone 0. Each MID-STREAM PRV starts the next zone for
    everything downstream of it. Every node/edge also gets a run_key: the
    element ID of the meter or PRV whose downstream system it belongs to, so
    a segment can be sized on its own system's longest run. Equipment PRVs
    (see _classify_midstream_prvs) do not start a zone.

    A PRV is itself in the zone it regulates FROM (it is the far end of the
    upstream system), and its outgoing pipes are in the next zone. A
    mid-stream PRV reached while already downstream of another one is
    recorded in graph.nested_prv_ids - only a single step down is supported.

    With no mid-stream PRV in the graph every node and edge stays in zone 0
    with the meter as run_key, so behaviour is unchanged.
    """
    if graph.origin_id is None:
        return

    origin = graph.nodes.get(graph.origin_id)
    if origin is None:
        return
    origin.zone = 0
    origin.run_key = graph.origin_id

    out_edges = {}
    for edge in graph.edges.values():
        out_edges.setdefault(edge.from_node_id, []).append(edge)

    seen = set()
    stack = [graph.origin_id]
    while stack:
        nid = stack.pop()
        if nid in seen:
            continue
        seen.add(nid)
        node = graph.nodes.get(nid)
        if node is None:
            continue

        if node.is_midstream_prv:
            if node.zone >= 1 and nid not in graph.nested_prv_ids:
                graph.nested_prv_ids.append(nid)
            out_zone, out_key = node.zone + 1, nid
        else:
            out_zone, out_key = node.zone, node.run_key

        for edge in out_edges.get(nid, []):
            edge.zone = out_zone
            edge.run_key = out_key
            child = graph.nodes.get(edge.to_node_id)
            if child is not None and child.element_id not in seen:
                child.zone = out_zone
                child.run_key = out_key
                stack.append(child.element_id)

        for child_id in graph.node_children.get(nid, []):
            child = graph.nodes.get(child_id)
            if child is not None and child_id not in seen:
                child.zone = out_zone
                child.run_key = out_key
                stack.append(child_id)


def _find_zone_runs(graph):
    """Longest developed length for the system fed by the meter and by each
    mid-stream PRV, stored in graph.zone_runs keyed by that element's ID.

    The meter's system ends at the farthest fixture OR mid-stream PRV on its
    side of the regulator; each PRV's system runs from the PRV to its
    farthest fixture. Same developed-length rule as _find_longest_run() (pipe
    length plus 5 ft per elbow along the path). Skipped when the graph has no
    mid-stream PRV.
    """
    graph.zone_runs = {}
    if graph.origin_id is None or not graph.midstream_prv_ids:
        return

    for start_id in [graph.origin_id] + list(graph.midstream_prv_ids):
        graph.zone_runs[start_id] = _longest_run_from(graph, start_id, True)


# =============================================================================
# HELPERS
# =============================================================================

def _get_element(doc, element_id_int):
    """Retrieve a Revit element by integer ID. Returns None on failure."""
    try:
        eid = ElementId(element_id_int)
        element = doc.GetElement(eid)
        return element
    except Exception as e:
        revit_helpers._log_entry(
            "ERROR", "_get_element", element_id_int,
            "doc.GetElement({}) failed: {}".format(element_id_int, str(e)))
        return None


def _is_pipe(element):
    """Return True if the element is a Revit pipe."""
    try:
        return element.GetType().Name == "Pipe"
    except Exception:
        return False


def _is_family_instance(element):
    """Return True if the element is a FamilyInstance."""
    try:
        return element.GetType().Name == "FamilyInstance"
    except Exception:
        return False


def _get_family_name(element):
    """Return the family name string or 'Unknown'."""
    try:
        return element.Symbol.Family.Name
    except Exception:
        try:
            return element.GetType().Name
        except Exception:
            return "Unknown"
