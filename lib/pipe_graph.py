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


class NetworkGraph(object):
    """The complete piping network graph."""

    def __init__(self):
        self.nodes              = {}    # element_id (int) -> NetworkNode
        self.edges              = {}    # element_id (int) -> NetworkEdge
        self.origin_id          = None  # meter element ID
        self.traversal_log      = []    # step-by-step traversal decisions
        self.disconnected       = []    # element IDs the traversal could not reach
        self.longest_run        = None  # populated by _find_longest_run()

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
    origin_id = origin_element.Id.IntegerValue
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
        current_id = current_element.Id.IntegerValue

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

    graph.log("TRAVERSAL COMPLETE: {} nodes, {} edges.".format(
        len(graph.nodes), len(graph.edges)))

    return graph


# =============================================================================
# ELEMENT PROCESSORS
# =============================================================================

def _process_pipe(graph, doc, pipe, parent_node_id, visited, queue):
    """Process a pipe element  -  create an edge and queue the far end."""
    pipe_id = pipe.Id.IntegerValue

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

    # Determine the to_node_id  -  will be set properly when far end is processed
    # For now use far_connected_id as placeholder
    to_node_id = far_connected_id if far_connected_id is not None else pipe_id

    edge = NetworkEdge(pipe_id, pipe, parent_node_id, to_node_id)
    graph.add_edge(edge)

    # Queue the far end element
    if far_connected_id is not None and far_connected_id not in visited:
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
    eid = element.Id.IntegerValue
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

    graph.add_node(node)

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

        # Otherwise sum loads from all downstream edges
        total = 0.0
        for edge in graph.edges.values():
            if edge.from_node_id == node_id:
                downstream_load = _sum_load(edge.to_node_id)
                edge.cumulative_load_mbh = downstream_load
                total += downstream_load

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

        if node.is_gas_fixture:
            graph.log(
                "PATH to fixture '{}': pipe={:.2f}ft, elbows={}, "
                "equiv={:.0f}ft, total={:.2f}ft".format(
                    node.fixture_name,
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
                longest["farthest_fixture_name"]    = node.fixture_name
            return

        for edge in graph.edges.values():
            if edge.from_node_id == node_id:
                _dfs(
                    edge.to_node_id,
                    pipe_length + edge.length_feet,
                    local_elbow_count,
                    current_path + [edge.element_id, edge.to_node_id]
                )

    _dfs(graph.origin_id, 0.0, 0, [graph.origin_id])
    graph.longest_run = longest


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
