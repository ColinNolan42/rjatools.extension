# -*- coding: ascii -*-
# tie_in.py
# Starting a sizing run at the first NEW pipe on an existing building, instead
# of at a gas meter or an RPZ. Shared by the gas and domestic water tools.
#
# Colin, 2026-10-13: on an existing building the meter or the RPZ is often
# existing and not modeled, so there is nothing to pick. The designer picks
# the new construction pipe that connects to existing. Check both sides of
# that pipe: the side that goes to existing is upstream, the side that goes
# to new construction is downstream. The existing piping is usually not
# modeled at all, so an open or capped end counts as the existing side too.
#
# pick_upstream_end() is pure and tested in CPython. The model read around it
# (phases, caps) needs a live Revit run to prove.
#
# IronPython 2.7

import shared_params
import revit_helpers


# What the far side of one connector on the picked pipe turned out to be.
END_OPEN     = "open"        # nothing connected: existing piping not modeled
END_EXISTING = "existing"    # connected to an element from another phase
END_CAP      = "cap"         # connected to a one-connector fitting
END_NEW      = "new"         # connected to new construction
END_TEE_ON_EXISTING = "tee_on_existing"   # a new phase tee Revit placed on a
                                          # modeled existing main (n=27)


def pick_upstream_end(ends):
    """Decide which end of a picked new pipe faces the existing piping.

    Args:
        ends: list of dicts with "connector_index" and "kind" (one of the
            END_ constants), one per connector considered on the pipe.

    Returns:
        dict with "status" ("ok", "no_upstream" or "ambiguous"),
        "connector_index" (when ok) and "reason", a clause for the user.
    """
    upstream = [e for e in ends if e["kind"] != END_NEW]
    if len(upstream) == 1:
        kind = upstream[0]["kind"]
        if kind == END_EXISTING:
            reason = "it connects to existing phase piping"
        elif kind == END_CAP:
            reason = "it is capped, the existing piping is not modeled"
        elif kind == END_TEE_ON_EXISTING:
            reason = ("it joins a new tee on the existing main (every other "
                      "connection of that tee is existing phase or open)")
        else:
            reason = "it is open, the existing piping is not modeled"
        return {"status": "ok",
                "connector_index": upstream[0]["connector_index"],
                "connector_id": upstream[0].get("connector_id"),
                "reason": reason}
    if not upstream:
        return {"status": "no_upstream",
                "reason": "every end of this pipe connects to new "
                          "construction, so neither is the existing side"}
    return {"status": "ambiguous",
            "reason": "more than one end of this pipe is open, capped or "
                      "connected to existing piping, so the upstream side "
                      "cannot be told from the downstream side"}


def resolve_new_pipe_upstream(pipe, doc, system=None):
    """Read a picked pipe's ends from the model and pick the upstream one.

    Args:
        pipe: the picked Revit Pipe.
        doc: the active Revit Document.
        system: a Connector.PipeSystemType name to restrict the connectors
            to (water passes Domestic Cold Water). None considers every
            connector, which is what gas needs.

    Returns:
        what pick_upstream_end() returns, plus:
          "ends"                     the classified connectors
          "upstream_xyz"             [x, y, z] of the upstream connector
          "upstream_neighbour_id"    element on the existing side, or None
          "downstream_neighbour_ids" elements on the new construction side
    """
    pipe_phase = revit_helpers.get_created_phase_id(pipe)

    ends = []
    for c in revit_helpers.get_connectors(pipe):
        if system is not None and c.get("system_type") != system:
            continue
        kind = END_NEW
        neighbour_id = c.get("connected_element_id")
        if not c.get("is_connected") or neighbour_id is None:
            kind = END_OPEN
            neighbour_id = None
        else:
            neighbour = revit_helpers.get_element_by_id(doc, neighbour_id)
            if neighbour is None:
                kind = END_OPEN
                neighbour_id = None
            else:
                phase = revit_helpers.get_created_phase_id(neighbour)
                if (pipe_phase is not None and phase is not None and
                        phase != pipe_phase):
                    kind = END_EXISTING
                elif _is_cap(neighbour):
                    kind = END_CAP
                elif _is_tee_on_existing(neighbour, pipe, pipe_phase, doc):
                    kind = END_TEE_ON_EXISTING
        ends.append({"connector_index": c.get("connector_index"),
                     "connector_id": c.get("connector_id"),
                     "kind": kind,
                     "neighbour_id": neighbour_id,
                     "origin_xyz": c.get("origin_xyz")})

    result = pick_upstream_end(ends)
    result["ends"] = ends
    result["upstream_xyz"] = None
    result["upstream_neighbour_id"] = None
    result["downstream_neighbour_ids"] = []
    if result["status"] == "ok":
        for end in ends:
            if end["connector_index"] == result["connector_index"]:
                result["upstream_xyz"] = end["origin_xyz"]
                result["upstream_neighbour_id"] = end["neighbour_id"]
            elif end["neighbour_id"] is not None:
                result["downstream_neighbour_ids"].append(end["neighbour_id"])
    return result


GAS_START_HINT = ("Pick the gas meter, or the first new gas pipe at the "
                  "connection to existing.")


def check_gas_start(element, doc, pipe_only=False):
    """The one check of a gas start pick, for Size Gas, One-Line and
    Diagnose (Colin, 2026-11-20). A pipe is checked as the first new pipe at
    an existing building (pipe_graph.check_gas_pipe, then
    resolve_new_pipe_upstream); anything else as the gas meter
    (revit_helpers.validate_selected_element).

    Args:
        element: the picked element.
        doc: the active Document.
        pipe_only: True with Connect into Existing: only a pipe is taken.

    Returns:
        dict with "ok" (bool), "kind" ("meter" or "tie_in", None when the
        pick is refused before it is read), "upstream" (for a pipe, what
        resolve_new_pipe_upstream returned; pass it to build_network),
        "gas_check" (for a pipe, what check_gas_pipe returned) and
        "message": why not, or what passed. The caller adds its own
        "Nothing was changed" line.
    """
    import pipe_graph     # pipe_graph imports tie_in: imported here
    is_pipe = revit_helpers.is_pipe(element)
    eid = revit_helpers._safe_eid(element)
    if pipe_only and not is_pipe:
        message = ("Connect into Existing was chosen, so the pick shall be "
                   "the first new gas pipe at the connection to existing, "
                   "not the gas meter or other equipment. Pick that pipe, or "
                   "choose New Construction to start at the meter.")
        revit_helpers._log_entry("WARNING", "check_gas_start", eid, message)
        return {"ok": False, "kind": None, "upstream": None,
                "gas_check": None, "message": message}
    if is_pipe:
        gas_check = pipe_graph.check_gas_pipe(element)
        if not gas_check.get("ok"):
            message = "This is not a gas pipe: {}.\n\n{}".format(
                gas_check.get("reason", "unknown"), GAS_START_HINT)
            revit_helpers._log_entry("WARNING", "check_gas_start", eid,
                                     message)
            return {"ok": False, "kind": "tie_in", "upstream": None,
                    "gas_check": gas_check, "message": message}
        upstream = resolve_new_pipe_upstream(element, doc)
        if upstream.get("status") != "ok":
            message = ("Could not tell which end of the picked pipe connects "
                       "to the existing gas piping: {}.\n\nPick the first new "
                       "gas pipe, the one with one end at the existing "
                       "connection (open, capped, or joined to existing "
                       "phase piping) and the other end on new "
                       "construction.".format(
                           upstream.get("reason", "unknown")))
            revit_helpers._log_entry("WARNING", "check_gas_start", eid,
                                     message)
            return {"ok": False, "kind": "tie_in", "upstream": upstream,
                    "gas_check": gas_check, "message": message}
        message = ("New gas piping connection. Upstream end found because "
                   "{}.".format(upstream.get("reason", "")))
        revit_helpers._log_entry("INFO", "check_gas_start", eid, message)
        return {"ok": True, "kind": "tie_in", "upstream": upstream,
                "gas_check": gas_check, "message": message}
    validation = revit_helpers.validate_selected_element(element)
    if not validation.get("is_valid"):
        message = "{}\n\n{}".format(GAS_START_HINT,
                                    validation.get("reason", ""))
        revit_helpers._log_entry("WARNING", "check_gas_start", eid, message)
        return {"ok": False, "kind": "meter", "upstream": None,
                "gas_check": None, "message": message}
    revit_helpers._log_entry("INFO", "check_gas_start", eid,
                             "Meter validation passed.")
    return {"ok": True, "kind": "meter", "upstream": None,
            "gas_check": None, "message": "Meter validation passed."}


def pick_tee_branch(ends):
    """Decide which connector of a picked tee leads to new construction.

    Colin, 2026-10-13: on an existing building the designer picks the TEE on
    the existing main. It shows both sides: the existing main runs through it
    and the new branch leaves it. Exactly one connector shall lead to new
    construction; every other one faces the existing piping.

    Args:
        ends: list of dicts with "connector_index" and "kind" (END_ constants).

    Returns:
        dict with "status" ("ok", "no_new" or "ambiguous"), "reason", and
        when ok "branch_connector" and "upstream_connectors" (a list).
    """
    new = [e for e in ends if e["kind"] == END_NEW]
    if len(new) == 1:
        return {"status": "ok",
                "branch_connector": new[0]["connector_index"],
                "upstream_connectors": [e["connector_index"] for e in ends
                                        if e is not new[0]],
                "upstream_connector_ids": [e.get("connector_id") for e in ends
                                           if e is not new[0]],
                "reason": "the tee's branch is the only connection to new "
                          "construction, the rest of the tee is on the "
                          "existing main"}
    if not new:
        return {"status": "no_new",
                "reason": "no connector of this tee leads to new "
                          "construction piping, so there is nothing new to "
                          "size from it"}
    return {"status": "ambiguous",
            "reason": "{} connectors of this tee lead to new construction, "
                      "so the existing main cannot be told from the new "
                      "branch".format(len(new))}


def resolve_tee_branch(tee, doc, system=None):
    """Read a picked tee's connectors from the model and find its new branch.

    A connector leads to new construction when the element on it was created
    in the project's LAST phase (New Construction). Anything earlier, open or
    capped is the existing side.

    Returns:
        what pick_tee_branch() returns, plus "ends" and "new_phase_id".
    """
    new_phase_id = revit_helpers.get_last_phase_id(doc)
    ends = []
    for c in revit_helpers.get_connectors(tee):
        if system is not None and c.get("system_type") != system:
            continue
        kind = END_OPEN
        neighbour_id = c.get("connected_element_id")
        if c.get("is_connected") and neighbour_id is not None:
            neighbour = revit_helpers.get_element_by_id(doc, neighbour_id)
            if neighbour is not None:
                phase = revit_helpers.get_created_phase_id(neighbour)
                if _is_cap(neighbour):
                    kind = END_CAP
                elif new_phase_id is not None and phase == new_phase_id:
                    kind = END_NEW
                else:
                    kind = END_EXISTING
        ends.append({"connector_index": c.get("connector_index"),
                     "connector_id": c.get("connector_id"),
                     "kind": kind, "neighbour_id": neighbour_id})
    result = pick_tee_branch(ends)
    result["ends"] = ends
    result["new_phase_id"] = new_phase_id
    return result


def _is_tee_on_existing(neighbour, pipe, pipe_phase, doc):
    """True when neighbour is a NEW phase fitting with 3 or more connectors
    whose every other connection (all but the one to the picked pipe) is
    existing phase or open, and at least one is existing phase.

    Review n=27: when the existing main is modeled and the designer draws
    the new branch into it, Revit splits the existing pipe and places a tee
    in the NEW phase. The first new pipe then had both ends on new
    construction and the start was refused as no_upstream. That tee is on
    the existing main, so it is the existing side.
    """
    if pipe_phase is None or revit_helpers.is_pipe(neighbour):
        return False
    pipe_id = revit_helpers.eid_int(pipe.Id)
    nid = revit_helpers.eid_int(neighbour.Id)
    connectors = revit_helpers.piping_connectors(
        revit_helpers.get_connectors(neighbour), nid)
    if len(connectors) < 3:
        return False
    existing = 0
    for c in connectors:
        other_id = c.get("connected_element_id")
        if other_id == pipe_id:
            continue
        if not c.get("is_connected") or other_id is None:
            continue                      # open: the main not modeled on
                                          # that side
        other = revit_helpers.get_element_by_id(doc, other_id)
        phase = (revit_helpers.get_created_phase_id(other)
                 if other is not None else None)
        if phase is None or phase == pipe_phase:
            revit_helpers._log_entry(
                "INFO", "_is_tee_on_existing", nid,
                "Connection to {} is new construction or unreadable: not a "
                "tee on the existing main.".format(other_id))
            return False
        existing += 1
    revit_helpers._log_entry(
        "INFO", "_is_tee_on_existing", nid,
        "{} existing phase connection(s): {}.".format(
            existing, "tee on the existing main" if existing else
            "NOT a tee on the existing main"))
    return existing > 0


def _is_cap(element):
    """A pipe cap: a one-connector pipe fitting that is not a load.

    Gas appliance and water fixture families can be one-connector families
    too, so anything flagged as a fixture is never a cap.
    """
    if revit_helpers.is_pipe(element):
        return False
    try:
        category = element.Category.Name
    except Exception as e:
        revit_helpers._log_entry(
            "INFO", "_is_cap", None,
            "Category of the connected element could not be read ({}), so "
            "it is not treated as a cap.".format(str(e)))
        return False
    if not category or "Fitting" not in category:
        return False
    for name in (shared_params.PARAM_IS_GAS_FIXTURE,
                 shared_params.PARAM_IS_WATER_FIXTURE):
        if revit_helpers.get_parameter_value(element, name):
            return False
    # Piping connectors only (n=30): an electrical connector never makes a
    # one-connector cap look like a two-connector fitting.
    return len(revit_helpers.piping_connectors(
        revit_helpers.get_connectors(element))) == 1


# ---------------------------------------------------------------------------
# Existing conditions, remembered between Size Gas and One-Line
# ---------------------------------------------------------------------------
# The designer types the existing pipe size, length and load into Size Gas and
# would otherwise type them again into One-Line. Kept per picked pipe for as
# long as this module stays loaded. If pyRevit reloads it the box just opens
# empty, so nothing depends on this.
# 2026-11-20: no button calls remember_existing or recall_existing any more.
# The existing conditions are entered in the gas dialog's Scope before the
# pick, empty every run (Colin). Kept for its test until Colin decides.
_EXISTING_BY_PIPE_ID = {}


def remember_existing(pipe_id, existing):
    _EXISTING_BY_PIPE_ID[pipe_id] = dict(existing)


def recall_existing(pipe_id):
    found = _EXISTING_BY_PIPE_ID.get(pipe_id)
    return dict(found) if found is not None else None


def existing_from_scope(scope_existing):
    """The existing dict the gas tools carry after a PIPE pick, from the
    dialog's Scope (Colin, 2026-11-20).

    scope_existing: the dialog's "existing" value, None for New Construction
    or a dict with "size_nominal", "length_ft" and "load_mbh" for Connect
    into Existing. Returns {"new_construction": True} for New Construction,
    else a copy of the entered values with "new_construction": False.
    """
    if not scope_existing:
        return {"new_construction": True}
    existing = dict(scope_existing)
    existing["new_construction"] = False
    return existing


NEW_CONSTRUCTION_TEXT = ("New construction. The system starts at the picked "
                         "pipe, with no existing piping upstream.")


def is_new_construction(existing):
    """True when the existing dict says New Construction (a pipe pick with
    the New Construction scope, tie_in.existing_from_scope)."""
    return bool(existing) and bool(existing.get("new_construction"))


def existing_for_sizing(existing):
    """The existing dict to apply to the graph, or None for new construction
    (nothing to apply, nothing to check)."""
    if existing is None or is_new_construction(existing):
        return None
    return existing


def existing_inputs_text(existing):
    """One line for the inputs table: the entered existing piping, or the
    new construction statement."""
    if is_new_construction(existing):
        return NEW_CONSTRUCTION_TEXT
    return ("{0}\" pipe, {1:.1f} ft from the meter to the new connection, "
            "{2:.1f} MBH existing load.".format(
                existing["size_nominal"], float(existing["length_ft"]),
                float(existing["load_mbh"])))
