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
    return len(revit_helpers.get_connectors(element)) == 1


# ---------------------------------------------------------------------------
# Existing conditions, remembered between Size Gas and One-Line
# ---------------------------------------------------------------------------
# The designer types the existing pipe size, length and load into Size Gas and
# would otherwise type them again into One-Line. Kept per picked pipe for as
# long as this module stays loaded. If pyRevit reloads it the box just opens
# empty, so nothing depends on this.
_EXISTING_BY_PIPE_ID = {}


def remember_existing(pipe_id, existing):
    _EXISTING_BY_PIPE_ID[pipe_id] = dict(existing)


def recall_existing(pipe_id):
    found = _EXISTING_BY_PIPE_ID.get(pipe_id)
    return dict(found) if found is not None else None
