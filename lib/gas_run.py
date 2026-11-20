# -*- coding: ascii -*-
# gas_run.py
# The Size Gas pipeline (IFGC Longest Run Method), shared by the Size Gas
# button and the headless entry (pyRevit agent runtime).
#
#   dialog (Scope: New Construction or Connect into Existing) -> pick meter
#   or first new pipe -> traverse -> PRV zones -> size -> report -> pre
#   sizing warnings (yes/no) -> write pipes, fittings, fixture stubs (three
#   steps, each its own transaction through revit_helpers.run_in_transaction)
#
# The button path is run(); the agent path is run_headless(). Both go through
# resolve_start() and size_core(), so there is one code path after the pick.
# The step bodies follow Size Gas.pushbutton/script.py on main (gas review
# fixes, 2026-11-20); only `output`, the yes/no question and the transaction
# runner are passed in.
#
# The IFGC table has NO default: the delivery pressure, pipe material and
# allowable drop are asked at the start of every gas job (Colin, 2026-10-13).
#
# IronPython 2.7 / pyRevit, CPython 3 for the tests.

import console_report as cr
import gas_console
import shared_params
import revit_helpers
import pipe_graph
import report_generator
import gas_tables
import sizing_engine
import tie_in


TOOL = "size_gas"
START_METER = "meter"
START_PIPE = "pipe"
DETAILS = ("flagged", "all")
QID_WARNINGS = "presizing_warnings"
QIDS = (QID_WARNINGS,)
TITLE = "Size Gas"

PIPE_STEP = "pipe sizes"
FITTING_STEP = "fittings"
STUB_STEP = "fixture stub pipes"

# The pipe write approach that worked first this run (shown in the Summary).
_confirmed_approach = [None]


class _Console(object):
    """The themed console for the button, plain markdown for a headless
    recorder. themed=True sends each section through output.print_html (the
    markdown twin prints if that fails); themed=False prints the markdown
    twin only, so a headless log stays readable text."""

    def __init__(self, output, themed=False):
        self.output = output
        self.themed = themed

    def emit(self, html, md):
        if self.themed:
            cr.emit(self.output, html, lambda: self.output.print_md(md),
                    None, TITLE)
        else:
            self.output.print_md(md)

    def say(self, text, kind="info"):
        self.emit(gas_console.message(text, kind),
                  gas_console.message_md(text, kind))


class _RjaPrompter(object):
    """Button prompter: the RJA themed popup (ui_helpers.rja_alert) in place
    of forms.alert, as main's Size Gas does."""

    def confirm(self, qid, text, title="", suggested="continue"):
        import ui_helpers
        return bool(ui_helpers.rja_alert(TITLE, text, "warn", yes_no=True))

    def notify(self, text, title="", **kwargs):
        import ui_helpers
        ui_helpers.rja_alert(TITLE, text, "info")


# ---------------------------------------------------------------------------
# Button path
# ---------------------------------------------------------------------------

def run(doc, uidoc, output, forms):
    """Size Gas behind the button: dialog (with its Scope), pick, core.
    Returns True if it finished. forms is kept for the button's signature;
    every popup is the RJA themed ui_helpers.rja_alert."""
    import ui_helpers

    con = _Console(output, themed=True)
    revit_helpers.clear_log()

    choice = ui_helpers.show_size_gas_dialog("Size Gas - Select IFGC Table")
    if choice is None:
        con.say("Cancelled at table selection. No changes were made to the "
                "model.")
        return False
    settings = settings_from_dialog(choice)

    # Scope, chosen in the dialog before the pick (Colin, 2026-11-20): None
    # is New Construction, else the existing gas piping entered there.
    connect_existing = choice.get("existing")

    # the shared pick step, same words and window as One-Line (2026-11-19);
    # Connect into Existing takes the first new pipe only
    selected_element = ui_helpers.pick_element(
        uidoc, doc, "Size Gas",
        ui_helpers.PICK_GAS_EXISTING if connect_existing
        else ui_helpers.PICK_GAS)
    if selected_element is None:
        pick_error = ui_helpers.last_pick_error()
        if pick_error:
            con.say("The pick failed: {}. No changes were made to the "
                    "model.".format(pick_error), "error")
        else:
            con.say("Selection cancelled. No changes were made to the "
                    "model.")
        return False

    start = resolve_start(doc, selected_element, settings)
    if start["stop"]:
        ui_helpers.rja_alert(TITLE, start["stop"], "error")
        return False

    inputs = {"selected": "Element ID {}".format(
        revit_helpers.eid_int(selected_element.Id)),
        "check": start["check"]}
    if start["existing_text"]:
        inputs["existing"] = start["existing_text"]

    outcome = size_core(doc, settings, selected_element, start["upstream"],
                        start["existing"], output, _RjaPrompter(),
                        write=True, themed=True, inputs_extra=inputs)
    if outcome["stop"]:
        ui_helpers.rja_alert(TITLE, outcome["stop"],
                             outcome.get("stop_kind") or "error")
    return outcome["status"] == "ok"


def _propane_heat_content(choice):
    """The propane heat content from the dialog choice, BTU/CF. The key is
    'propane_heat_content_btu_per_cf'; a dialog without the box gives the
    firm default (Colin, 2026-11-20)."""
    value = choice.get("propane_heat_content_btu_per_cf")
    if value is None:
        return shared_params.DEFAULT_PROPANE_HEAT_CONTENT_BTU_PER_CF
    return float(value)


def settings_from_dialog(choice):
    """The dialog's answer as headless settings (same keys as
    default_settings). existing is the dialog's Scope value: None for New
    Construction, else {size_nominal, length_ft, load_mbh}."""
    material = choice["pipe_material"]
    opt = gas_tables.get_table_option_by_material_and_short_label(
        material, choice["table_label"])
    down = None
    if choice.get("mid_stream_prv"):
        down = gas_tables.get_table_option_by_material_and_short_label(
            material, choice["downstream_table_label"])["table_id"]
    return {
        "table_id": opt["table_id"],
        "pipe_material": material,
        "heat_content_btu_per_cf": float(choice["heat_content"]),
        "propane_heat_content_btu_per_cf": _propane_heat_content(choice),
        "downstream_table_id": down,
        "existing": choice.get("existing"),
    }


# ---------------------------------------------------------------------------
# Settings (headless)
# ---------------------------------------------------------------------------

def table_options():
    """Every IFGC table the dialog offers, for the agent to ask Colin from."""
    return [{"table_id": o["table_id"], "material": o["material"],
             "gas": o["gas"], "inlet_pressure_psi": o["inlet_pressure_psi"],
             "label": o["label"]} for o in gas_tables.TABLE_OPTIONS]


def default_settings():
    """What the dialog prefills. table_id is None on purpose: the table is
    asked at the start of every gas job, never assumed. The propane heat
    content is used only on the propane CFH Table 402.4(37)."""
    return {
        "table_id": None,
        "pipe_material": None,
        "heat_content_btu_per_cf":
            float(shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF),
        "propane_heat_content_btu_per_cf":
            float(shared_params.DEFAULT_PROPANE_HEAT_CONTENT_BTU_PER_CF),
        "downstream_table_id": None,
        "existing": None,
    }


def merge_settings(overrides):
    """default_settings with the caller's values on top. An unknown key is an
    error, so a typo never runs silently on the default."""
    settings = default_settings()
    unknown = sorted(set(overrides or {}) - set(settings))
    if unknown:
        raise ValueError("unknown Size Gas setting(s): {}".format(
            ", ".join(unknown)))
    settings.update(overrides or {})
    return settings


def _option(table_id, material=None):
    """The TABLE_OPTIONS entry for a table id (and material when given)."""
    for o in gas_tables.TABLE_OPTIONS:
        if o["table_id"] == table_id and (material is None
                                          or o["material"] == material):
            return o
    return None


def validate_settings(settings):
    """Problems as text, empty when the settings can run. Same rules as the
    dialog: a known table, both heat contents in the plausible range
    (gas_tables.read_heat_content, 500 to 3,000 BTU/CF), a downstream table
    of the same material and gas that is not a higher pressure."""
    problems = []
    up = None
    if not settings.get("table_id"):
        problems.append(
            "table_id is required: ask the delivery pressure, pipe material "
            "and allowable pressure drop (gas_run.table_options())")
    else:
        up = _option(settings["table_id"], settings.get("pipe_material"))
        if up is None:
            problems.append("table_id {} is not an offered IFGC table{}".format(
                settings["table_id"],
                " for {}".format(settings["pipe_material"])
                if settings.get("pipe_material") else ""))
    for key in ("heat_content_btu_per_cf", "propane_heat_content_btu_per_cf"):
        _value, message = gas_tables.read_heat_content(settings.get(key), key)
        if message:
            problems.append(message)

    down_id = settings.get("downstream_table_id")
    if down_id and up is not None:
        down = _option(down_id, up["material"])
        if down is None or down["gas"] != up["gas"]:
            problems.append(
                "downstream_table_id {} shall be a {} {} table".format(
                    down_id, up["material"], up["gas"]))
        elif down["inlet_pressure_psi"] > up["inlet_pressure_psi"]:
            problems.append(
                "downstream_table_id {} is a higher pressure than {}".format(
                    down_id, up["table_id"]))

    existing = settings.get("existing")
    if existing is not None:
        problems.extend(_existing_problems(existing, settings.get("table_id")))
    return problems


def _existing_problems(existing, table_id):
    problems = []
    if not isinstance(existing, dict):
        return ["existing shall be {size_nominal, length_ft, load_mbh}"]
    if tie_in.is_new_construction(existing):
        return problems
    missing = [k for k in ("size_nominal", "length_ft", "load_mbh")
               if k not in existing]
    if missing:
        return ["existing is missing {}".format(", ".join(missing))]
    if table_id and _option(table_id) is not None:
        sizes = gas_tables.list_pipe_sizes(table_id)
        if existing["size_nominal"] not in sizes:
            problems.append(
                "existing size_nominal {} is not in Table {} ({})".format(
                    existing["size_nominal"], table_id, ", ".join(sizes)))
    for k in ("length_ft", "load_mbh"):
        try:
            if float(existing[k]) < 0:
                problems.append("existing {} shall be 0 or more".format(k))
        except (TypeError, ValueError):
            problems.append("existing {} is not a number".format(k))
    return problems


# ---------------------------------------------------------------------------
# Start point (shared by the pick and run_headless)
# ---------------------------------------------------------------------------

def _connect_existing(settings):
    """The entered existing piping when the run connects into existing
    (the dialog's Connect into Existing), else None (New Construction)."""
    existing = settings.get("existing")
    if not existing or tie_in.is_new_construction(existing):
        return None
    return existing


def resolve_start(doc, element, settings):
    """Check the picked meter or first new pipe and find where the walk
    begins, through the one gas start check (tie_in.check_gas_start).

    settings["existing"] is the Scope: None or {"new_construction": True}
    is New Construction (a pipe pick starts the system with no existing
    piping, Colin 2026-11-20); {size_nominal, length_ft, load_mbh} is
    Connect into Existing, which takes the first new pipe only.
    Returns dict: kind ("meter" | "pipe"), upstream, existing, check (one line
    for the inputs table), existing_text (inputs line, or None), stop (None
    to go on, or the message that stopped the run), title.
    """
    out = {"kind": START_METER, "upstream": None, "existing": None,
           "check": None, "existing_text": None,
           "stop": None, "title": "Size Gas - Invalid Selection"}
    if element is None:
        out["stop"] = "Could not retrieve selected element."
        out["title"] = "Size Gas - Selection Error"
        return out

    connect_existing = _connect_existing(settings)
    start = tie_in.check_gas_start(element, doc,
                                   pipe_only=bool(connect_existing))
    if start["kind"] == "tie_in" or revit_helpers.is_pipe(element):
        out["kind"] = START_PIPE
    if not start["ok"]:
        out["stop"] = start["message"] + "\n\nNothing was changed."
        return out
    out["upstream"] = start["upstream"]
    out["check"] = start["message"]
    if start["kind"] == "tie_in":
        # New Construction from a pipe: full new construction, new meter.
        existing = tie_in.existing_from_scope(connect_existing)
        out["existing"] = existing
        out["existing_text"] = tie_in.existing_inputs_text(existing)
    return out


# ---------------------------------------------------------------------------
# Run checks
# ---------------------------------------------------------------------------

def _table_limit_check(table_id, graph, heat_content, propane_heat_content):
    """Table Intended Use load limit warnings (2024 IFGC Table 402.4(13):
    150 CFH through the house line regulator). The meter carries the whole
    system, so the total is the new fixtures plus the existing load entered
    for Connect into Existing (n=35/52, Colin 2026-11-20)."""
    total_mbh = pipe_graph.total_connected_mbh(graph)
    existing = getattr(graph, "existing", None) or {}
    total_mbh += float(existing.get("load_mbh") or 0.0)
    demand, _units = sizing_engine.demand_in_table_units(
        total_mbh, table_id, heat_content, propane_heat_content)
    return sizing_engine.table_limit_warnings(table_id, demand, total_mbh)


def _graph_notices(graph, result):
    """(text, kind) lines for the pipes left unsized because they lead only
    to 0 MBH (data error) fixtures. The data errors and tapped pipes
    themselves print in the pre-sizing warnings (gas_console)."""
    lines = []
    for z in result.get("unsized_zero_load") or []:
        lines.append((
            "Pipe {} leads only to 0 MBH fixture(s) {} and was NOT sized or "
            "written. Fix the fixture load and run Size Gas again.".format(
                z["pipe_id"], ", ".join(str(i) for i in z["fixture_ids"])),
            "error"))
    return lines


# ---------------------------------------------------------------------------
# Core (one code path after the pick)
# ---------------------------------------------------------------------------

def size_core(doc, settings, element, upstream, existing, output, prompter,
              write=True, run_step=None, themed=False, inputs_extra=None):
    """Traverse, zone, size, report and (write=True) write the sizes.

    The yes/no pre sizing warnings question goes through prompter.confirm
    (qid "presizing_warnings") and is asked only when writing; with
    HeadlessPrompter it raises NeedsDecision before anything is written.
    run_step(doc, name, action) runs one write step and returns None when it
    committed, else the error text (default revit_helpers.run_in_transaction,
    warnings deleted, no Revit dialog). themed=True (the button) prints the
    RJA themed console sections through gas_console; themed=False prints
    their markdown twins. inputs_extra adds the pick lines (selected, check,
    existing) to the inputs section.

    Returns dict: status ("ok" | "cancelled" | "refused"), stop (message
    that stopped the run, or None), stop_title, stop_kind (the popup kind),
    graph, result, existing_check, warnings (dict of lists), writes (None
    when nothing was written).
    """
    outcome = {"status": "ok", "stop": None, "stop_title": "Size Gas",
               "stop_kind": "error", "graph": None, "result": None,
               "existing_check": None, "warnings": {}, "writes": None}
    con = _Console(output, themed)

    def _alert(text, title, kind="error"):
        """The stop the caller shows: the button's popup, the headless
        envelope's first error."""
        outcome["status"] = "refused"
        outcome["stop"] = text
        outcome["stop_title"] = title
        outcome["stop_kind"] = kind
        return outcome

    _confirmed_approach[0] = None
    # revit_helpers keeps its own cache and stays loaded between runs
    revit_helpers.reset_pipe_diameter_approach()

    table_id = settings["table_id"]
    selected_opt = _option(table_id, settings.get("pipe_material"))
    pipe_material = selected_opt["material"]
    inlet_pressure_psi = selected_opt["inlet_pressure_psi"]
    heat_content_btu_per_cf = float(settings["heat_content_btu_per_cf"])
    propane_heat_content = settings.get("propane_heat_content_btu_per_cf")
    if propane_heat_content is None:
        propane_heat_content = (
            shared_params.DEFAULT_PROPANE_HEAT_CONTENT_BTU_PER_CF)
    propane_heat_content = float(propane_heat_content)
    downstream_table_id = settings.get("downstream_table_id") or None
    mid_stream_prv = downstream_table_id is not None
    downstream_opt = (_option(downstream_table_id, pipe_material)
                      if mid_stream_prv else None)

    # The inputs section prints once the pick is known, so the table,
    # material, gas, heat content and start show together.
    inputs = {
        "lib": revit_helpers.lib_folder(),
        "material": pipe_material,
        "table": selected_opt["label"].split("  [")[0],
        "table_id": table_id,
        "gas": selected_opt["gas"],
        # gas_console.heat_content_text prints what the table used: the
        # propane value on 402.4(37), none on an MBH table (n=24)
        "heat_content": heat_content_btu_per_cf,
        "propane_heat_content": propane_heat_content,
        "downstream_table": (downstream_opt["label"].split("  [")[0]
                             if downstream_opt is not None else None),
        "downstream_table_id": downstream_table_id,
    }
    inputs.update(inputs_extra or {})
    con.emit(gas_console.sizegas_inputs(inputs),
             gas_console.sizegas_inputs_md(inputs))

    # STEP 4 - Traverse piping network
    try:
        graph = pipe_graph.build_network(element, doc, upstream)
    except Exception as e:
        con.say("Traversal ERROR: {}".format(str(e)), "error")
        return _alert("Traversal failed:\n\n{}".format(str(e)),
                      "Size Gas - Traversal Error")
    outcome["graph"] = graph

    trav = {"nodes": len(graph.nodes), "segments": len(graph.edges),
            "existing_connections": list(graph.existing_connections or [])}
    con.emit(gas_console.traversal(trav), gas_console.traversal_md(trav))
    if getattr(graph, "existing_appliances", None):
        appliance_rows = report_generator.existing_appliance_rows(graph)
        con.emit(gas_console.existing_appliances(appliance_rows),
                 gas_console.existing_appliances_md(appliance_rows))

    fixture_nodes = pipe_graph.gas_fixtures(graph)
    if len(fixture_nodes) == 0:
        return _alert(
            "No gas fixtures found (IS_GAS_FIXTURE = Yes).\n\n"
            "Common causes:\n"
            "  1. Pipe caps were inadvertently changed when resizing all "
            "piping in Revit -- verify cap families have not replaced fixture "
            "families at pipe terminations.\n"
            "  2. Fixture families are not physically connected to the piping.\n"
            "  3. IS_GAS_FIXTURE parameter is not set to Yes on the appliances.\n\n"
            "Run Diagnose to see which terminal elements were found.",
            "Size Gas - No Fixtures")

    zero_load = pipe_graph.zero_load_stop_message(graph)
    if zero_load is not None:
        con.say("Stopped: total connected load is 0 MBH. No changes were "
                "made to the model.", "error")
        return _alert(zero_load, "Size Gas - Zero Load")

    if graph.longest_run is None:
        return _alert("Could not determine longest run. "
                      "Run Diagnose first and resolve all errors.",
                      "Size Gas - Sizing Error")

    # Existing building: the existing length joins the developed length that
    # picks the table row, before anything is sized or zoned. New construction
    # has nothing to apply and nothing to check.
    new_construction = tie_in.is_new_construction(existing)
    if tie_in.existing_for_sizing(existing) is not None:
        pipe_graph.apply_existing_conditions(
            graph, existing["size_nominal"], existing["length_ft"],
            existing["load_mbh"])

    # STEP 4b - Pressure regulating valves (only a mid-stream one is a step
    # down; nothing changes unless a downstream table is given).
    min_ft = shared_params.PRV_MIDSTREAM_MIN_DOWNSTREAM_FT
    prv_nodes = [graph.nodes[i] for i in graph.prv_ids]
    midstream_prvs = [graph.nodes[i] for i in graph.midstream_prv_ids]
    prv_warnings = []

    # the same stop and words as One-Line (pipe_graph; the PRV keywords come
    # from shared_params.PRV_FAMILY_KEYWORDS)
    prv_stop = pipe_graph.midstream_prv_stop_message(graph, mid_stream_prv)
    if prv_stop is not None:
        return _alert(prv_stop, "Size Gas - No Mid-Stream PRV", "warn")

    if mid_stream_prv and graph.nested_prv_ids:
        return _alert(
            "Mid-stream PRV(s) {} sit downstream of another mid-stream "
            "PRV.\n\nOnly a single step down is supported.".format(
                ", ".join(str(i) for i in graph.nested_prv_ids)),
            "Size Gas - Nested PRVs")

    if prv_nodes:
        prv_rows = [{
            "id": n.element_id, "family": n.family_name,
            "downstream_ft": n.downstream_pipe_ft or 0.0,
            "treated": ("mid-stream step down" if n.is_midstream_prv else
                        "equipment PRV, ignored (within {:g} ft)".format(
                            min_ft))} for n in prv_nodes]
        con.emit(gas_console.prv_table(prv_rows),
                 gas_console.prv_table_md(prv_rows))
        if midstream_prvs and not mid_stream_prv:
            prv_warnings.append(
                "{} mid-stream PRV(s) found but Mid Stream PRV is unchecked - "
                "the whole system is sized as one system on the meter's "
                "table, including the piping after the regulator(s).".format(
                    len(midstream_prvs)))

    # STEP 5 - IFGC sizing calculation
    try:
        result = sizing_engine.size_system(
            graph, pipe_material, inlet_pressure_psi, table_id,
            heat_content_btu_per_cf=heat_content_btu_per_cf,
            downstream_table_id=downstream_table_id,
            propane_heat_content_btu_per_cf=propane_heat_content)
    except ValueError as e:
        con.say("Sizing ERROR: {}".format(str(e)), "error")
        return _alert("Sizing failed:\n\n{}".format(str(e)),
                      "Size Gas - Sizing Error")
    outcome["result"] = result

    # Sizing results: the basis, then one table per system.
    total_load = pipe_graph.total_connected_mbh(graph)
    units = gas_tables.table_units(result["table_id"])
    con.emit(gas_console.sizing_basis(result, total_load, len(fixture_nodes),
                                      units),
             gas_console.sizing_basis_md(result, total_load,
                                         len(fixture_nodes), units))
    rows = sizing_engine.sizing_rows(result, graph)
    con.emit(gas_console.sizing_tables(rows, result),
             gas_console.sizing_tables_md(rows, result))
    for text, kind in _graph_notices(graph, result):
        con.say(text, kind)

    existing_check = sizing_engine.check_existing_pipe(result, graph)
    outcome["existing_check"] = existing_check
    existing_warnings = []
    if existing_check is not None:
        con.emit(gas_console.existing_check(existing_check),
                 gas_console.existing_check_md(existing_check))
        if not existing_check["adequate"]:
            existing_warnings.append(_existing_warning(existing_check))

    # STEP 6 - Pre-sizing validation warning. 0 MBH fixtures are listed from
    # graph.data_errors only (the walk's one record), not tested again here.
    fixture_nodes = pipe_graph.gas_fixtures(graph)
    unnamed = [n for n in fixture_nodes
               if n.fixture_name in ("", "UNNAMED", None)]
    # Table Intended Use load limits (2024 IFGC Table 402.4(13): 150 CFH),
    # on the new load plus any existing load entered (n=35/52).
    table_warnings = _table_limit_check(
        table_id, graph, heat_content_btu_per_cf, propane_heat_content)
    non_gas = list(graph.non_gas_skipped)
    bare_end = pipe_graph.non_fixture_ends(graph)
    data_errors = list(getattr(graph, "data_errors", None) or [])
    tapped = list(getattr(graph, "tapped_pipes", None) or [])

    outcome["warnings"] = {
        "existing_pipe": existing_warnings,
        "prv": prv_warnings,
        "table_limit": table_warnings,
        "data_errors": [dict(d) for d in data_errors],
        "unsized_zero_load": list(result.get("unsized_zero_load") or []),
        "unnamed": [{"id": n.element_id, "mbh": float(n.gas_load_mbh)}
                    for n in unnamed],
        "non_gas": [{"id": i, "classification": c} for i, c in non_gas],
        "bare_end": [{"id": n.element_id, "family": n.family_name}
                     for n in bare_end],
        "disconnected": list(graph.disconnected),
        "tapped": tapped,
    }

    groups = gas_console.build_warning_groups(
        existing=existing_warnings, prv=prv_warnings,
        table_limit=table_warnings,
        unnamed=[(n.element_id, n.gas_load_mbh) for n in unnamed],
        non_gas=list(non_gas),
        bare_end=[(n.element_id, n.family_name) for n in bare_end],
        disconnected=list(graph.disconnected),
        n_fixtures=len(fixture_nodes),
        data_errors=data_errors,
        tapped=tapped)
    if groups:
        con.emit(gas_console.presize_warnings(groups),
                 gas_console.presize_warnings_md(groups))
        if write:
            proceed = prompter.confirm(
                QID_WARNINGS,
                "{} pre-sizing warning(s) found (see the output window). "
                "Proceed with sizing anyway?".format(len(groups)),
                title="Size Gas - Warnings", suggested="stop")
            if not proceed:
                con.say("Sizing cancelled at the pre-sizing warnings. No "
                        "changes were made to the model.")
                outcome["status"] = "cancelled"
                return outcome

    if not write:
        con.say("Plan only (write=False): nothing was written.")
        return outcome

    outcome["writes"] = _write_sizes(doc, con, graph, result, existing_check,
                                     new_construction, run_step)
    if outcome["writes"].get("pipe_transaction_error"):
        return _alert(
            "The pipe sizes could not be committed and Revit rolled them "
            "back. No size was written to the model.\n\n{}".format(
                outcome["writes"]["pipe_transaction_error"]),
            "Size Gas - Transaction Error")
    return outcome


def _existing_warning(existing_check):
    """The existing pipe shortfall text (capacity and demand in the table's
    units: CFH, or MBH on a propane table 402.4(25) to (36))."""
    if existing_check["capacity_cfh"] is None:
        return ("The entered existing pipe size {}\" is not in Table {} at "
                "{:.0f} ft, so it could not be checked.".format(
                    existing_check["size_nominal"], existing_check["table_id"],
                    float(existing_check["run_ft"])))
    units = existing_check.get("units", "CFH")
    if units == "MBH":
        load_text = "{:.0f} MBH".format(float(existing_check["total_load_mbh"]))
    else:
        load_text = "{:.0f} {} ({:.0f} MBH)".format(
            float(existing_check["demand_cfh"]), units,
            float(existing_check["total_load_mbh"]))
    return ("The existing {}\" pipe carries {:.0f} {} at {:.0f} ft "
            "developed length, but existing plus new load is {}. "
            "The existing pipe is NOT adequate and was not "
            "resized.".format(
                existing_check["size_nominal"],
                float(existing_check["capacity_cfh"]), units,
                float(existing_check["run_ft"]), load_text))


# ---------------------------------------------------------------------------
# Write steps. Each step is one transaction through run_step (default
# revit_helpers.run_in_transaction: warnings deleted, an error rolls the step
# back with no Revit dialog, and the Commit status is checked, n=4). A step
# that did not commit reports zero successes, and the steps that depend on
# it do not run.
# ---------------------------------------------------------------------------

def _run_step(run_step, doc, name, action):
    """run_step(doc, name, action), revit_helpers.run_in_transaction when
    none was given (read at call time)."""
    return (run_step or revit_helpers.run_in_transaction)(doc, name, action)


def _write_sizes(doc, con, graph, result, existing_check, new_construction,
                 run_step=None):
    """STEPS 7 to 9: pipes, fittings, fixture stubs. Returns the counts for
    the envelope; rolled_back names the steps Revit rolled back and
    transaction_errors gives their text."""
    approaches = ("RBS_PIPE_NOMINAL_DIAMETER, RBS_PIPE_DIAMETER_PARAM, "
                  "LookupParameter(Diameter)")
    con.emit(gas_console.write_start(approaches),
             gas_console.write_start_md(approaches))
    # Transitions are identified now, before the resize: Revit may delete one
    # while the pipes change size (Colin, 2026-11-19).
    transition_ids = _transition_ids(graph)

    rolled_back = []
    tx_errors = []
    w = _write_pipe_sizes(doc, graph, result, run_step)
    if not w["committed"]:
        rolled_back.append(PIPE_STEP)
        tx_errors.append("{}: {}".format(PIPE_STEP, w["error"]))
        con.say("Pipe transaction rolled back: {}".format(w["error"]),
                "error")

    # The existing pipe audit can never stop the fitting and stub steps
    # (Colin, 2026-11-02).
    try:
        # new construction: print only if the walk still met existing phase
        # pipes, so a skip is never silent
        if not new_construction or w["existing_skip_count"]:
            con.emit(gas_console.existing_audit(w["existing_skip_count"],
                                                w["existing_overloaded"],
                                                w["existing_unchecked"]),
                     gas_console.existing_audit_md(w["existing_skip_count"],
                                                   w["existing_overloaded"],
                                                   w["existing_unchecked"]))
    except Exception as e:
        con.say("Existing pipe audit could not be printed: {}".format(str(e)),
                "warn")

    # STEP 8 - fittings, only when the pipe sizes committed. A new fitting
    # that joins an existing pipe is never sized below it; transitions are
    # left to Revit; the meter and PRVs are never resized.
    existing_sizes = {}
    if w["committed"]:
        existing_sizes = _existing_pipe_sizes(graph, w["new_sizes"], result,
                                              doc)
    fit, fit_committed = _write_fitting_sizes(
        doc, graph, w["new_sizes"], transition_ids, existing_sizes,
        w["committed"], run_step)
    if w["committed"] and not fit_committed:
        rolled_back.append(FITTING_STEP)
        tx_errors.extend(fit["failures"])

    # STEP 9 - fixture stub pipes, after their caps were sized in STEP 8, so
    # only when STEP 8 committed.
    st = _write_stub_sizes(doc, graph, w["skipped_stubs"], fit_committed,
                           run_step)
    if st["rolled_back"]:
        rolled_back.append(STUB_STEP)
        tx_errors.extend(r["reason"] for r in st["rows"])

    try:
        failures = gas_console.failure_rows(
            w["fail_rows"], fit.get("failure_rows") or [], st["rows"])
        wr = {"pipes_sized": w["success_count"], "stubs_sized": st["success"],
              "fit": fit, "failures": failures}
        con.emit(gas_console.write_results(wr),
                 gas_console.write_results_md(wr))
    except Exception as e:
        con.say("Write results could not be printed: {}. See the summary "
                "below.".format(str(e)), "warn")

    try:
        _print_summary(con, result, existing_check, w, st, fit, rolled_back)
    except Exception as e:
        con.say("Summary could not be printed: {}. Check the write results "
                "above.".format(str(e)), "warn")

    return {
        "pipes_sized": w["success_count"],
        "pipe_failures": ["Pipe {}: {}".format(r["id"], r["reason"])
                          for r in w["fail_rows"]],
        "stubs_sized": st["success"], "stub_failures": st["fail"],
        "existing_skipped": w["existing_skip_count"],
        "existing_overloaded": w["existing_overloaded"],
        "existing_unchecked": w["existing_unchecked"],
        "fittings": fit,
        "approach": _confirmed_approach[0] if w["committed"] else None,
        "written_ids": list(w["written"]) + list(st["written"]),
        "pipe_transaction_error": w["error"],
        "rolled_back": rolled_back,
        "transaction_errors": tx_errors,
    }


def _new_pipe_writes():
    return {"success_count": 0, "fail_count": 0, "fail_rows": [],
            "skipped_stubs": [], "existing_skip_count": 0,
            "existing_overloaded": [], "existing_unchecked": [],
            "new_sizes": {}, "written": [], "committed": False,
            "error": None}


def _write_one_pipe(w, pipe_id, nominal_size, graph, result, doc):
    """Write or queue one pipe of result["sizes"] (STEP 7)."""
    edge = graph.edges.get(pipe_id)
    if edge is None or edge.pipe is None:
        w["fail_rows"].append({
            "id": pipe_id, "family": "Pipe",
            "reason": "edge or pipe element not found in graph."})
        w["fail_count"] += 1
        return

    # One test for new construction, failing closed: an unreadable phase is
    # existing and is never resized (Colin, 2026-11-02).
    is_new = revit_helpers.is_new_construction(edge.pipe, doc)

    to_node = graph.nodes.get(edge.to_node_id)
    if to_node is not None and to_node.is_gas_fixture:
        # An existing phase appliance fed by new piping keeps its stub pipe
        # as drawn: it is audited, never resized (review n=26), so it is
        # not in new_sizes either (no fitting is sized from it).
        existing_appliance = to_node.element_id in (
            getattr(graph, "existing_appliances", None) or [])
        if is_new and not existing_appliance:
            w["new_sizes"][pipe_id] = nominal_size
            # New construction stub: queued for STEP 9 (cap first resize)
            w["skipped_stubs"].append({
                "pipe_id": pipe_id,
                "fixture_name": to_node.fixture_name or "UNNAMED",
                "demand_mbh": edge.cumulative_load_mbh,
                "recommended_size": nominal_size,
            })
        else:
            w["existing_skip_count"] += 1
            _audit_existing_pipe(pipe_id, edge, nominal_size, result,
                                 w["existing_overloaded"],
                                 w["existing_unchecked"])
        return

    if not is_new:
        w["existing_skip_count"] += 1
        _audit_existing_pipe(pipe_id, edge, nominal_size, result,
                             w["existing_overloaded"], w["existing_unchecked"])
        return
    w["new_sizes"][pipe_id] = nominal_size

    nominal_inches = sizing_engine.NOMINAL_TO_INCHES.get(nominal_size)
    if nominal_inches is None:
        w["fail_rows"].append({
            "id": pipe_id, "family": "Pipe",
            "reason": "unrecognised nominal size '{}'.".format(nominal_size)})
        w["fail_count"] += 1
        return

    ok, _approach = _set_pipe_diameter(edge.pipe, nominal_inches)
    if ok:
        w["success_count"] += 1
        w["written"].append(pipe_id)
    else:
        w["fail_count"] += 1
        w["fail_rows"].append({
            "id": pipe_id, "family": "Pipe",
            "reason": "all API approaches failed ({}\").".format(
                nominal_size)})


def _write_pipe_sizes(doc, graph, result, run_step=None):
    """STEP 7: write the new construction pipe sizes in one transaction.
    Returns the _new_pipe_writes() dict, with committed False and zero
    successes when Revit rolled the transaction back."""
    w = _new_pipe_writes()

    def action():
        for pipe_id, nominal_size in result["sizes"].items():
            _write_one_pipe(w, pipe_id, nominal_size, graph, result, doc)

    err = _run_step(run_step, doc, "RJA Tools - Size Gas Pipes", action)
    if err is None:
        w["committed"] = True
        return w
    revit_helpers._log_entry("ERROR", "_write_pipe_sizes", None,
                             "Pipe transaction rolled back: {}".format(err))
    attempted = w["success_count"] + w["fail_count"]
    w["success_count"] = 0
    w["fail_count"] = attempted
    w["written"] = []
    w["error"] = err
    w["fail_rows"].append({
        "id": "-", "family": "All pipes",
        "reason": "transaction rolled back, no pipe size was written: "
                  "{}".format(err)})
    return w


def _write_fitting_sizes(doc, graph, new_sizes, transition_ids,
                         existing_sizes, pipes_committed, run_step=None):
    """STEP 8: resize the fittings in one transaction, only when the pipe
    sizes committed. Returns (fit counts, committed)."""
    if not pipes_committed:
        fit = _new_fit_counts()
        fit["failures"].append(
            "Fittings not resized: the pipe sizes were rolled back.")
        fit["failure_rows"].append({
            "id": "-", "family": "All fittings",
            "reason": "not run, the pipe sizes were rolled back."})
        return fit, False

    holder = {"fit": _new_fit_counts()}

    def action():
        holder["fit"] = _resize_fittings(graph, new_sizes, doc,
                                         transition_ids, existing_sizes)

    err = _run_step(run_step, doc, "RJA Tools - Resize Pipe Fittings", action)
    if err is None:
        return holder["fit"], True
    revit_helpers._log_entry("ERROR", "_write_fitting_sizes", None,
                             "Fitting transaction rolled back: {}".format(err))
    fit = _new_fit_counts()
    fit["failures"].append("Fitting transaction rolled back: {}".format(err))
    fit["failure_rows"].append({
        "id": "-", "family": "All fittings",
        "reason": "transaction rolled back: {}".format(err)})
    return fit, False


def _write_stub_sizes(doc, graph, skipped_stubs, fittings_committed,
                      run_step=None):
    """STEP 9: size the fixture stub pipes in one transaction, after their
    caps were sized in STEP 8, so only when STEP 8 committed.

    Returns {"success", "fail", "rows", "rolled_back", "written"}. A rolled
    back or skipped step counts every stub as failed (n=49)."""
    st = {"success": 0, "fail": 0, "rows": [], "rolled_back": False,
          "written": []}
    if not skipped_stubs:
        return st
    if not fittings_committed:
        st["fail"] = len(skipped_stubs)
        st["rows"].append({
            "id": "-", "family": "Fixture stub pipes",
            "reason": "not run: the fittings and caps were not sized first "
                      "(their transaction did not commit)."})
        return st

    def action():
        for s in skipped_stubs:
            # One guard per stub: one bad stub never rolls back the rest.
            try:
                if _size_one_stub(s, graph):
                    st["success"] += 1
                    st["written"].append(s["pipe_id"])
                else:
                    st["fail"] += 1
                    st["rows"].append({
                        "id": s["pipe_id"], "family": "Fixture stub pipe",
                        "reason": "size {}\" not written (see log)".format(
                            s["recommended_size"])})
            except Exception as e:
                st["fail"] += 1
                revit_helpers._log_entry(
                    "ERROR", "size_gas_stubs", s["pipe_id"],
                    "Stub resize raised: {}".format(str(e)))
                st["rows"].append({
                    "id": s["pipe_id"], "family": "Fixture stub pipe",
                    "reason": str(e)})

    err = _run_step(run_step, doc, "RJA Tools - Size Fixture Stub Pipes",
                    action)
    if err is None:
        return st
    revit_helpers._log_entry("ERROR", "_write_stub_sizes", None,
                             "Stub transaction rolled back: {}".format(err))
    st["success"] = 0
    st["fail"] = len(skipped_stubs)
    st["rolled_back"] = True
    st["written"] = []
    # the per-stub rows describe writes Revit undid: only the rollback row
    # stands
    st["rows"] = [{"id": "-", "family": "Fixture stub pipes",
                   "reason": "transaction rolled back: {}".format(err)}]
    return st


def _new_fit_counts():
    """The counts _resize_fittings fills. failures holds the text lines,
    failure_rows the same failures as data for the console table. prv counts
    the meter and PRVs left as drawn."""
    return {"resized": 0, "no_sized_pipe": 0, "existing": 0,
            "no_element": 0, "appliance": 0, "replaced": 0,
            "transitions": 0, "transitions_removed": 0, "prv": 0,
            "failures": [], "failure_rows": []}


def _fit_fail(counts, node, family, reason):
    """Record one fitting failure, as a text line and as a table row."""
    counts["failures"].append("Fitting {} ({}): {}".format(
        node.element_id, family, reason))
    counts["failure_rows"].append({"id": node.element_id,
                                   "family": family or "",
                                   "reason": reason})


def _transition_ids(graph):
    """Ids of the fittings that are transitions, read BEFORE the pipes are
    resized: Revit may delete a transition during the resize, after which
    its part type can no longer be read."""
    ids = set()
    for node in graph.nodes.values():
        if node.node_type not in ("tee", "fitting", "elbow", "fixture"):
            continue
        if node.is_gas_fixture:
            continue
        if node.is_transition or revit_helpers.is_transition(node.element):
            ids.add(node.element_id)
    return ids


def _set_pipe_diameter(pipe, nominal_inches):
    """revit_helpers.set_pipe_diameter, recording the approach that worked
    first (shown in the Summary). Nothing prints per pipe."""
    ok, name = revit_helpers.set_pipe_diameter(pipe, nominal_inches)
    if ok and _confirmed_approach[0] is None:
        _confirmed_approach[0] = name
    return ok, name


def _existing_pipe_sizes(graph, new_sizes, result, doc):
    """{pipe id: modeled nominal size or None} for the existing phase pipes
    in the graph, read from the drawn diameter on the table of each pipe's
    own system (sizing_engine.edge_table_id, by its run_key). A new fitting
    is never sized below these (n=3). An open ended pipe has no system table
    and is skipped (logged), never read on the meter table."""
    sizes = {}
    for pipe_id, edge in graph.edges.items():
        if pipe_id in new_sizes or edge.pipe is None:
            continue
        if not revit_helpers.is_valid_element(edge.pipe):
            continue
        if revit_helpers.is_new_construction(edge.pipe, doc):
            continue    # a new pipe with no size this run (open end)
        table_id = sizing_engine.edge_table_id(result, edge)
        if table_id is None:
            revit_helpers._log_entry(
                "INFO", "_existing_pipe_sizes", pipe_id,
                "Existing pipe with no downstream node (or no system "
                "table): its size is not read.")
            continue
        sizes[pipe_id] = _get_pipe_nominal_size(edge.pipe, table_id)
    return sizes


def _resize_fittings(graph, new_sizes, doc, transition_ids=None,
                     existing_sizes=None):
    """Resize the new construction fittings to the pipes sized this run.

    For each fitting the target size is the highest-load connected pipe in
    new_sizes; a new fitting that also joins existing phase pipes takes the
    largest of all its pipes, the existing ones at their modeled size from
    existing_sizes, so a new tee on an existing main never shrinks the main
    (n=3). A valve with a transition on each side reads the pipes past them.
    An existing phase fitting is never resized (Colin, 2026-11-02); the meter
    and every PRV are never resized (n=29). Transitions are never resized
    here (Colin, 2026-11-19): Revit updates or removes them with their
    pipes. transition_ids is from _transition_ids(), taken before the pipe
    resize. Only 'Nominal Radius' is written; Connector.Radius writes
    disconnect fittings (verified 2026-10-12). One guard per fitting (G7C,
    2026-11-04). A "joint" node (no Revit element) is never in this list.
    """
    counts = _new_fit_counts()
    transition_ids = transition_ids or set()
    done_ids = set()
    for node in graph.nodes.values():
        if node.node_type not in ("tee", "fitting", "elbow", "fixture"):
            continue
        try:
            _resize_one_fitting(node, graph, new_sizes, doc, counts, done_ids,
                                transition_ids, existing_sizes)
        except Exception as e:
            _fit_fail(counts, node, node.family_name, str(e))
            revit_helpers._log_entry(
                "ERROR", "_resize_fittings", node.element_id,
                "Resize raised: {}".format(str(e)))
    return counts


def _resize_one_fitting(node, graph, new_sizes, doc, counts, done_ids,
                        transition_ids=(), existing_sizes=None):
    """Resize one fitting node. A fitting Revit deleted when its pipes were
    resized is found again on the connected pipes and sized instead."""
    if node.element is None:
        counts["no_element"] += 1
        return

    # The meter and every PRV keep the size drawn (n=29).
    exempt = sizing_engine.fitting_resize_exempt(node, graph)
    if exempt:
        counts["prv"] += 1
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "Not resized: {}.".format(exempt))
        return

    element = node.element
    replaced = False
    if not revit_helpers.is_valid_element(element):
        element = revit_helpers.current_element(doc, node.element_id)
    if element is None:
        pipes = [e.pipe for e in graph.edges.values()
                 if node.element_id in (e.from_node_id, e.to_node_id)
                 and e.pipe is not None]
        # A transition Revit removed because its pipes became the same size
        # is the expected result, not a failure: its pipes now join.
        if (node.element_id in transition_ids
                and (revit_helpers.pipes_joined(pipes)
                     or revit_helpers.ends_connected(pipes,
                                                     node.location_xyz))):
            counts["transitions_removed"] += 1
            revit_helpers._log_entry(
                "INFO", "_resize_fittings", node.element_id,
                "Transition removed or replaced by Revit, its pipe ends "
                "are connected.")
            return
        element, _d = revit_helpers.find_replacement_fitting(
            pipes, node.location_xyz, set(graph.nodes.keys()) | done_ids)
        if element is None:
            _fit_fail(counts, node, node.family_name,
                      "deleted by Revit when its pipes were resized, and no "
                      "replacement fitting was found on its pipes.")
            return
        replaced = True
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "Deleted by Revit; replacement {} sized instead.".format(
                revit_helpers.eid_int(element.Id)))

    # Transitions are left to Revit (a replacement counts the same way).
    if revit_helpers.is_transition(element):
        counts["transitions"] += 1
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "Transition: size comes from its pipes, left to Revit.")
        return

    if not revit_helpers.is_new_construction(element, doc):
        counts["existing"] += 1
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "Existing phase (or phase unreadable): not resized.")
        return

    # The highest-load new pipe that touches it sets the size, never below
    # an existing pipe it joins; a valve with a transition on each side
    # reads the pipes past them (Horse Creek).
    target_size, via_transition = sizing_engine.fitting_target_size(
        node.element_id, graph, new_sizes, transition_ids, existing_sizes)
    if via_transition:
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "No sized pipe touches it; size {} read through a transition."
            .format(target_size))

    if target_size is None:
        counts["no_sized_pipe"] += 1
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "No new construction pipe sized this run touches it: not "
            "resized.")
        return

    # An appliance is resized only when it is a cap style family with a
    # 'Nominal Radius' (the cap first, then the stub pipe, STEP 9).
    if node.is_gas_fixture and _lookup_param(
            element, "Nominal Radius") is None:
        counts["appliance"] += 1
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "Appliance with no 'Nominal Radius': not resized.")
        return

    nominal_inches = sizing_engine.NOMINAL_TO_INCHES.get(target_size)
    if nominal_inches is None:
        _fit_fail(counts, node, node.family_name,
                  "unrecognised nominal size '{}'.".format(target_size))
        return

    # piping connectors only on a Type swap (n=30): no pipe end is moved
    # onto an electrical connector of a gas valve
    ok, approach = revit_helpers.set_fitting_size(element, nominal_inches,
                                                  piping_only=True)
    if ok:
        counts["resized"] += 1
        if replaced:
            counts["replaced"] += 1
            done_ids.add(revit_helpers.eid_int(element.Id))
    else:
        _fit_fail(counts, node, node.family_name,
                  "could not set size to {}\" ({}).".format(
                      target_size, approach))


def _size_one_stub(s, graph):
    """Size one fixture stub pipe (STEP 9). True when it was written."""
    edge = graph.edges.get(s["pipe_id"])
    if (edge is None or edge.pipe is None
            or not revit_helpers.is_valid_element(edge.pipe)):
        revit_helpers._log_entry(
            "ERROR", "size_gas_stubs", s["pipe_id"],
            "Stub pipe NOT FOUND in graph, or deleted.")
        return False
    nominal_inches = sizing_engine.NOMINAL_TO_INCHES.get(
        s["recommended_size"])
    if nominal_inches is None:
        revit_helpers._log_entry(
            "ERROR", "size_gas_stubs", s["pipe_id"],
            "Unrecognised nominal size '{}'.".format(s["recommended_size"]))
        return False
    ok, _approach = _set_pipe_diameter(edge.pipe, nominal_inches)
    return ok


def _lookup_param(element, name):
    """element.LookupParameter(name), or None when it raises (logged)."""
    try:
        return element.LookupParameter(name)
    except Exception as e:
        revit_helpers._log_entry(
            "WARNING", "_lookup_param", revit_helpers.eid_int(element.Id),
            "LookupParameter('{}') raised: {}".format(name, str(e)))
        return None


def _get_pipe_nominal_size(pipe, table_id):
    """The nominal size name from table_id's own size list for a pipe's drawn
    diameter, or None (logged). The table's own list keeps a copper or CSST
    pipe from reading as the steel name sharing its diameter."""
    pipe_id = revit_helpers.eid_int(pipe.Id)
    dia_in = revit_helpers.get_pipe_diameter_inches(pipe)
    if dia_in is None:
        revit_helpers._log_entry(
            "WARNING", "_get_pipe_nominal_size", pipe_id,
            "Diameter NOT READABLE, existing pipe not checked.")
        return None
    nominal = sizing_engine.nominal_for_diameter(table_id, dia_in)
    if nominal is None:
        revit_helpers._log_entry(
            "WARNING", "_get_pipe_nominal_size", pipe_id,
            "Diameter {:.3f} in. matches no size in Table {}.".format(
                float(dia_in), table_id))
    else:
        revit_helpers._log_entry(
            "INFO", "_get_pipe_nominal_size", pipe_id,
            "Diameter {:.3f} in. reads as {} in Table {}.".format(
                float(dia_in), nominal, table_id))
    return nominal


def _audit_existing_pipe(pipe_id, edge, recommended_nom, result,
                         overloaded_list, unchecked_list):
    """Check whether an existing pipe's current size handles the new load.
    Undersized (or no published capacity) goes to overloaded_list; a size
    that matches no table size goes to unchecked_list."""
    demand = edge.cumulative_load_mbh
    if demand <= 0:
        return
    ctx = result.get("edge_context", {}).get(pipe_id, {})
    table_id = ctx.get("table_id", result["table_id"])
    run_ft = ctx.get("longest_run_ft", result["longest_run_ft"])

    current_nom = _get_pipe_nominal_size(edge.pipe, table_id)
    if current_nom is None:
        unchecked_list.append({"pipe_id": pipe_id, "table_id": table_id})
        return

    # Demand in the table's units: CFH via heat content (the propane value
    # on 402.4(37)), or the MBH load itself on a propane table 402.4(25) to
    # (36) (Colin, 2026-11-04, 2026-11-20).
    demand_cfh, units = sizing_engine.demand_in_table_units(
        demand, table_id,
        result.get("heat_content_btu_per_cf",
                   shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF),
        result.get("propane_heat_content_btu_per_cf"))
    try:
        current_cap = float(gas_tables.get_capacity(
            table_id, run_ft, current_nom))
    except ValueError as e:
        revit_helpers._log_entry(
            "WARNING", "_audit_existing_pipe", pipe_id,
            "No published capacity: {}".format(str(e)))
        current_cap = None

    if current_cap is None or current_cap < demand_cfh:
        overloaded_list.append({
            "pipe_id": pipe_id,
            "current_nom": current_nom,
            "demand_mbh": float(demand),
            "current_cap": current_cap,
            "units": units,
            "recommended_nom": recommended_nom,
        })


def _summary_info(result, existing_check, w, st, fit, rolled_back):
    """The info dict gas_console.summary reads. Counts come from the
    transactions as committed: a rolled back step has zero successes and is
    named in rolled_back (n=4/49)."""
    return {
        "success": w["success_count"], "stub_success": st["success"],
        "stub_fail": st["fail"], "existing_skip": w["existing_skip_count"],
        "existing_overloaded": len(w["existing_overloaded"]),
        "existing_unchecked": len(w["existing_unchecked"]),
        "fit": fit, "fail_count": w["fail_count"],
        # no approach stands when Revit rolled the pipe sizes back
        "approach": (_confirmed_approach[0] if w.get("committed")
                     else None),
        "table_id": result["table_id"],
        # gas_console.heat_content_text prints what the table used (n=24)
        "heat_content": result["heat_content_btu_per_cf"],
        "propane_heat_content": result.get(
            "propane_heat_content_btu_per_cf"),
        "existing_check": existing_check,
        "zones": result.get("zones") or [],
        "downstream": bool(result.get("downstream_table_id")),
        "longest_run": result["longest_run_ft"],
        "table_row": result["table_length_used_ft"],
        "rolled_back": list(rolled_back),
        "stub_rolled_back": bool(st.get("rolled_back")),
    }


def _print_summary(con, result, existing_check, w, st, fit, rolled_back):
    """The run summary table (themed console, or its markdown twin)."""
    info = _summary_info(result, existing_check, w, st, fit, rolled_back)
    con.emit(gas_console.summary(info), gas_console.summary_md(info))


# ---------------------------------------------------------------------------
# Headless entry (pyRevit agent runtime: no dialog, no pick, no alert)
# ---------------------------------------------------------------------------

def run_headless(doc, start_id, settings=None, decisions=None, write=True,
                 detail="flagged", include_log=False, output=None):
    """Size Gas with no UI. Returns the headless.result envelope.

    start_id: the gas meter, or the first new gas pipe at the connection to
    existing. settings["existing"] is the dialog's Scope: None (or
    {"new_construction": True}) is New Construction, a pipe start being the
    start of the system; {size_nominal, length_ft, load_mbh} is Connect into
    Existing, which takes the first new pipe only (a meter is refused).
    settings: table_id (required, see table_options()), pipe_material
    (optional, picks the material when a table id is shared),
    heat_content_btu_per_cf (840), propane_heat_content_btu_per_cf (2500,
    used on the propane CFH Table 402.4(37) only), downstream_table_id (None
    = no mid-stream PRV), existing. Heat contents shall be 500 to 3,000
    BTU/CF. decisions: {"presizing_warnings": "continue" | "stop"}, asked
    only when writing. write=False sizes and reports only (run_query);
    write=True writes pipes, fittings and stubs, each step its own
    transaction through revit_helpers.run_in_transaction (warnings deleted,
    Commit status checked, a later step skipped after a rollback;
    run_modify). detail "flagged" lists pipes whose drawn size differs from
    the sized one, existing pipes and zero load pipes; "all" lists every
    pipe.
    """
    import headless
    recorder = headless.RecordingOutput(forward=output)
    state = {"prompter": None}

    def done(status, data=None, errors=None, decision_exc=None):
        log = recorder.log if include_log else None
        prompter = state["prompter"]
        warnings = list(prompter.warnings) if prompter else []
        if decision_exc is not None:
            return headless.needs_decision(TOOL, decision_exc, warnings, log)
        return headless.result(TOOL, status, data=data, warnings=warnings,
                               errors=errors, log=log)

    try:
        state["prompter"] = headless.HeadlessPrompter(decisions)
        settings = merge_settings(settings)
    except ValueError as exc:
        return done("refused", errors=[str(exc)])
    problems = validate_settings(settings)
    if detail not in DETAILS:
        problems.append("detail shall be 'flagged' or 'all'")
    if problems:
        return done("refused", errors=problems)

    recorder.print_md("# Size Gas (headless)")
    try:
        revit_helpers.clear_log()
        element = revit_helpers.get_element_by_id(doc, start_id)
        start = resolve_start(doc, element, settings)
        if start["stop"]:
            return done("refused", errors=[start["stop"]])
        outcome = size_core(doc, settings, element, start["upstream"],
                            start["existing"], recorder, state["prompter"],
                            write=write)
    except headless.NeedsDecision as exc:
        return done("needs_decision", decision_exc=exc)
    except Exception as exc:
        return done("error", errors=[revit_helpers.exc_text(exc)])

    errors = []
    writes = outcome.get("writes") or {}
    if outcome["stop"]:
        errors.append(outcome["stop"])
    for text in writes.get("transaction_errors") or []:
        if not (outcome["stop"] and text.startswith(PIPE_STEP + ":")):
            errors.append(text)
    data = envelope_data(outcome, start_id, start["kind"], settings, write,
                         detail, doc)
    return done(outcome["status"], data=data, errors=errors)


def envelope_data(outcome, start_id, kind, settings, write, detail, doc=None):
    """The JSON data block of the headless result. Pure over the outcome
    (doc only for the phase of each listed pipe)."""
    graph = outcome["graph"]
    result = outcome["result"] or {}
    detail_by_id = dict((d["pipe_id"], d)
                        for d in (result.get("segment_detail") or []))
    pipes = []
    if graph is not None:
        for pipe_id, size in sorted((result.get("sizes") or {}).items()):
            edge = graph.edges.get(pipe_id)
            seg = detail_by_id.get(pipe_id, {})
            drawn = None
            drawn_in = getattr(edge, "diameter_inches", None) if edge else None
            ctx = (result.get("edge_context") or {}).get(pipe_id, {})
            table_id = ctx.get("table_id", result.get("table_id"))
            if drawn_in:
                drawn = sizing_engine.nominal_for_diameter(table_id, drawn_in)
            is_new = None
            if doc is not None and edge is not None and edge.pipe is not None:
                try:
                    is_new = bool(revit_helpers.is_new_construction(
                        edge.pipe, doc))
                except Exception:
                    is_new = None
            to_node = graph.nodes.get(edge.to_node_id) if edge else None
            changed = drawn != size
            mbh = float(seg.get("demand_mbh", getattr(
                edge, "cumulative_load_mbh", 0.0) or 0.0))
            listed = (detail == "all" or changed or is_new is False
                      or mbh <= 0)
            if not listed:
                continue
            pipes.append({
                "id": pipe_id, "size": size, "drawn": drawn,
                "drawn_in": drawn_in, "changed": changed,
                "mbh": mbh, "demand": seg.get("demand_cfh"),
                "units": seg.get("units"), "capacity": seg.get("capacity_mbh"),
                "table_id": table_id,
                "length_ft": getattr(edge, "length_feet", None),
                "is_new": is_new,
                "stub_to": (to_node.fixture_name or "UNNAMED")
                if (to_node is not None and to_node.is_gas_fixture) else None,
                "note": seg.get("note"),
            })

    fixtures = []
    if graph is not None:
        for n in graph.nodes.values():
            if n.is_gas_fixture:
                fixtures.append({"id": n.element_id, "name": n.fixture_name,
                                 "mbh": float(n.gas_load_mbh or 0.0)})
    lr = (graph.longest_run if graph is not None else None) or {}
    data = {
        "origin": {"id": int(start_id), "start": kind},
        "write": bool(write),
        "table_id": settings.get("table_id"),
        "downstream_table_id": settings.get("downstream_table_id"),
        "heat_content_btu_per_cf": float(settings["heat_content_btu_per_cf"]),
        "propane_heat_content_btu_per_cf": float(
            settings.get("propane_heat_content_btu_per_cf")
            or shared_params.DEFAULT_PROPANE_HEAT_CONTENT_BTU_PER_CF),
        "existing": settings.get("existing"),
        "nodes": len(graph.nodes) if graph is not None else 0,
        "pipe_segments": len(graph.edges) if graph is not None else 0,
        "fixtures": sorted(fixtures, key=lambda f: f["id"]),
        "total_load_mbh": float(sum([f["mbh"] for f in fixtures], 0.0)),
        "longest_run_ft": result.get("longest_run_ft"),
        "farthest_fixture": lr.get("farthest_fixture_name"),
        "table_row_ft": result.get("table_length_used_ft"),
        "zones": result.get("zones"),
        "existing_check": outcome["existing_check"],
        "presizing_warnings": outcome["warnings"],
        "pipes_listed": detail,
        "pipes": pipes,
        "writes": outcome["writes"],
    }
    if graph is not None:
        data["existing_connections"] = list(graph.existing_connections or [])
        data["existing_appliances"] = list(
            getattr(graph, "existing_appliances", None) or [])
    return data
