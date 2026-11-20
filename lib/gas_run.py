# -*- coding: ascii -*-
# gas_run.py
# The Size Gas pipeline (IFGC Longest Run Method), shared by the Size Gas
# button and the headless entry (pyRevit agent runtime).
#
#   dialog -> pick meter or first new pipe -> (pipe: existing conditions)
#   -> traverse -> PRV zones -> size -> report -> pre sizing warnings (yes/no)
#   -> write pipes, fittings, fixture stubs (three transactions)
#
# The button path is run(); the agent path is run_headless(). Both go through
# resolve_start() and size_core(), so there is one code path after the pick.
# The step bodies moved here verbatim from Size Gas.pushbutton/script.py; only
# `output`, the yes/no question and the transaction factory are passed in.
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
    """Size Gas behind the button: dialog, pick, existing conditions, core.
    Returns True if it finished. forms is kept for the button's signature;
    every popup is the RJA themed ui_helpers.rja_alert."""
    import ui_helpers
    from Autodesk.Revit.DB import Transaction

    con = _Console(output, themed=True)
    revit_helpers.clear_log()

    choice = ui_helpers.show_size_gas_dialog("Size Gas - Select IFGC Table")
    if choice is None:
        con.say("Cancelled at table selection. No changes were made to the "
                "model.")
        return False
    settings = settings_from_dialog(choice)

    # the shared pick step, same words and window as One-Line (2026-11-19)
    selected_element = ui_helpers.pick_element(
        uidoc, doc, "Size Gas", ui_helpers.PICK_GAS)
    if selected_element is None:
        con.say("Selection cancelled. No changes were made to the model.")
        return False

    def ask_existing(picked_id, table_id):
        existing = ui_helpers.show_existing_gas_dialog(
            "Size Gas - Existing Conditions",
            gas_tables.list_pipe_sizes(table_id),
            tie_in.recall_existing(picked_id))
        if existing is not None and not tie_in.is_new_construction(existing):
            tie_in.remember_existing(picked_id, existing)
        return existing

    start = resolve_start(doc, selected_element, settings, output,
                          ask_existing=ask_existing, themed=True)
    if start["stop"]:
        if start["stop"] != _CANCELLED:
            ui_helpers.rja_alert(TITLE, start["stop"], "error")
        return False

    inputs = {"selected": "Element ID {}".format(
        revit_helpers.eid_int(selected_element.Id)),
        "check": start["check"]}
    if start["existing_text"]:
        inputs["existing"] = start["existing_text"]

    outcome = size_core(doc, settings, selected_element, start["upstream"],
                        start["existing"], output, _RjaPrompter(),
                        write=True,
                        transaction_factory=lambda d, n: Transaction(d, n),
                        themed=True, inputs_extra=inputs)
    if outcome["stop"]:
        ui_helpers.rja_alert(TITLE, outcome["stop"], "error")
    return outcome["status"] == "ok"


def settings_from_dialog(choice):
    """The dialog's answer as headless settings (same keys as
    default_settings)."""
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
        "downstream_table_id": down,
        "existing": None,
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
    asked at the start of every gas job, never assumed."""
    return {
        "table_id": None,
        "pipe_material": None,
        "heat_content_btu_per_cf":
            float(shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF),
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
    dialog: a known table, heat content above 0, a downstream table of the
    same material and gas that is not a higher pressure."""
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
    try:
        if float(settings["heat_content_btu_per_cf"]) <= 0:
            problems.append("heat_content_btu_per_cf shall be above 0")
    except (TypeError, ValueError):
        problems.append("heat_content_btu_per_cf is not a number")

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

_CANCELLED = "cancelled"


def resolve_start(doc, element, settings, output, ask_existing=None,
                  themed=False):
    """Check the picked meter or first new pipe and find where the walk
    begins.

    For a pipe, the existing conditions come from settings["existing"], else
    ask_existing(picked_id, table_id) (the button's dialog). With neither
    (headless), the run is full new construction: {"new_construction": True},
    the picked pipe starts the system with no existing piping, as the button
    with Connect into Existing left off (Colin, 2026-11-20).
    Returns dict: kind ("meter" | "pipe"), upstream, existing, check (one line
    for the inputs table), existing_text (inputs line, or None), stop (None
    to go on, _CANCELLED, or the message that stopped the run), title.
    Nothing is printed here but the cancel line; the inputs section prints in
    size_core once the pick is known.
    """
    out = {"kind": START_METER, "upstream": None, "existing": None,
           "check": None, "existing_text": None,
           "stop": None, "title": "Size Gas - Invalid Selection"}
    if element is None:
        out["stop"] = "Could not retrieve selected element."
        out["title"] = "Size Gas - Selection Error"
        return out

    if revit_helpers.is_pipe(element):
        out["kind"] = START_PIPE
        gas_check = pipe_graph.check_gas_pipe(element)
        if not gas_check["ok"]:
            out["stop"] = (
                "This is not a gas pipe: {}.\n\nPick the gas meter, or the "
                "first new GAS pipe at the connection to existing.".format(
                    gas_check["reason"]))
            return out
        upstream = tie_in.resolve_new_pipe_upstream(element, doc)
        if upstream["status"] != "ok":
            out["stop"] = (
                "Could not tell which end of the picked pipe connects to the "
                "existing gas piping: {}.\n\nPick the first new gas pipe, the "
                "one with one end at the existing connection (open, capped, "
                "or joined to existing phase piping) and the other end on "
                "new construction.".format(upstream["reason"]))
            return out
        out["upstream"] = upstream
        out["check"] = ("New gas piping connection. Upstream end found "
                        "because {}.".format(upstream["reason"]))

        existing = settings.get("existing")
        if existing is None and ask_existing is not None:
            existing = ask_existing(revit_helpers.eid_int(element.Id),
                                    settings["table_id"])
            if existing is None:
                _Console(output, themed).say(
                    "Cancelled at existing conditions. No changes were made "
                    "to the model.")
                out["stop"] = _CANCELLED
                return out
        if existing is None:
            # Headless with no existing given: full new construction, the
            # same as the button with Connect into Existing left off (Colin,
            # 2026-11-20). Give settings["existing"] = {size_nominal,
            # length_ft, load_mbh} to connect into existing piping instead.
            existing = {"new_construction": True}
        out["existing"] = existing
        out["existing_text"] = tie_in.existing_inputs_text(existing)
        return out

    validation = revit_helpers.validate_selected_element(element)
    if not validation["is_valid"]:
        out["stop"] = (
            "Please select the gas meter element, or the first new gas pipe "
            "at the connection to existing.\n\n{}".format(
                validation["reason"]))
        return out
    out["check"] = "Meter validation passed."
    return out


# ---------------------------------------------------------------------------
# Core (one code path after the pick)
# ---------------------------------------------------------------------------

def size_core(doc, settings, element, upstream, existing, output, prompter,
              write=True, transaction_factory=None, themed=False,
              inputs_extra=None):
    """Traverse, zone, size, report and (write=True) write the sizes.

    The yes/no pre sizing warnings question goes through prompter.confirm
    (qid "presizing_warnings") and is asked only when writing; with
    HeadlessPrompter it raises NeedsDecision before anything is written.
    themed=True (the button) prints the RJA themed console sections through
    gas_console; themed=False prints their markdown twins. inputs_extra adds
    the pick lines (selected, check, existing) to the inputs section.

    Returns dict: status ("ok" | "cancelled" | "refused"), stop (message
    that stopped the run, or None), stop_title, graph, result, existing_check,
    warnings (dict of lists), writes (None when nothing was written).
    """
    outcome = {"status": "ok", "stop": None, "stop_title": "Size Gas",
               "graph": None, "result": None, "existing_check": None,
               "warnings": {}, "writes": None}
    con = _Console(output, themed)

    def stop(text, title):
        outcome["status"] = "refused"
        outcome["stop"] = text
        outcome["stop_title"] = title
        return outcome

    revit_helpers.reset_pipe_diameter_approach()

    table_id = settings["table_id"]
    selected_opt = _option(table_id, settings.get("pipe_material"))
    pipe_material = selected_opt["material"]
    inlet_pressure_psi = selected_opt["inlet_pressure_psi"]
    heat_content_btu_per_cf = float(settings["heat_content_btu_per_cf"])
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
        "heat_content": heat_content_btu_per_cf,
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
        return stop("Traversal failed:\n\n{}".format(str(e)),
                    "Size Gas - Traversal Error")
    outcome["graph"] = graph

    trav = {"nodes": len(graph.nodes), "segments": len(graph.edges),
            "existing_connections": list(graph.existing_connections or [])}
    con.emit(gas_console.traversal(trav), gas_console.traversal_md(trav))

    fixture_nodes = [n for n in graph.nodes.values() if n.is_gas_fixture]
    if len(fixture_nodes) == 0:
        return stop(
            "No gas fixtures found (IS_GAS_FIXTURE = Yes).\n\n"
            "Common causes:\n"
            "  1. Pipe caps were inadvertently changed when resizing all "
            "piping in Revit -- verify cap families have not replaced fixture "
            "families at pipe terminations.\n"
            "  2. Fixture families are not physically connected to the piping.\n"
            "  3. IS_GAS_FIXTURE parameter is not set to Yes on the appliances.\n\n"
            "Run Diagnose to see which terminal elements were found.",
            "Size Gas - No Fixtures")

    if graph.longest_run is None:
        return stop("Could not determine longest run. "
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

    if mid_stream_prv and not midstream_prvs:
        if prv_nodes:
            why = ("{} regulator(s) were found, but each has {:g} ft of pipe "
                   "or less downstream, so they are treated as equipment "
                   "PRVs and ignored.".format(len(prv_nodes), min_ft))
        else:
            why = ("No pressure regulating valve was found. A PRV is "
                   "recognised by 'PRV', 'regulator' or 'regulating' in its "
                   "family name; check it is connected to the piping.")
        return stop(
            "Mid Stream PRV is checked, but there is no mid-stream "
            "regulator in this system.\n\n{}\n\nUncheck Mid Stream PRV "
            "to size the system as one.".format(why),
            "Size Gas - No Mid-Stream PRV")

    if mid_stream_prv and graph.nested_prv_ids:
        return stop(
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
            downstream_table_id=downstream_table_id)
    except ValueError as e:
        con.say("Sizing ERROR: {}".format(str(e)), "error")
        return stop("Sizing failed:\n\n{}".format(str(e)),
                    "Size Gas - Sizing Error")
    outcome["result"] = result

    # Sizing results: the basis, then one table per system.
    total_load = sum(n.gas_load_mbh for n in fixture_nodes)
    units = gas_tables.table_units(result["table_id"])
    con.emit(gas_console.sizing_basis(result, total_load, len(fixture_nodes),
                                      units),
             gas_console.sizing_basis_md(result, total_load,
                                         len(fixture_nodes), units))
    rows = sizing_engine.sizing_rows(result, graph)
    con.emit(gas_console.sizing_tables(rows, result),
             gas_console.sizing_tables_md(rows, result))

    existing_check = sizing_engine.check_existing_pipe(result, graph)
    outcome["existing_check"] = existing_check
    existing_warnings = []
    if existing_check is not None:
        con.emit(gas_console.existing_check(existing_check),
                 gas_console.existing_check_md(existing_check))
        if not existing_check["adequate"]:
            existing_warnings.append(_existing_warning(existing_check))

    # STEP 6 - Pre-sizing validation warning
    fixture_nodes = [n for n in graph.nodes.values() if n.is_gas_fixture]
    no_load = [n for n in fixture_nodes if n.gas_load_mbh <= 0]
    unnamed = [n for n in fixture_nodes
               if n.fixture_name in ("", "UNNAMED", None)]
    total_mbh = sum(n.gas_load_mbh for n in fixture_nodes)
    table_warnings = sizing_engine.table_limit_warnings(
        table_id,
        sizing_engine.mbh_to_cfh(total_mbh, heat_content_btu_per_cf),
        total_mbh)
    non_gas = list(graph.non_gas_skipped)
    bare_end = pipe_graph.non_fixture_ends(graph)

    outcome["warnings"] = {
        "existing_pipe": existing_warnings,
        "prv": prv_warnings,
        "table_limit": table_warnings,
        "no_load": [{"id": n.element_id, "name": n.fixture_name or "UNNAMED"}
                    for n in no_load],
        "unnamed": [{"id": n.element_id, "mbh": float(n.gas_load_mbh)}
                    for n in unnamed],
        "non_gas": [{"id": i, "classification": c} for i, c in non_gas],
        "bare_end": [{"id": n.element_id, "family": n.family_name}
                     for n in bare_end],
        "disconnected": list(graph.disconnected),
    }

    groups = gas_console.build_warning_groups(
        existing=existing_warnings, prv=prv_warnings,
        table_limit=table_warnings,
        no_load=[(n.fixture_name, n.element_id) for n in no_load],
        unnamed=[(n.element_id, n.gas_load_mbh) for n in unnamed],
        non_gas=list(non_gas),
        bare_end=[(n.element_id, n.family_name) for n in bare_end],
        disconnected=list(graph.disconnected),
        n_fixtures=len(fixture_nodes))
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

    outcome["writes"] = _write_sizes(doc, output, graph, result,
                                     existing_check, transaction_factory,
                                     themed=themed,
                                     new_construction=new_construction)
    if outcome["writes"].get("pipe_transaction_error"):
        return stop("Transaction failed - no sizes were written to the "
                    "model.\n\n{}".format(
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
# Write (three transactions: pipes, fittings, fixture stubs)
# ---------------------------------------------------------------------------

def _write_sizes(doc, output, graph, result, existing_check,
                 transaction_factory, themed=False, new_construction=False):
    """STEPS 7 to 9 of the button. Returns the counts."""
    con = _Console(output, themed)
    state = {"approach": None}
    approaches = ("RBS_PIPE_NOMINAL_DIAMETER, RBS_PIPE_DIAMETER_PARAM, "
                  "LookupParameter(Diameter)")
    con.emit(gas_console.write_start(approaches),
             gas_console.write_start_md(approaches))
    success_count = 0
    fail_count = 0
    fail_rows = []     # pipe failures, {"id", "family", "reason"}
    skipped_stubs = []
    existing_skip_count = 0
    existing_overloaded = []
    existing_unchecked = []
    new_sizes = {}
    written = []

    writes = {"pipes_sized": 0, "pipe_failures": [], "stubs_sized": 0,
              "stub_failures": 0, "existing_skipped": 0,
              "existing_overloaded": [], "existing_unchecked": [],
              "fittings": None, "approach": None, "written_ids": written,
              "pipe_transaction_error": None}

    # Transitions are identified now, before the resize: Revit may delete one
    # while the pipes change size (Colin, 2026-11-19).
    transition_ids = _transition_ids(graph)

    t = transaction_factory(doc, "RJA Tools - Size Gas Pipes")
    t.Start()
    try:
        for pipe_id, nominal_size in result["sizes"].items():
            edge = graph.edges.get(pipe_id)
            if edge is None or edge.pipe is None:
                fail_rows.append({
                    "id": pipe_id, "family": "Pipe",
                    "reason": "edge or pipe element not found in graph."})
                fail_count += 1
                continue

            is_new = revit_helpers.is_new_construction(edge.pipe, doc)
            if is_new:
                new_sizes[pipe_id] = nominal_size

            to_node = graph.nodes.get(edge.to_node_id)
            is_stub = (to_node is not None and to_node.is_gas_fixture)

            if is_stub:
                if is_new:
                    skipped_stubs.append({
                        "pipe_id": pipe_id,
                        "fixture_name": to_node.fixture_name or "UNNAMED",
                        "demand_mbh": edge.cumulative_load_mbh,
                        "recommended_size": nominal_size,
                    })
                else:
                    existing_skip_count += 1
                    _audit_existing_pipe(
                        pipe_id, edge, nominal_size, result,
                        existing_overloaded, existing_unchecked)
                continue

            if not is_new:
                existing_skip_count += 1
                _audit_existing_pipe(
                    pipe_id, edge, nominal_size, result,
                    existing_overloaded, existing_unchecked)
                continue

            nominal_inches = sizing_engine.NOMINAL_TO_INCHES.get(nominal_size)
            if nominal_inches is None:
                fail_rows.append({
                    "id": pipe_id, "family": "Pipe",
                    "reason": "unrecognised nominal size '{}'.".format(
                        nominal_size)})
                fail_count += 1
                continue

            ok, approach = _set_pipe_diameter(edge.pipe, nominal_inches,
                                              state)
            if ok:
                success_count += 1
                written.append(pipe_id)
            else:
                fail_count += 1
                fail_rows.append({
                    "id": pipe_id, "family": "Pipe",
                    "reason": "all API approaches failed ({}\").".format(
                        nominal_size)})

        doc.Regenerate()
        t.Commit()
    except Exception as e:
        try:
            t.RollBack()
        except Exception:
            pass
        con.say("Transaction ERROR: {}".format(str(e)), "error")
        writes["pipe_transaction_error"] = str(e)
        return writes

    # The existing pipe audit can never stop the fitting and stub
    # transactions (Colin, 2026-11-02).
    try:
        # new construction: print only if the walk still met existing phase
        # pipes, so a skip is never silent
        if not new_construction or existing_skip_count:
            con.emit(gas_console.existing_audit(existing_skip_count,
                                                existing_overloaded,
                                                existing_unchecked),
                     gas_console.existing_audit_md(existing_skip_count,
                                                   existing_overloaded,
                                                   existing_unchecked))
    except Exception as e:
        con.say("Existing pipe audit could not be printed: {}".format(str(e)),
                "warn")

    # STEP 8 - Resize fittings. Transitions are counted and left to Revit.
    fit = _new_fit_counts()
    t2 = transaction_factory(doc, "RJA Tools - Resize Pipe Fittings")
    t2.Start()
    try:
        fit = _resize_fittings(graph, new_sizes, doc, transition_ids)
        doc.Regenerate()
        t2.Commit()
    except Exception as e:
        try:
            t2.RollBack()
        except Exception:
            pass
        fit = _new_fit_counts()
        fit["failures"].append(
            "Fitting transaction rolled back: {}".format(str(e)))
        fit["failure_rows"].append({
            "id": "-", "family": "All fittings",
            "reason": "transaction rolled back: {}".format(str(e))})

    # STEP 9 - Resize fixture stub pipes (after the caps)
    stub_success = 0
    stub_fail = 0
    stub_rows = []
    if skipped_stubs:
        t3 = transaction_factory(doc, "RJA Tools - Size Fixture Stub Pipes")
        t3.Start()
        try:
            for s in skipped_stubs:
                try:
                    if _size_one_stub(s, graph, state):
                        stub_success += 1
                        written.append(s["pipe_id"])
                    else:
                        stub_fail += 1
                        stub_rows.append({
                            "id": s["pipe_id"], "family": "Fixture stub pipe",
                            "reason": "size {}\" not written (see log)".format(
                                s["recommended_size"])})
                except Exception as e:
                    stub_fail += 1
                    revit_helpers._log_entry(
                        "ERROR", "size_gas_stubs", s["pipe_id"],
                        "Stub resize raised: {}".format(str(e)))
                    stub_rows.append({
                        "id": s["pipe_id"], "family": "Fixture stub pipe",
                        "reason": str(e)})
            doc.Regenerate()
            t3.Commit()
        except Exception as e:
            try:
                t3.RollBack()
            except Exception:
                pass
            stub_success = 0
            stub_rows.append({
                "id": "-", "family": "Fixture stub pipes",
                "reason": "transaction rolled back: {}".format(str(e))})

    try:
        failures = gas_console.failure_rows(
            fail_rows, fit.get("failure_rows") or [], stub_rows)
        wr = {"pipes_sized": success_count, "stubs_sized": stub_success,
              "fit": fit, "failures": failures}
        con.emit(gas_console.write_results(wr),
                 gas_console.write_results_md(wr))
    except Exception as e:
        con.say("Write results could not be printed: {}. The sizes were "
                "written.".format(str(e)), "warn")

    try:
        _print_summary(con, result, existing_check, success_count,
                       stub_success, stub_fail, existing_skip_count,
                       existing_overloaded, existing_unchecked, fit,
                       fail_count, state["approach"])
    except Exception as e:
        con.say("Summary could not be printed: {}. The sizes above were "
                "written.".format(str(e)), "warn")

    writes.update({
        "pipes_sized": success_count,
        "pipe_failures": ["Pipe {}: {}".format(r["id"], r["reason"])
                          for r in fail_rows],
        "stubs_sized": stub_success, "stub_failures": stub_fail,
        "existing_skipped": existing_skip_count,
        "existing_overloaded": existing_overloaded,
        "existing_unchecked": existing_unchecked,
        "fittings": fit, "approach": state["approach"],
    })
    return writes


def _new_fit_counts():
    """The counts _resize_fittings fills. failures holds the text lines,
    failure_rows the same failures as data for the console table."""
    return {"resized": 0, "no_sized_pipe": 0, "existing": 0,
            "no_element": 0, "appliance": 0, "replaced": 0,
            "transitions": 0, "transitions_removed": 0,
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


def _set_pipe_diameter(pipe, nominal_inches, state):
    """revit_helpers.set_pipe_diameter plus the approach that worked (shown
    in the Summary). Nothing prints per pipe."""
    ok, name = revit_helpers.set_pipe_diameter(pipe, nominal_inches)
    if ok and state["approach"] is None:
        state["approach"] = name
    return ok, name


def _resize_fittings(graph, new_sizes, doc, transition_ids=None):
    """Resize the new construction fittings to the pipes sized this run.

    For each fitting the target size is the highest-load connected pipe in
    new_sizes (new construction pipes only); a valve with a transition on
    each side reads the pipes past them. An existing phase fitting is never
    resized (Colin, 2026-11-02). Transitions are never resized here (Colin,
    2026-11-19): a transition takes its two sizes from its pipes, and Revit
    updates or removes it when the pipes change. transition_ids is from
    _transition_ids(), taken before the pipe resize. Only 'Nominal Radius'
    is written; Connector.Radius writes disconnect fittings (verified
    2026-10-12). One guard per fitting (G7C, 2026-11-04).
    """
    counts = _new_fit_counts()
    transition_ids = transition_ids or set()
    done_ids = set()
    for node in graph.nodes.values():
        if node.node_type not in ("tee", "fitting", "elbow", "fixture"):
            continue
        try:
            _resize_one_fitting(node, graph, new_sizes, doc, counts, done_ids,
                                transition_ids)
        except Exception as e:
            _fit_fail(counts, node, node.family_name, str(e))
            revit_helpers._log_entry(
                "ERROR", "_resize_fittings", node.element_id,
                "Resize raised: {}".format(str(e)))
    return counts


def _resize_one_fitting(node, graph, new_sizes, doc, counts, done_ids,
                        transition_ids=()):
    """Resize one fitting node. A fitting Revit deleted when its pipes were
    resized is found again on the connected pipes and sized instead."""
    if node.element is None:
        counts["no_element"] += 1
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

    # The highest-load new pipe that touches it sets the size; a valve with
    # a transition on each side reads the pipes past them (Horse Creek).
    target_size, via_transition = sizing_engine.fitting_target_size(
        node.element_id, graph, new_sizes, transition_ids)
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

    ok, approach = revit_helpers.set_fitting_size(element, nominal_inches)
    if ok:
        counts["resized"] += 1
        if replaced:
            counts["replaced"] += 1
            done_ids.add(revit_helpers.eid_int(element.Id))
    else:
        _fit_fail(counts, node, node.family_name,
                  "could not set size to {}\" ({}).".format(
                      target_size, approach))


def _size_one_stub(s, graph, state):
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
    ok, _approach = _set_pipe_diameter(edge.pipe, nominal_inches,
                                       state)
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

    demand_cfh, units = sizing_engine.demand_in_table_units(
        demand, table_id, result.get(
            "heat_content_btu_per_cf",
            shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF))
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


def _print_summary(con, result, existing_check, success_count,
                   stub_success, stub_fail, existing_skip_count,
                   existing_overloaded, existing_unchecked, fit, fail_count,
                   approach):
    """The run summary table (themed console, or its markdown twin)."""
    info = {
        "success": success_count, "stub_success": stub_success,
        "stub_fail": stub_fail, "existing_skip": existing_skip_count,
        "existing_overloaded": len(existing_overloaded),
        "existing_unchecked": len(existing_unchecked),
        "fit": fit, "fail_count": fail_count,
        "approach": approach,
        "table_id": result["table_id"],
        "heat_content": result["heat_content_btu_per_cf"],
        "existing_check": existing_check,
        "zones": result.get("zones") or [],
        "downstream": bool(result.get("downstream_table_id")),
        "longest_run": result["longest_run_ft"],
        "table_row": result["table_length_used_ft"],
    }
    con.emit(gas_console.summary(info), gas_console.summary_md(info))


# ---------------------------------------------------------------------------
# Headless entry (pyRevit agent runtime: no dialog, no pick, no alert)
# ---------------------------------------------------------------------------

def run_headless(doc, start_id, settings=None, decisions=None, write=True,
                 detail="flagged", include_log=False, output=None):
    """Size Gas with no UI. Returns the headless.result envelope.

    start_id: the gas meter, or the first new gas pipe at the connection to
    existing. A pipe start with no settings["existing"] is full new
    construction, the picked pipe starting the system; to connect into
    existing piping give {size_nominal, length_ft, load_mbh}). settings: table_id
    (required, see table_options()), pipe_material (optional, picks the
    material when a table id is shared), heat_content_btu_per_cf (840),
    downstream_table_id (None = no mid-stream PRV), existing.
    decisions: {"presizing_warnings": "continue" | "stop"}, asked only when
    writing. write=False sizes and reports only (run_query); write=True writes
    pipes, fittings and stubs, each transaction with the warning deleting
    failure handler (run_modify). detail "flagged" lists pipes whose drawn
    size differs from the sized one, existing pipes and zero load pipes;
    "all" lists every pipe.
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
        start = resolve_start(doc, element, settings, recorder)
        if start["stop"]:
            return done("refused", errors=[start["stop"]])
        if write:
            factory, check_tx = revit_helpers.quiet_transaction_factory()
        else:
            factory, check_tx = None, (lambda: [])
        outcome = size_core(doc, settings, element, start["upstream"],
                            start["existing"], recorder, state["prompter"],
                            write=write, transaction_factory=factory)
    except headless.NeedsDecision as exc:
        return done("needs_decision", decision_exc=exc)
    except Exception as exc:
        return done("error", errors=[revit_helpers.exc_text(exc)])

    errors = check_tx()
    if outcome["stop"]:
        errors.insert(0, outcome["stop"])
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
                                 "mbh": float(n.gas_load_mbh)})
    lr = (graph.longest_run if graph is not None else None) or {}
    data = {
        "origin": {"id": int(start_id), "start": kind},
        "write": bool(write),
        "table_id": settings.get("table_id"),
        "downstream_table_id": settings.get("downstream_table_id"),
        "heat_content_btu_per_cf": float(settings["heat_content_btu_per_cf"]),
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
    return data
