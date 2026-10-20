# -*- coding: ascii -*-
# water_performance.py
# Domestic water pipe performance: classify each drawn pipe against the firm
# sizing table, and report its computed speed.
#
# Pure logic, no Revit, so it can be unit tested in CPython.
#
# Basis (Colin, 2026-10-15): the TRUTH is the drawn size against the firm
# table size for the pipe's demand (water_sizing_engine.size_network, run read
# only). Computed speed is shown beside it and only raises a pipe that is at
# table size to YELLOW, because the firm table is not exactly reproducible
# from an inside diameter.
#
# Colors:
#   GREEN   drawn size equals the firm table size, speed within the limit
#   YELLOW  drawn size equals the firm table size, computed speed above limit
#   RED     drawn size is smaller than the firm table size
#   PURPLE  drawn size is larger than the firm table size
#   GRAY    could not be judged (over the table, no demand, no size, no ID)
#
# IronPython 2.7

import water_tables

GREEN = "GREEN"
YELLOW = "YELLOW"
RED = "RED"
PURPLE = "PURPLE"
GRAY = "GRAY"

# Worst first. Used to color fittings from the pipes next to them.
PRIORITY = {RED: 4, YELLOW: 3, PURPLE: 2, GREEN: 1, GRAY: 0}

# Copper Type L actual inside diameters, inches, keyed by nominal size.
# Source: ASTM B88 Type L, ID = OD - 2 x wall. OD 0.625 to 3.125 in, walls
# 0.040, 0.045, 0.050, 0.055, 0.060, 0.070, 0.080, 0.090 in (Engineering
# ToolBox reproduction of ASTM B88, engineeringtoolbox.com
# astm-copper-tubes-d_779.html). Verify against the current standard.
COPPER_L_ID_IN = {
    0.5:  0.545,
    0.75: 0.785,
    1.0:  1.025,
    1.25: 1.265,
    1.5:  1.505,
    2.0:  1.985,
    2.5:  2.465,
    3.0:  2.945,
}

# Q (gpm) = 2.4483 x V (fps) x d^2 (in^2), the same relation the sizing
# method note uses to check the firm table.
GPM_PER_FPS_IN2 = 2.4483

SIZE_TOLERANCE_IN = 0.001

# Yellow only when computed speed is this far over the limit. At the firm
# table's own limits the computed speed lands from 4.1 to 5.04 fps hot and
# 6.9 to 8.02 fps cold (checked 2026-10-15), so a bare "over the limit" test
# would flag 2-1/2 and 3 in pipes sized exactly per the table. Same 10 percent
# default as the duct mains' yellow band.
SPEED_TOLERANCE = 0.10


def inside_diameter_in(nominal_in):
    """Type L inside diameter for a nominal size, or None if not tabulated."""
    if nominal_in is None:
        return None
    for key, value in COPPER_L_ID_IN.items():
        if abs(key - float(nominal_in)) < SIZE_TOLERANCE_IN:
            return value
    return None


def velocity_fps(gpm, nominal_in):
    """Speed in fps for a flow in a nominal size, or None when unknown."""
    if gpm is None:
        return None
    diameter = inside_diameter_in(nominal_in)
    if diameter is None:
        return None
    return float(gpm) / (GPM_PER_FPS_IN2 * diameter * diameter)


class PipeResult(object):
    """The verdict for one pipe."""

    def __init__(self, element_id):
        self.element_id = element_id
        self.label = GRAY
        self.reason = ""
        self.system = None
        self.is_return = False
        self.drawn_in = None
        self.required_in = None
        self.required_nominal = None
        self.gpm = None
        self.fps = None
        self.limit_fps = None
        self.length_feet = 0.0


def classify(segment, system_key=None, speed_tolerance=SPEED_TOLERANCE):
    """Judge one water_sizing_engine.SegmentResult.

    Args:
        segment: SegmentResult from size_network.
        system_key: 'cold' or 'hot'. For return pipes pass 'hot'.
        speed_tolerance: fraction over the speed limit before a pipe at table
            size turns yellow (dialog input, default SPEED_TOLERANCE).

    Returns:
        PipeResult.
    """
    result = PipeResult(segment.element_id)
    result.system = system_key
    result.is_return = bool(segment.is_return)
    result.drawn_in = segment.existing_size_in
    result.required_in = segment.size_inches
    result.required_nominal = segment.nominal_size
    result.length_feet = segment.length_feet
    result.gpm = segment.circ_gpm if segment.is_return else segment.demand_gpm
    if system_key is not None:
        result.limit_fps = water_tables.velocity_limit_fps(system_key)

    if segment.exceeds_table:
        result.reason = "load above the firm table, not judged"
        return result
    if segment.size_inches is None:
        result.reason = segment.flag or "no table size for this pipe"
        return result
    if not segment.existing_size_in:
        result.reason = "drawn size not read"
        return result
    if (not segment.is_return and
            float(segment.demand_wsfu or 0.0) <= 0.0):
        result.reason = "no fixture demand downstream"
        return result

    drawn = float(segment.existing_size_in)
    required = float(segment.size_inches)
    # Speed at the DRAWN size, which is what the pipe actually does.
    result.fps = velocity_fps(result.gpm, drawn)

    if drawn < required - SIZE_TOLERANCE_IN:
        result.label = RED
        result.reason = "drawn {} in, table requires {}".format(
            _fmt(drawn), segment.nominal_size)
    elif drawn > required + SIZE_TOLERANCE_IN:
        result.label = PURPLE
        result.reason = "drawn {} in, table size is {}".format(
            _fmt(drawn), segment.nominal_size)
    else:
        result.label = GREEN
        result.reason = "at table size"
        if (result.fps is not None and result.limit_fps is not None and
                result.fps > result.limit_fps * (1.0 + speed_tolerance)):
            result.label = YELLOW
            result.reason = ("at table size, computed {:.1f} fps is more than "
                             "{:.0f} percent over the {:.0f} fps limit".format(
                                 result.fps, float(speed_tolerance) * 100.0,
                                 result.limit_fps))
    return result


def summarize(results):
    """Counts by label."""
    counts = {GREEN: 0, YELLOW: 0, RED: 0, PURPLE: 0, GRAY: 0}
    for r in results:
        counts[r.label] = counts.get(r.label, 0) + 1
    return counts


def worst_label(labels):
    """Highest priority label of an iterable, or None when empty."""
    worst = None
    for label in labels:
        if worst is None or PRIORITY.get(label, 0) > PRIORITY.get(worst, 0):
            worst = label
    return worst


def basis_lines(speed_tolerance=SPEED_TOLERANCE):
    """Basis of design block for the performance report."""
    cold = water_tables.velocity_limit_fps("cold")
    hot = water_tables.velocity_limit_fps("hot")
    lines = list(water_tables.basis_of_design_lines())
    lines.append("")
    lines.append("Performance check basis:")
    lines.append("  Pass/fail: drawn nominal size against the firm WSFU table "
                 "size for the pipe's demand (sizing run read only, minimums "
                 "applied). Return pipes against circulation flow.")
    lines.append("  Speed shown: gpm / (2.4483 x ID^2), ID from ASTM B88 "
                 "Copper Type L (OD - 2 x wall), at the DRAWN size.")
    lines.append("  Speed limits: {:.0f} fps cold, {:.0f} fps hot and hot "
                 "water return (firm table). Yellow above {:.0f} percent "
                 "over.".format(cold, hot, float(speed_tolerance) * 100.0))
    lines.append("  Flow: 2024 IPC Table E103.3(3), flush tank column. "
                 "Flushometer heavy systems read low.")
    return lines


def _fmt(value):
    return "{:.3f}".format(float(value)).rstrip("0").rstrip(".")
