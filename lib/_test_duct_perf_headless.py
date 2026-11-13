# -*- coding: ascii -*-
"""Tests for the Duct Performance headless entry (duct_perf_run): dialog
controls to the dialog answer, refusals and the needs_decision for a ceiling
exhaust fan.

Plain CPython, Revit stubbed. The judging and pricing are covered by
_test_critical_path, _test_fan_system_effect, _test_duct_perf_sheet and
_test_level_plans; the model side needs a live Revit run.

Run:  python _test_duct_perf_headless.py
"""

import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


class _StubMeta(type):
    _next = [1000]

    def __getattr__(cls, name):
        if name.startswith("__"):
            raise AttributeError(name)
        _StubMeta._next[0] += 1
        setattr(cls, name, _StubMeta._next[0])
        return _StubMeta._next[0]


_Stub = _StubMeta("_Stub", (object,), {
    "__init__": lambda self, *a, **k: setattr(self, "args", a)})


class _StubModule(types.ModuleType):
    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)
        cls = _StubMeta(name, (_Stub,), {})
        setattr(self, name, cls)
        return cls


for _n in ("Autodesk", "Autodesk.Revit", "Autodesk.Revit.DB",
           "Autodesk.Revit.DB.Mechanical", "Autodesk.Revit.DB.Plumbing",
           "Autodesk.Revit.UI", "Autodesk.Revit.UI.Selection", "System",
           "System.Collections", "System.Collections.Generic", "pyrevit"):
    sys.modules[_n] = _StubModule(_n)
sys.modules["pyrevit"].HOST_APP = type("H", (object,), {"version": "2024"})()

import duct_perf_run as dp      # noqa: E402
import fan_system_effect        # noqa: E402
import fitting_tables           # noqa: E402
import vav_checks               # noqa: E402

FAILS = []
COUNT = [0]


def check(name, cond):
    COUNT[0] += 1
    if not cond:
        FAILS.append(name)
        print("FAIL " + name)


# --- defaults equal the dialog as it opens --------------------------------------
r, err = dp.settings_from_controls("rtu")
check("rtu no errors", err == [])
check("limits", r["limits"] == dict((n, (float(f), float(x))) for n, f, x in dp._DV_ROWS))
check("tol", r["tol_pct"] == 10.0 and r["safety_pct"] == 10.0)
check("report off", r["full_diag"] is False and r["selected_cols"] == set())
check("static off, no filter", r["ext_static"] is False and r["filter_inwc"] == 0.0)
check("scope rtu", r["scope"] == "rtu" and r["vav"] is False and r["exhaust"] is False)
check("curb on rtu", r["fan_effect"]["curb_in"] == fan_system_effect.DEFAULT_CURB_IN)
check("fse on, default basis", r["fan_effect"]["on"] is True
      and r["fan_effect"]["velocity_basis"] == "default")
check("rtu fan fpm", r["fan_effect"]["fan_fpm"] == fan_system_effect.scope_fan_fpm("rtu"))
check("ranked off", r["ranked"] is None)
check("c values 2 places", all(abs(v - round(v, 2)) < 1e-12 for v in r["c_values"].values()))
check("no exhaust drops on rtu", "backdraft" not in r["comp_values"] and "cap" not in r["comp_values"])
check("basis", r["calc_basis"]["air_density"] == 0.075)

r, err = dp.settings_from_controls("rtu", {"ext_static": True})
check("static on takes filter", r["filter_inwc"] == fitting_tables.DEFAULT_RETURN_FILTER_INWC)
r, err = dp.settings_from_controls("rtu", {"ext_static": True, "filter_on": False})
check("filter off", r["filter_inwc"] == 0.0)
r, err = dp.settings_from_controls("vavbox", {"ext_static": True})
check("vav box no filter, no curb", r["filter_inwc"] == 0.0
      and r["fan_effect"]["curb_in"] == 0.0 and r["scope"] == "vavbox")

r, err = dp.settings_from_controls("exhaust", {"ext_static": True, "cap_on": True})
check("exhaust swsi", r["fan_effect"]["fan_type"] == "SWSI")
check("exhaust fpm", r["fan_effect"]["fan_fpm"] == fan_system_effect.scope_fan_fpm("exhaust"))
check("exhaust drops", "backdraft" in r["comp_values"] and "cap" in r["comp_values"])
check("exhaust cap", r["exhaust_cap"] is True and r["exhaust"] is True)
check("backdraft default off", r["exhaust_settings"] == {"backdraft": False})
check("no filter on exhaust", r["filter_inwc"] == 0.0)

r, err = dp.settings_from_controls("vav", {"vav": {"unit_max_esp": 1.0}})
check("vav mp limits", vav_checks.MP_KEY in r["limits"] and r["vav"] is True)
check("vav esp", r["vav_settings"]["unit_max_esp"] == 1.0)
check("vav min blank", r["vav_settings"]["min_box_inwc"] is None)

r, err = dp.settings_from_controls("rtu", {"ext_static": True, "ranked_on": True,
                                           "rank_return": 3})
check("ranked", r["ranked"] == {"supply": 2, "return": 3})
r, err = dp.settings_from_controls("exhaust", {"ext_static": True, "ranked_on": True,
                                               "rank_supply": 9})
check("exhaust supply rank ignored", r["ranked"]["supply"] == 2)
r, err = dp.settings_from_controls("rtu", {"limits": {"Supply Air": (1000, 0.1)}})
check("limit override", r["limits"]["Supply Air"] == (1000.0, 0.1)
      and r["limits"]["Return Air"] == (600.0, 0.05))

# --- errors -------------------------------------------------------------------------
check("bad scope", dp.settings_from_controls("ahu")[0] is None)
check("unknown key", "unknown" in dp.settings_from_controls("rtu", {"tolerance": 5})[1][0])
check("unknown nested", "c_values" in dp.settings_from_controls(
    "rtu", {"c_values": {"nope": 1}})[1][0])
check("tol range", dp.settings_from_controls("rtu", {"tol_pct": 100})[0] is None)
check("bad rank", dp.settings_from_controls(
    "rtu", {"ext_static": True, "ranked_on": True, "rank_supply": 1})[0] is None)
check("bad fan type", dp.settings_from_controls("rtu", {"fan_type": "TUBE"})[0] is None)
check("negative c", dp.settings_from_controls(
    "rtu", {"c_values": {list(dp.default_controls()["c_values"])[0]: -1}})[0] is None)

# --- run_headless refusals and the ceiling fan question ----------------------------
r = dp.run_headless(None, 5, scope="nope")
check("refused scope", r["status"] == "refused")
r = dp.run_headless(None, 5, detail="some")
check("refused detail", r["status"] == "refused")


class Fan(object):
    Id = 77


dp._element = lambda d, i: Fan() if i == 77 else None
dp.hvac_graph.is_equipment = lambda e: True
dp.hvac_graph.is_exhaust_fan = lambda e: True
dp.hvac_graph.is_ceiling_exhaust_fan = lambda e: True
dp.hvac_graph.fan_rated_cfm = lambda e: 150.0
dp._elem_name = lambda e: "EF-1"
dp.eid_int = lambda i: int(i)
r = dp.run_headless(None, 5, scope="exhaust")
check("missing unit refused", r["status"] == "refused" and "No element" in r["errors"][0])
r = dp.run_headless(None, 77, scope="rtu")
check("wrong scope refused", r["status"] == "refused" and "not listed" in r["errors"][0])
r = dp.run_headless(None, 77, scope="exhaust")
check("fan cfm asked", r["status"] == "needs_decision"
      and r["decision"]["qid"] == "fan_cfm" and r["decision"]["suggested"] == 150.0)
r = dp.run_headless(None, 77, scope="exhaust", fan_cfm=-3)
check("fan cfm checked", r["status"] == "refused")
CALL = []
dp.perf_core = lambda *a, **k: CALL.append((a, k)) or dict(dp._core_result(), status="no_results")
r = dp.run_headless(None, 77, scope="exhaust", fan_cfm=150)
check("fan cfm passed", CALL[-1][0][3] == {77: {"cfm": 150.0}}
      and CALL[-1][1]["make_sheets"] is False)
check("no results refused", r["status"] == "refused" and r["data"]["ducts"] == [])

print("{} checks, {} failed".format(COUNT[0], len(FAILS)))
sys.exit(1 if FAILS else 0)
