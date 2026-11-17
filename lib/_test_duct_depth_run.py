# -*- coding: ascii -*-
"""Tests for duct_depth_run, the agent_api duct space check.

Uses the REAL friction function (the friction block of hvac_graph.py run on
its own, as _test_duct_sizing does) and the real firm limits.

Run:  python _test_duct_depth_run.py
"""

import io
import math
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import duct_depth_run as dd  # noqa: E402
import duct_sizing as ds  # noqa: E402


def _load_hvac():
    src = io.open(os.path.join(HERE, "hvac_graph.py"), encoding="utf-8").read()
    start = src.index("STANDARD_AIR_DENSITY_LB_FT3 = ")
    end = src.index("def fitting_family_name(")
    ns = {"math": math}
    exec(src[start:end], ns)
    firm_src = src[src.index("FIRM_DEFAULTS = {"):]
    firm_src = firm_src[:firm_src.index("}") + 1]
    fns = {}
    exec(firm_src, fns)
    return ns["duct_friction_loss_per_100ft"], fns["FIRM_DEFAULTS"]


FRICTION, FIRM = _load_hvac()


def row(**kw):
    r = {"id": "S1", "cfm": 1000, "structure_bottom_elev_ft": 14.0,
         "ceiling_elev_ft": 10.0, "shape": "rect", "system": "Supply Air",
         "preferred_h_in": 12}
    r.update(kw)
    return r


def run(rows, settings=None):
    return dd.run_headless("DOC", rows=rows, settings=settings,
                           friction_fn=FRICTION, firm=FIRM)


class TestDepthRun(unittest.TestCase):

    def test_supply_fits(self):
        env = run([row()])
        self.assertEqual(env["status"], "ok")
        r = env["data"]["rows"][0]
        # 4 ft between structure and ceiling, minus 4 in hangers = 44 in
        self.assertEqual(r["depth_in"], 44.0)
        self.assertTrue(r["fits"])
        w, h = r["size_in"]
        # the design height holds; the space is deeper
        self.assertEqual(h, 12)
        self.assertTrue(w % 2 == 0 and w <= 4 * h)
        self.assertLessEqual(r["fpm"], 800)
        self.assertLessEqual(r["friction"], 0.08)
        self.assertEqual(r["size"], "{}x{}".format(w, h))

    def test_shallow_has_no_fit(self):
        # 8 in between structure and ceiling, 4 in left after hangers
        env = run([row(structure_bottom_elev_ft=10.0 + 8.0 / 12.0)])
        self.assertEqual(env["status"], "needs_decision")
        self.assertEqual(env["decision"]["qid"], "no_fit")
        r = env["data"]["rows"][0]
        self.assertFalse(r["fits"])
        self.assertIn("under the 6 in minimum", r["reason"])
        self.assertIn("S1", env["decision"]["text"])

    def test_pancake_at_depth(self):
        # 14 in clear after hangers: the height is held to the depth
        env = run([row(cfm=2000, preferred_h_in=16,
                       structure_bottom_elev_ft=10.0 + 18.0 / 12.0)])
        r = env["data"]["rows"][0]
        self.assertTrue(r["fits"])
        self.assertEqual(r["size_in"][1], 14)

    def test_round_and_rect(self):
        env = run([row(id="R", shape="round"), row(id="T", shape="rect")])
        self.assertEqual(env["status"], "ok")
        rnd, rect = env["data"]["rows"]
        self.assertTrue(rnd["size"].endswith("in round"))
        self.assertEqual(rnd["size_in"], ds.size_round(1000, 800, 0.08,
                                                       FRICTION))

    def test_round_too_deep(self):
        env = run([row(shape="round", cfm=4000,
                       structure_bottom_elev_ft=10.0 + 16.0 / 12.0)])
        self.assertEqual(env["status"], "needs_decision")
        self.assertIn("deep", env["data"]["rows"][0]["reason"])

    def test_return_uses_return_limits(self):
        env = run([row(system="Return Air")])
        r = env["data"]["rows"][0]
        self.assertEqual(r["limits"], {"max_fpm": 600, "max_friction": 0.05})
        self.assertLessEqual(r["fpm"], 600)
        sup = run([row()])["data"]["rows"][0]
        self.assertGreaterEqual(r["size_in"][0] * r["size_in"][1],
                                sup["size_in"][0] * sup["size_in"][1])

    def test_mixed_rows_return_all(self):
        env = run([row(id="A"), row(id="B", structure_bottom_elev_ft=10.5)])
        self.assertEqual(env["status"], "needs_decision")
        self.assertEqual([r["fits"] for r in env["data"]["rows"]],
                         [True, False])

    def test_settings(self):
        env = run([row()], settings={"limits": {"Supply Air": [600, 0.05]},
                                     "hanger_in": 2})
        r = env["data"]["rows"][0]
        self.assertEqual(r["depth_in"], 46.0)
        self.assertEqual(r["limits"]["max_fpm"], 600)
        self.assertIn("Supply Air", env["data"]["basis"]["limits"])

    def test_refusals(self):
        bad = [
            ([row(cfm="1000")], "cfm is not a number"),
            ([row(cfm=True)], "cfm is not a number"),
            ([row(cfm=0)], "cfm shall be above 0"),
            ([row(structure_bottom_elev_ft=float("nan"))], "not a number"),
            ([row(ceiling_elev_ft=float("inf"))], "not a number"),
            ([row(preferred_h_in=float("nan"))], "preferred_h_in"),
            ([row(system="Kitchen")], "system"),
            ([row(shape="oval")], "shape"),
            ([row(id=None)], "id is missing"),
            ([row(preferred_h_in=None)], "preferred_h_in"),
            ([row(preferred_h_in=4)], "preferred_h_in"),
            ([row(), row()], "unique"),
            ([], "non empty"),
            ("rows", "non empty"),
        ]
        for rows, text in bad:
            env = run(rows)
            self.assertEqual(env["status"], "refused", text)
            self.assertTrue(any(text in e for e in env["errors"]),
                            (text, env["errors"]))
        for settings, text in (({"limits": {"Kitchen": [1, 1]}}, "unknown"),
                               ({"limits": {"Supply Air": [0, 0.08]}}, "above 0"),
                               ({"hanger_in": -1}, "hanger_in"),
                               ({"hanger_in": float("nan")}, "hanger_in"),
                               ({"limits": {"Supply Air": [float("inf"), 1]}},
                                "above 0"),
                               ({"limits": ["x"]}, "limits shall be"),
                               ({"speed": 1}, "unknown setting"),
                               ("x", "object")):
            env = run([row()], settings=settings)
            self.assertEqual(env["status"], "refused", text)
            self.assertTrue(any(text in e for e in env["errors"]), text)
        env = run([row(id=str(i)) for i in range(dd.MAX_ROWS + 1)])
        self.assertEqual(env["status"], "refused")

    def test_envelope_shape(self):
        env = run([row()])
        self.assertTrue(set(env) >= {"tool", "status", "data", "warnings",
                                     "errors", "created"})
        self.assertEqual(env["tool"], "duct_size_in_depth")


if __name__ == "__main__":
    unittest.main(verbosity=1)
