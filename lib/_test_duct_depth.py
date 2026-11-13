# -*- coding: ascii -*-
"""Plain CPython tests for duct_sizing.available_depth_in and size_in_depth
(AI design pathway: size to the limits AND the space). Uses a simple
friction stand in, so these check the search rules, not the friction values
(those are covered by hvac_graph and _test_duct_sizing).

Run:  python _test_duct_depth.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import duct_sizing as ds  # noqa: E402

FAILS = []
COUNT = [0]


def check(name, cond):
    COUNT[0] += 1
    if not cond:
        FAILS.append(name)
        print("FAIL " + name)


def no_friction(fpm, dh):
    return 0.0


# depth: (12.0 - 10.0) ft = 24 in, minus 4 in hangers = 20 in
check("depth", abs(ds.available_depth_in(12.0, 10.0) - 20.0) < 1e-9)
check("depth never negative", ds.available_depth_in(10.0, 10.0) == 0.0)
check("depth custom hanger", abs(ds.available_depth_in(11.0, 10.0, 2.0) - 10.0) < 1e-9)

# 1000 cfm at 800 fpm needs 1.25 ft2 = 180 in2
size, why = ds.size_in_depth(1000, 20.0, 800, 0.08, no_friction)
check("rect at depth", size == (10, 20) and why is None)        # 200 in2
size, why = ds.size_in_depth(1000, 11.9, 800, 0.08, no_friction)
check("pancaked", size == (18, 10))                               # 180 in2
size, why = ds.size_in_depth(1000, 20.0, 800, 0.08, no_friction,
                             preferred_h_in=12)
check("preferred shallower", size == (16, 12))                    # 192 in2
size, why = ds.size_in_depth(3000, 8.0, 800, 0.08, no_friction)
check("too shallow asks", size is None and "4:1" in why)          # 540 in2 > 32x8
size, why = ds.size_in_depth(500, 5.0, 800, 0.08, no_friction)
check("under minimum", size is None and "minimum" in why)
size, why = ds.size_in_depth(1000, 20.0, 800, 0.08, no_friction,
                             min_size=(24, 12))
check("covers face", size == (24, 20))
# round: 1000 cfm needs 1.25 ft2 -> 16 in (1.396 ft2)
size, why = ds.size_in_depth(1000, 20.0, 800, 0.08, no_friction,
                             shape=ds.SHAPE_ROUND)
check("round fits", size == 16 and why is None)
size, why = ds.size_in_depth(1000, 14.0, 800, 0.08, no_friction,
                             shape=ds.SHAPE_ROUND)
check("round too deep", size is None and "16 in round" in why)
check("no depth", ds.size_in_depth(1000, None, 800, 0.08, no_friction)[0] is None)


def friction(fpm, dh):
    return 0.10 if dh < 20 else 0.05


size, why = ds.size_in_depth(1000, 11.9, 800, 0.08, friction)
check("friction stops pancake", size is None)

print("{} checks, {} failed".format(COUNT[0], len(FAILS)))
sys.exit(1 if FAILS else 0)
