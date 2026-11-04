# -*- coding: ascii -*-
"""Tests for the Size Gas fitting fixes and the 402.4(13) limit (2026-11-04).

Runs in plain CPython with the Revit assemblies stubbed (the stubs come from
_test_prv_zones). Covers:
  - get_created_phase_id / is_new_construction on an element Revit deleted
    (every property read raises) return None / False, never raise;
  - is_valid_element and find_replacement_fitting (the fitting Revit placed
    on the end of a pipe in place of a deleted one);
  - sizing_engine.table_limit_warnings for Table 402.4(13) (150 CFH).
It does NOT run Size Gas's _resize_fittings (the pushbutton imports pyRevit)
or a real Revit model. Live behaviour still needs a Revit run.

Run:  python -I _test_gas_fittings.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_prv_zones import _Id   # noqa: E402  (installs the Revit stubs)

import revit_helpers   # noqa: E402
import sizing_engine   # noqa: E402


class _TypeName(object):
    def __init__(self, name):
        self.Name = name


class _Elem(object):
    """A stand-in element with a class name, a validity flag and connectors."""
    def __init__(self, eid, type_name="FamilyInstance", valid=True,
                 connectors=None):
        self._eid = eid
        self._type_name = type_name
        self.IsValidObject = valid
        self.connectors = connectors or []

    @property
    def Id(self):
        if not self.IsValidObject:
            raise Exception("The referenced object is not valid")
        return _Id(self._eid)

    @property
    def CreatedPhaseId(self):
        if not self.IsValidObject:
            raise Exception("The referenced object is not valid")
        return _Id(200)

    def GetType(self):
        return _TypeName(self._type_name)

    @property
    def ConnectorManager(self):
        return _Manager(self.connectors)


class _Manager(object):
    def __init__(self, connectors):
        self.Connectors = connectors


class _XYZ(object):
    def __init__(self, x, y, z):
        self.X, self.Y, self.Z = x, y, z


class _Ref(object):
    def __init__(self, owner):
        self.Owner = owner


class _Conn(object):
    def __init__(self, xyz, owners):
        self.Origin = _XYZ(*xyz)
        self.IsConnected = bool(owners)
        self.AllRefs = [_Ref(o) for o in owners]


class _Phases(object):
    def __init__(self, ids):
        self._ids = ids
        self.Size = len(ids)

    def __iter__(self):
        return iter([type("P", (), {"Id": _Id(i)})() for i in self._ids])

    def get_Item(self, i):
        return type("P", (), {"Id": _Id(self._ids[i])})()


class _Doc(object):
    def __init__(self, ids):
        self.Phases = _Phases(ids)


def _pipe_with_new_fitting(new_fitting, system=None, pipe_id=10):
    """A pipe from (0,0,0) to (5,0,0). Its (5,0,0) end now connects to
    new_fitting (and to the piping system element, as AllRefs does)."""
    pipe = _Elem(pipe_id, "Pipe")
    owners = [pipe, new_fitting] + ([system] if system else [])
    pipe.connectors = [_Conn((0, 0, 0), []), _Conn((5, 0, 0), owners)]
    return pipe


class TestDeletedElementNeverRaises(unittest.TestCase):

    def test_created_phase_of_deleted_element_is_none(self):
        gone = _Elem(5, valid=False)
        self.assertIsNone(revit_helpers.get_created_phase_id(gone))

    def test_deleted_element_is_not_new_construction(self):
        gone = _Elem(5, valid=False)
        self.assertFalse(revit_helpers.is_new_construction(
            gone, _Doc([100, 200])))

    def test_live_element_still_reads_new(self):
        self.assertTrue(revit_helpers.is_new_construction(
            _Elem(5), _Doc([100, 200])))


class TestIsValidElement(unittest.TestCase):

    def test_none_and_deleted(self):
        self.assertFalse(revit_helpers.is_valid_element(None))
        self.assertFalse(revit_helpers.is_valid_element(_Elem(1, valid=False)))

    def test_live_and_stub_without_flag(self):
        self.assertTrue(revit_helpers.is_valid_element(_Elem(1)))
        self.assertTrue(revit_helpers.is_valid_element(object()))


class TestFindReplacementFitting(unittest.TestCase):

    def test_finds_the_new_fitting_at_the_old_location(self):
        new = _Elem(99)
        system = _Elem(500, "PipingSystem")
        pipe = _pipe_with_new_fitting(new, system)
        found, d = revit_helpers.find_replacement_fitting(
            [pipe], [5.2, 0, 0], set([1, 2, 10]))
        self.assertIs(found, new)
        self.assertAlmostEqual(d, 0.2)

    def test_reads_the_end_nearest_the_old_fitting(self):
        new = _Elem(99)
        pipe = _pipe_with_new_fitting(new)
        found, _d = revit_helpers.find_replacement_fitting(
            [pipe], [0.1, 0, 0], set())
        self.assertIsNone(found)   # the (0,0,0) end is open

    def test_a_graph_node_is_never_returned(self):
        known = _Elem(42)
        pipe = _pipe_with_new_fitting(known)
        found, _d = revit_helpers.find_replacement_fitting(
            [pipe], [5, 0, 0], set([42]))
        self.assertIsNone(found)

    def test_deleted_pipe_and_no_location(self):
        pipe = _pipe_with_new_fitting(_Elem(99))
        pipe.IsValidObject = False
        self.assertEqual(revit_helpers.find_replacement_fitting(
            [pipe], [5, 0, 0], set()), (None, None))
        live = _pipe_with_new_fitting(_Elem(99))
        self.assertEqual(revit_helpers.find_replacement_fitting(
            [live], None, set()), (None, None))

    def test_a_pipe_on_the_end_is_not_a_fitting(self):
        other_pipe = _Elem(77, "Pipe")
        pipe = _pipe_with_new_fitting(other_pipe)
        found, _d = revit_helpers.find_replacement_fitting(
            [pipe], [5, 0, 0], set())
        self.assertIsNone(found)


class TestTableLimitWarnings(unittest.TestCase):
    """2024 IFGC Table 402.4(13) Intended Use: 150 CFH single house line
    regulator load (S38, Colin approved 2026-11-04)."""

    def test_at_the_limit_is_silent(self):
        self.assertEqual(
            sizing_engine.table_limit_warnings("402.4(13)", 150.0, 126.0), [])

    def test_over_the_limit_warns(self):
        w = sizing_engine.table_limit_warnings("402.4(13)", 1391.7, 1169.0)
        self.assertEqual(len(w), 1)
        self.assertIn("402.4(13)", w[0])
        self.assertIn("150 CFH", w[0])
        self.assertIn("1392 CFH (1169 MBH)", w[0])
        w[0].encode("ascii")

    def test_other_tables_have_no_limit(self):
        for tid in ("402.4(12)", "402.4(14)", "402.4(10)", "402.4(28)"):
            self.assertEqual(
                sizing_engine.table_limit_warnings(tid, 5000.0, 4200.0), [],
                tid)

    def test_none_demand(self):
        self.assertEqual(
            sizing_engine.table_limit_warnings("402.4(13)", None, 0.0), [])


if __name__ == "__main__":
    unittest.main()
