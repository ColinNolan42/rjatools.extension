# -*- coding: ascii -*-
"""Tests for the Size Gas fitting fixes and the 402.4(13) limit (2026-11-04).

Runs in plain CPython with the Revit assemblies stubbed (the stubs come from
_test_prv_zones). Covers:
  - get_created_phase_id / is_new_construction on an element Revit deleted
    (every property read raises) return None / False, never raise;
  - is_valid_element and find_replacement_fitting (the fitting Revit placed
    on the end of a pipe in place of a deleted one);
  - revit_helpers.is_transition (part type first, family name last) and
    pipes_joined, on fake fittings (2026-11-19);
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


class _Model(object):
    def __init__(self, part_type):
        self._pt = part_type

    @property
    def PartType(self):
        if isinstance(self._pt, Exception):
            raise self._pt
        return self._pt


class _Fitting(_Elem):
    """A fitting with a MEPModel part type and a family name."""
    def __init__(self, eid, part_type=None, family=None, valid=True):
        _Elem.__init__(self, eid, "FamilyInstance", valid)
        self._family = family
        if part_type is not None:
            self.MEPModel = _Model(part_type)

    @property
    def Symbol(self):
        family = type("Family", (), {"Name": self._family})()
        return type("Symbol", (), {"Family": family})()


class TestIsTransition(unittest.TestCase):
    """is_transition reads the part type first, the family name last."""

    def test_part_type_transition(self):
        self.assertTrue(revit_helpers.is_transition(
            _Fitting(1, "Transition", family="Elbow - Generic")))

    def test_part_type_wins_over_the_name(self):
        # A defined part type that is not Transition is trusted.
        self.assertFalse(revit_helpers.is_transition(
            _Fitting(2, "Elbow", family="Transition Elbow")))
        self.assertFalse(revit_helpers.is_transition(
            _Fitting(3, "Tee", family="Transition Tee")))

    def test_name_fallback_only_when_no_part_type(self):
        self.assertTrue(revit_helpers.is_transition(
            _Fitting(4, None, family="Transition - Generic")))
        self.assertTrue(revit_helpers.is_transition(
            _Fitting(5, "Undefined", family="TRANSITION - Steel")))
        self.assertFalse(revit_helpers.is_transition(
            _Fitting(6, None, family="Elbow - Generic")))

    def test_unreadable_part_type_falls_to_the_name(self):
        f = _Fitting(7, RuntimeError("no model"), family="Transition - Gen")
        self.assertTrue(revit_helpers.is_transition(f))

    def test_family_content_part_type_is_the_second_read(self):
        # No MEPModel part type; the family parameter says Transition (7)
        # even though the name does not.
        class _PT(object):
            Transition = 7
            Elbow = 3
        db = sys.modules["Autodesk.Revit.DB"]
        db.PartType = _PT
        revit_helpers.BuiltInParameter.FAMILY_CONTENT_PART_TYPE = "FCPT"
        try:
            class _F(_Fitting):
                value = None

                @property
                def Symbol(self):
                    val = self.value
                    fam = type("Family", (), {
                        "Name": "Generic Fitting",
                        "get_Parameter": lambda me, bip: type(
                            "P", (), {"AsInteger": lambda m: val})()})()
                    return type("Sym", (), {"Family": fam})()

            def fitting(value):
                f = _F(20, None, family="Generic Fitting")
                f.value = value
                return f
            self.assertTrue(revit_helpers.is_transition(fitting(7)))
            self.assertFalse(revit_helpers.is_transition(fitting(3)))
        finally:
            del db.PartType
            del revit_helpers.BuiltInParameter.FAMILY_CONTENT_PART_TYPE

    def test_valve_is_not_a_transition(self):
        self.assertFalse(revit_helpers.is_transition(
            _Fitting(8, "Valve", family="RJA - Gas Shutoff Valve (Ansul)")))
        self.assertFalse(revit_helpers.is_transition(
            _Fitting(9, None, family="RJA - Gas Shutoff Valve (Ansul)")))

    def test_deleted_none_and_unreadable_never_raise(self):
        self.assertFalse(revit_helpers.is_transition(None))
        self.assertFalse(revit_helpers.is_transition(
            _Fitting(10, "Transition", valid=False)))
        self.assertFalse(revit_helpers.is_transition(_Elem(11)))


class TestPipesJoined(unittest.TestCase):

    def _pair(self, joined=True):
        a = _Elem(1, "Pipe")
        b = _Elem(2, "Pipe")
        a.connectors = [_Conn((0, 0, 0), []),
                        _Conn((5, 0, 0), [a, b] if joined else [a])]
        b.connectors = [_Conn((5, 0, 0), [a, b] if joined else [b]),
                        _Conn((9, 0, 0), [])]
        return a, b

    def test_pipes_that_now_connect(self):
        self.assertTrue(revit_helpers.pipes_joined(list(self._pair(True))))

    def test_pipes_with_open_ends_are_not_joined(self):
        self.assertFalse(revit_helpers.pipes_joined(list(self._pair(False))))

    def test_deleted_pipe_and_single_pipe(self):
        a, b = self._pair(True)
        b.IsValidObject = False
        self.assertFalse(revit_helpers.pipes_joined([a, b]))
        self.assertFalse(revit_helpers.pipes_joined([a]))
        self.assertFalse(revit_helpers.pipes_joined([]))

    def test_a_fitting_between_is_not_joined(self):
        a, b = self._pair(False)
        fit = _Elem(50)
        a.connectors[1] = _Conn((5, 0, 0), [a, fit])
        b.connectors[0] = _Conn((5, 0, 0), [b, fit])
        self.assertFalse(revit_helpers.pipes_joined([a, b]))


class TestEndsConnected(unittest.TestCase):
    """A transition Revit removed or replaced: the pipe end where it sat is
    joined to something (Horse Creek, 2026-11-19)."""

    def test_end_joined_to_a_new_fitting(self):
        new_fit = _Elem(77)
        pipe = _pipe_with_new_fitting(new_fit)
        self.assertTrue(revit_helpers.ends_connected([pipe], [5.1, 0, 0]))

    def test_open_end_is_a_break(self):
        pipe = _Elem(10, "Pipe")
        pipe.connectors = [_Conn((0, 0, 0), []), _Conn((5, 0, 0), [])]
        self.assertFalse(revit_helpers.ends_connected([pipe], [5.0, 0, 0]))

    def test_only_itself_on_the_connector_is_open(self):
        pipe = _Elem(10, "Pipe")
        pipe.connectors = [_Conn((0, 0, 0), []), _Conn((5, 0, 0), [pipe])]
        self.assertFalse(revit_helpers.ends_connected([pipe], [5.0, 0, 0]))

    def test_no_live_pipe_or_no_location(self):
        pipe = _pipe_with_new_fitting(_Elem(77))
        pipe.IsValidObject = False
        self.assertFalse(revit_helpers.ends_connected([pipe], [5.0, 0, 0]))
        self.assertFalse(revit_helpers.ends_connected([], [5.0, 0, 0]))
        self.assertFalse(revit_helpers.ends_connected(
            [_pipe_with_new_fitting(_Elem(77))], None))


class _Edge(object):
    def __init__(self, eid, a, b, load):
        self.element_id, self.from_node_id, self.to_node_id = eid, a, b
        self.cumulative_load_mbh = load


class _Graph(object):
    def __init__(self, edges, children):
        self.edges = dict((e.element_id, e) for e in edges)
        self.node_children = children


class TestFittingTargetSize(unittest.TestCase):
    """Horse Creek 2026-11-19: ball valve 3633493 drawn 1-1/4" on a 2" line,
    a transition on each side (3661791, 3661818), no pipe touching it."""

    def _valve_between_transitions(self):
        # tee 1 -> pipe 100 (2") -> T 2 -> valve 3 -> T 4 -> pipe 101 -> cap
        edges = [_Edge(100, 1, 2, 400.0), _Edge(101, 4, 5, 400.0)]
        return _Graph(edges, {2: [3], 3: [4]})

    def test_direct_pipe_wins(self):
        g = _Graph([_Edge(100, 1, 3, 400.0), _Edge(102, 3, 9, 90.0)], {})
        self.assertEqual(sizing_engine.fitting_target_size(
            3, g, {100: '2"', 102: '1"'}, {2}), ('2"', False))

    def test_valve_reads_through_its_transitions(self):
        g = self._valve_between_transitions()
        self.assertEqual(sizing_engine.fitting_target_size(
            3, g, {100: '2"', 101: '2"'}, {2, 4}), ('2"', True))

    def test_not_through_a_fitting_that_is_not_a_transition(self):
        g = self._valve_between_transitions()
        self.assertEqual(sizing_engine.fitting_target_size(
            3, g, {100: '2"', 101: '2"'}, set([99])), (None, False))

    def test_unsized_pipes_give_nothing(self):
        g = self._valve_between_transitions()
        self.assertEqual(sizing_engine.fitting_target_size(
            3, g, {}, {2, 4}), (None, False))


class TestSizedTypes(unittest.TestCase):
    """Valves made with one Type per size (RJA - Ball Valve - .25 thru 3,
    Types '0 1/4"' to '3"', size in the Type parameter 'Nominal Pipe Size'):
    Size Gas switches the Type (Horse Creek, 2026-11-19)."""

    def test_parse_size_texts(self):
        p = revit_helpers.parse_size_inches
        self.assertEqual(p('1 1/2"'), 1.5)
        self.assertEqual(p('0 3/4"'), 0.75)
        self.assertEqual(p('1-1/4"'), 1.25)
        self.assertEqual(p('2"'), 2.0)
        self.assertEqual(p('3/8"'), 0.375)
        self.assertEqual(p('2.5 in'), 2.5)
        self.assertIsNone(p("Standard"))
        self.assertIsNone(p(""))
        self.assertIsNone(p(None))
        self.assertIsNone(p('0"'))

    def test_pick_matching_type(self):
        cands = [("half", 0.5), ("one-half", 1.5), ("two", 2.0), ("x", None)]
        self.assertEqual(revit_helpers.pick_sized_type(cands, 1.5), "one-half")
        self.assertEqual(revit_helpers.pick_sized_type(cands, 2.0), "two")
        self.assertIsNone(revit_helpers.pick_sized_type(cands, 1.25))


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
