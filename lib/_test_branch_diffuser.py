# -*- coding: utf-8 -*-
"""Tests for the branch diffuser check (hvac_graph.branch_diffuser_check).

RJA rules (Colin, 2026-11-02):
  - Ceiling and sidewall diffusers from ANY family are judged on the RJA
    diffuser tables (MEP Design Standards), not only the RJA auto sizing
    families.
  - An undersized diffuser is RED and the reason names the size it needs.
  - A Flow above the largest published size is RED, 'Max CFM reached,
    split diffuser'.

hvac_graph imports the Revit API, so the API and revit_helpers are stubbed
and the element lookups the check makes are replaced; the table logic is the
real module's.

Run:  python _test_branch_diffuser.py
"""

import os
import sys
import types
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.dont_write_bytecode = True


class _Enum(object):
    def __getattr__(self, name):
        return abs(hash(name)) % 100000


_db = types.ModuleType('Autodesk.Revit.DB')
for _n in ('BuiltInCategory', 'BuiltInParameter', 'ConnectorProfileType',
           'FlowDirectionType', 'ConnectorType', 'Domain'):
    setattr(_db, _n, _Enum())
for _n in ('FillPatternElement', 'ElementId', 'MEPSystem'):
    setattr(_db, _n, type(_n, (object,), {}))
sys.modules.setdefault('Autodesk', types.ModuleType('Autodesk'))
sys.modules.setdefault('Autodesk.Revit', types.ModuleType('Autodesk.Revit'))
sys.modules['Autodesk.Revit.DB'] = _db
_rh = types.ModuleType('revit_helpers')
_rh.eid_int = lambda x: getattr(x, 'i', x)
sys.modules.setdefault('revit_helpers', _rh)

import hvac_graph as HG  # noqa: E402


class _Id(object):
    def __init__(self, i):
        self.i = i


class _Terminal(object):
    def __init__(self, geom, category, sys_class='Supply Air'):
        self.Id = _Id(77)
        self.geom, self.category, self.sys_class = geom, category, sys_class


_REAL_CLASSIFY = HG.classify_diffuser_branch
HG.terminal_connector_geometry = lambda t: t.geom
HG.classify_diffuser_branch = lambda t: t.category
HG.terminal_sys_class = lambda t: t.sys_class


def check(term, cfm, duct):
    return HG.branch_diffuser_check(term, cfm, duct)


CEIL10 = _Terminal(('round', 10.0), 'ceiling')
CEIL8 = _Terminal(('round', 8.0), 'ceiling')
CEIL14 = _Terminal(('round', 14.0), 'ceiling')


class TestCeiling(unittest.TestCase):

    def test_adequate_neck_passes(self):
        r = check(CEIL10, 300.0, 10.0)
        self.assertEqual(r.status, 'GREEN')

    def test_undersized_neck_names_the_size_needed(self):
        r = check(CEIL8, 300.0, 8.0)
        self.assertEqual(r.status, 'RED')
        self.assertTrue(r.reason.startswith('Diffuser undersized for CFM, use '), r.reason)
        self.assertIn(r.required_size, r.reason)
        self.assertIn('10', r.required_size)

    def test_undersized_neck_and_branch(self):
        r = check(CEIL8, 300.0, 6.0)
        self.assertEqual(r.status, 'RED')
        self.assertTrue(r.reason.startswith('Diffuser + branch undersized, use '), r.reason)

    def test_over_largest_neck_is_split_diffuser(self):
        # Ceiling supply table tops out at 14 in., 750 cfm.
        r = check(CEIL14, 800.0, 14.0)
        self.assertEqual(r.status, 'RED')
        self.assertEqual(r.reason, 'Max CFM reached, split diffuser')

    def test_branch_smaller_than_neck(self):
        r = check(CEIL10, 300.0, 8.0)
        self.assertEqual(r.status, 'RED')
        self.assertEqual(r.reason, 'Branch smaller than diffuser')


class _Param(object):
    def __init__(self, ft):
        self.ft = ft
        self.HasValue = True

    def AsDouble(self):
        return self.ft


class _Symbol(object):
    """Type parameters "Diffuser Width" / "Diffuser Height", in feet."""
    def __init__(self, w_in, h_in):
        self.p = {'Diffuser Width': _Param(w_in / 12.0),
                  'Diffuser Height': _Param(h_in / 12.0)}

    def LookupParameter(self, name):
        return self.p.get(name)


def _ceiling(neck, w_in=None, h_in=None, sys_class='Supply Air'):
    t = _Terminal(('round', float(neck)), 'ceiling', sys_class)
    t.LookupParameter = lambda name: None
    if w_in is not None:
        t.Symbol = _Symbol(w_in, h_in)
    return t


class TestFamilyBands(unittest.TestCase):
    """Tables follow the auto sizing families, 2026-11-02: the family steps to
    the next neck AT its first cfm, and a module's last neck runs to its cap."""

    def test_24x24_steps_at_110(self):
        self.assertEqual(check(_ceiling(6, 24, 24), 109.0, 6.0).status, 'GREEN')
        self.assertEqual(check(_ceiling(6, 24, 24), 110.0, 6.0).status, 'RED')

    def test_12x12_8in_runs_to_200_cap(self):
        self.assertEqual(check(_ceiling(8, 12, 12), 200.0, 8.0).status, 'GREEN')
        r = check(_ceiling(8, 12, 12), 201.0, 8.0)
        self.assertEqual((r.status, r.reason), ('RED', 'Max CFM reached, split diffuser'))

    def test_24x24_8in_at_200_needs_10in(self):
        r = check(_ceiling(8, 24, 24), 200.0, 8.0)
        self.assertEqual(r.status, 'RED')
        self.assertIn('10', r.required_size)

    def test_no_module_uses_24x24(self):
        self.assertEqual(check(_ceiling(8), 200.0, 8.0).status, 'RED')

    def test_24x12_grate_10in_to_325(self):
        self.assertEqual(check(_ceiling(10, 24, 12, 'Return Air'), 325.0, 10.0).status,
                         'GREEN')
        self.assertEqual(check(_ceiling(10, 24, 24, 'Return Air'), 325.0, 10.0).status,
                         'RED')


class TestSidewall(unittest.TestCase):

    def test_over_largest_face_is_split_diffuser(self):
        term = _Terminal(('rect', 24.0, 12.0), 'sidewall')
        r = check(term, 100000.0, (24.0, 12.0))
        self.assertEqual(r.status, 'RED')
        self.assertEqual(r.reason, 'Max CFM reached, split diffuser')


class TestClassifyRectNeck(unittest.TestCase):
    """A rectangular neck pointing up or down is 'rect_ceiling', never a
    sidewall (Colin, 2026-11-02). The real classify_diffuser_branch, with the
    connector reads stubbed."""

    def setUp(self):
        self._slot = HG._is_slot_diffuser
        self._horiz = HG._terminal_neck_is_horizontal
        HG._is_slot_diffuser = lambda t: False
        HG._terminal_neck_is_horizontal = lambda t: t.horizontal

    def tearDown(self):
        HG._is_slot_diffuser = self._slot
        HG._terminal_neck_is_horizontal = self._horiz

    def kind(self, geom, horizontal):
        term = _Terminal(geom, None)
        term.horizontal = horizontal
        return _REAL_CLASSIFY(term)

    def test_rect_neck_down_is_rect_ceiling(self):
        self.assertEqual(self.kind(('rect', 12.0, 12.0), False), 'rect_ceiling')

    def test_rect_neck_sideways_is_sidewall(self):
        self.assertEqual(self.kind(('rect', 20.0, 8.0), True), 'sidewall')

    def test_rect_neck_direction_unknown_stays_sidewall(self):
        self.assertEqual(self.kind(('rect', 20.0, 8.0), None), 'sidewall')

    def test_round_necks_unchanged(self):
        self.assertEqual(self.kind(('round', 10.0), False), 'ceiling')
        self.assertEqual(self.kind(('round', 10.0), True), 'round_sidewall')

    def test_rect_ceiling_is_not_judged_on_a_table(self):
        term = _Terminal(('rect', 12.0, 12.0), 'rect_ceiling')
        r = check(term, 300.0, (12.0, 12.0))
        self.assertEqual(r.status, 'GRAY')


if __name__ == '__main__':
    unittest.main()
