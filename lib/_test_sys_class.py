# -*- coding: utf-8 -*-
"""Tests for the duct system class and the one supply test (Colin, 2026-11-02).

  - hvac_graph.duct_sys_class: a System Type name that is a firm row is used
    as is (Outside Air and Transfer Air exist only as type names); any other
    type name ('Supply Air 2') takes the System Classification; with no
    classification the type name is the logged fallback.
  - vav_checks.is_supply_class is the one supply test; fitting_tables uses it.

hvac_graph imports the Revit API, so the API and revit_helpers are stubbed,
the same way as _test_branch_diffuser.py.

Run:  python _test_sys_class.py
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

import hvac_graph as HG       # noqa: E402
import fitting_tables as FT   # noqa: E402
import vav_checks as VC       # noqa: E402


class _Id(object):
    def __init__(self, i):
        self.i = i


class _Param(object):
    def __init__(self, value):
        self.value = value

    def AsValueString(self):
        return self.value

    def AsString(self):
        return self.value


class _Duct(object):
    """type_name: the System Type name; cls: the System Classification."""

    def __init__(self, type_name, cls):
        self.Id = _Id(5)
        self.type_name, self.cls = type_name, cls

    def get_Parameter(self, _bip):
        return None if self.type_name is None else _Param(self.type_name)

    def LookupParameter(self, name):
        if name == 'System Classification' and self.cls is not None:
            return _Param(self.cls)
        return None


class TestDuctSysClass(unittest.TestCase):

    def test_numbered_type_name_reads_the_classification(self):
        self.assertEqual(HG.duct_sys_class(_Duct('Supply Air 2', 'Supply Air')),
                         'Supply Air')

    def test_firm_row_type_names_are_kept(self):
        # Revit has no Outside Air or Transfer Air classification.
        self.assertEqual(HG.duct_sys_class(_Duct('Outside Air', 'Supply Air')),
                         'Outside Air')
        self.assertEqual(HG.duct_sys_class(_Duct('Transfer Air', 'Other Air')),
                         'Transfer Air')

    def test_no_classification_falls_back_to_the_type_name(self):
        self.assertEqual(HG.duct_sys_class(_Duct('SA-RTU1', None)), 'SA-RTU1')

    def test_nothing_readable_is_unknown(self):
        self.assertEqual(HG.duct_sys_class(_Duct(None, None)), 'Unknown')

    def test_undefined_entries_are_dropped(self):
        self.assertEqual(HG.duct_sys_class(_Duct('RA 3', 'Return Air,Undefined')),
                         'Return Air')


class TestOneSupplyTest(unittest.TestCase):

    def test_fitting_tables_uses_the_shared_test(self):
        for name in ('Supply Air', 'Supply Air 2', 'supply air (MP)',
                     'Return Air', 'Exhaust Air', 'Outside Air', '', None):
            self.assertEqual(FT._is_supply(name), VC.is_supply_class(name), name)
        self.assertTrue(FT._is_supply('Supply Air 2'))
        self.assertFalse(FT._is_supply('Return Air'))



class _Xyz(object):
    def __init__(self, x, y, z):
        self.X, self.Y, self.Z = x, y, z


class _Conn(object):
    """A junction connector: Shape from the stubbed ConnectorProfileType."""

    def __init__(self, shape, w=1.0, h=0.5, r=0.5):
        self.Shape = shape
        self.Width, self.Height, self.Radius = w, h, r
        self.CoordinateSystem = type('CS', (object,), {'BasisZ': _Xyz(1.0, 0.0, 0.0)})()


class TestJunctionLegsShape(unittest.TestCase):
    """Review 2026-11-11: junction_legs reads round and rectangular legs
    only; an oval leg is left out (unreadable), so the tee takes the worksheet
    FALLBACK instead of being priced as rectangular with Width x Height."""

    def test_oval_leg_is_unreadable(self):
        cpt = _db.ConnectorProfileType
        conns = {11: _Conn(cpt.Round), 12: _Conn(cpt.Rectangular, 1.0, 0.5),
                 13: _Conn(cpt.Oval, 1.0, 0.5)}
        old = HG._connector_to
        HG._connector_to = lambda elem, nid: conns.get(nid)
        try:
            legs = HG.junction_legs(type('E', (object,), {'Id': _Id(9)})(), [11, 12, 13])
        finally:
            HG._connector_to = old
        self.assertEqual(sorted(legs), [11, 12])
        self.assertEqual(legs[11]['shape'], 'round')
        self.assertEqual(legs[12]['shape'], 'rect')
        self.assertAlmostEqual(legs[12]['area_ft2'], 0.5)


if __name__ == '__main__':
    unittest.main()
