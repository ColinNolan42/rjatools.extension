# -*- coding: utf-8 -*-
"""Tests for the ceiling exhaust fan detection in hvac_graph (Colin, 2026-11-02).

  - is_ceiling_exhaust_fan: every CONNECTED air connector is Exhaust Air,
    direction Out. The unconnected grille face connector is ignored. An In
    connector (an inline fan) or another system is not a ceiling fan.
  - discharge_ends: the ductwork nodes with no children; the unit
    connector's logical Duct System element is not one.

hvac_graph imports the Revit API, so the API and revit_helpers are stubbed,
the same way as _test_sys_class.py.

Run:  python _test_ceiling_fan.py
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


class _Id(object):
    def __init__(self, i):
        self.i = i


class _Conn(object):
    def __init__(self, system, direction, connected=True):
        self.Domain = HG.Domain.DomainHvac
        self.DuctSystemType = system      # str() of the enum, as in IronPython
        self.Direction = direction
        self.IsConnected = connected


class _Fan(object):
    def __init__(self, conns):
        self.Id = _Id(4845923)
        self.conns = conns


OUT = HG.FlowDirectionType.Out
IN = HG.FlowDirectionType.In


class TestIsCeilingExhaustFan(unittest.TestCase):

    def setUp(self):
        self._hc = HG._hvac_connectors
        HG._hvac_connectors = lambda e: e.conns

    def tearDown(self):
        HG._hvac_connectors = self._hc

    def test_ceiling_exhaust_sp_a(self):
        # As read live on The Hub: 4 in. round Out connected, 6x6 grille
        # face Out unconnected (no flow).
        fan = _Fan([_Conn('ExhaustAir', OUT, False), _Conn('ExhaustAir', OUT)])
        self.assertTrue(HG.is_ceiling_exhaust_fan(fan))

    def test_inline_fan_has_an_in_connector(self):
        fan = _Fan([_Conn('ExhaustAir', IN), _Conn('ExhaustAir', OUT)])
        self.assertFalse(HG.is_ceiling_exhaust_fan(fan))

    def test_supply_out_is_not_an_exhaust_fan(self):
        fan = _Fan([_Conn('SupplyAir', OUT), _Conn('ReturnAir', IN)])
        self.assertFalse(HG.is_ceiling_exhaust_fan(fan))

    def test_nothing_connected_is_not_detected(self):
        fan = _Fan([_Conn('ExhaustAir', OUT, False)])
        self.assertFalse(HG.is_ceiling_exhaust_fan(fan))

    def test_bidirectional_is_not_detected(self):
        # The fallback NOTE case: a family authored Bidirectional reads as
        # an ordinary unit, which then finds no terminals.
        fan = _Fan([_Conn('ExhaustAir', HG.FlowDirectionType.Bidirectional)])
        self.assertFalse(HG.is_ceiling_exhaust_fan(fan))


class TestIsExhaustFan(unittest.TestCase):
    """Hides exhaust fans from the list until Exhaust System is ticked."""

    def setUp(self):
        self._hc = HG._hvac_connectors
        HG._hvac_connectors = lambda e: e.conns

    def tearDown(self):
        HG._hvac_connectors = self._hc

    def test_ceiling_and_inline_fans_are_exhaust_fans(self):
        self.assertTrue(HG.is_exhaust_fan(_Fan([_Conn('ExhaustAir', OUT)])))
        self.assertTrue(HG.is_exhaust_fan(_Fan([_Conn('ExhaustAir', IN),
                                                _Conn('ExhaustAir', OUT)])))

    def test_an_rtu_is_not(self):
        self.assertFalse(HG.is_exhaust_fan(_Fan([_Conn('SupplyAir', OUT),
                                                 _Conn('ReturnAir', IN)])))

    def test_nothing_connected_is_not(self):
        self.assertFalse(HG.is_exhaust_fan(_Fan([_Conn('ExhaustAir', OUT, False)])))


class TestDischargeEnds(unittest.TestCase):

    def setUp(self):
        self._saved = (HG.is_duct, HG.is_fitting_or_accessory, HG.is_terminal)
        HG.is_duct = lambda e: e == 'duct'
        HG.is_fitting_or_accessory = lambda e: e == 'fitting'
        HG.is_terminal = lambda e: e == 'terminal'

    def tearDown(self):
        HG.is_duct, HG.is_fitting_or_accessory, HG.is_terminal = self._saved

    def test_one_open_duct_end(self):
        nodes = {1: 'fan', 2: 'duct', 3: 'fitting', 4: 'duct', 9: 'system'}
        children = {1: [2, 9], 2: [3], 3: [4], 4: [], 9: []}
        self.assertEqual(HG.discharge_ends(1, nodes, children), [4])

    def test_a_modeled_cap_terminal_is_the_end(self):
        nodes = {1: 'fan', 2: 'duct', 5: 'terminal'}
        children = {1: [2], 2: [5], 5: []}
        self.assertEqual(HG.discharge_ends(1, nodes, children), [5])

    def test_a_discharge_joining_other_ductwork_has_two_ends(self):
        nodes = {1: 'fan', 2: 'duct', 3: 'fitting', 4: 'duct', 6: 'duct'}
        children = {1: [2], 2: [3], 3: [4, 6], 4: [], 6: []}
        self.assertEqual(HG.discharge_ends(1, nodes, children), [4, 6])


if __name__ == '__main__':
    unittest.main()
