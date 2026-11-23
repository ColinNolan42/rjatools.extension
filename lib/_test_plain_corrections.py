# -*- coding: utf-8 -*-
"""Tests for the plain language System Corrections (duct_perf_console.plain,
plain_groups, corrections_sheet, system_corrections; Colin, 2026-11-23).

Every engine message template is fed through plain(): it shall land in the
right group (Fix the model, Fix the design, For information) with its own
sentence, never the generic one, and the sentence the sheet shows shall hold
no element ids and no ASHRAE table codes. The texts come from the real
builders where they run under CPython (fan_system_effect, the clearance
block of hvac_graph), else from the templates as the engine writes them.

Run:  python _test_plain_corrections.py
"""

import io
import math
import os
import re
import unittest

import duct_perf_console as dpc
import fan_system_effect as F

HERE = os.path.dirname(os.path.abspath(__file__))


def _clearance():
    src = io.open(os.path.join(HERE, 'hvac_graph.py'), encoding='utf-8').read()
    ns = {'math': math}
    exec(src[src.index('CLEARANCE_MARGIN_IN = '):src.index('def effective_height_in(')], ns)
    return ns['clearance_texts']


CLEAR = _clearance()
M, D, I = dpc.MODEL, dpc.DESIGN, dpc.INFO
GENERIC = ('One item needs a check; see the console for the details.',
           'See the console for the details.')
ID_RUN = re.compile(r'\d{4,}')
TABLE_CODE = re.compile(r'\b[A-Z]{2}\d-\d+\b|Cc,|Le\b|Dh\b|\bid\b')

# (kind, engine text, expected group, a word the sentence shall contain)
CASES = [
    # Fix the design
    ('warn', CLEAR('branch', 'Main', 12.0, 11.0, '3640926 (12x12)', '3640950 (11x8)')[1], D,
     'Increase the main to 14 in.'),
    ('warn', CLEAR('neck', 'Branch', 10.0, 10.0, '3640926 (10 dia)', '3641000 (10 dia)')[1], D,
     'reduce the neck to 8 in.'),
    ('warn', CLEAR('step', 'Main', 12.0, 11.0, '3640926 (12x12)', '3640950 (11x12)')[1], D,
     'Reduce the smaller duct to 10 in.'),
    ('red', F.swirl_text(F.INLET, 'return inlet', ((123, 1.0, None), (456, 2.5, None)), 3.0, 6.0),
     D, 'return inlet'),
    ('red', F.swirl_text(F.OUTLET, 'supply outlet', ((123, 1.0, None), (456, 2.5, None)), 3.0,
                         6.0), D, 'turning vanes'),
    ('warn', F.double_elbow_text('supply outlet', [(123, 1.0, None), (456, 2.0, None)], 3.0, 6.0),
     D, 'Two elbows in a row'),
    ('warn', F.damper_window_text('supply outlet', 3650001, 'Balancing Damper', 1.0, 3.5, 'Le'),
     D, 'Move it at least 3.5 ft'),
    ('red', F.le_text('supply outlet', 3650002, 1.2, 3.9, 'Eq. 37'), D,
     'at least 2.0 ft of straight duct'),
    ('red', F.inlet_red_text('return inlet', 3650003, 1.0, 20.0, True), D,
     'at least 3.4 ft of straight duct'),
    ('warn', F.cabinet_warn_text('return inlet', 3650004, 3.8, 20.0, True, 600.0), D,
     'the fan loss may be higher'),
    ('gray', F.high_text(F.outlet_effect('DWDI', 0.5, 4.0, 8000.0, 0.0), 'supply outlet',
                         3650005), D, 'Fan connection loss is too high at the supply outlet'),
    ('warn', F.tap_close_text('supply outlet', 3650006, 1.0, 2.0, '0.5 Le', 3.9, '1.0 Le'), D,
     'A tap is too close to the supply outlet'),
    ('warn', F.disruption_text('tee', 3650007, 'supply outlet', 1.0, 3.5, 'Le'), D,
     'A fitting is too close'),
    ('warn', 'UNPRICED: WARNING: flex duct 3650008 on the exhaust discharge; exhaust '
             'ductwork shall be rigid', D, 'shall be rigid'),
    ('warn', '2 diffuser(s) on an MP main with no box upstream (flagged red on the duct '
             'feeding each).', D, 'medium pressure'),
    ('warn', 'LOW: VAV box VAV-1 (id 3650009), pressure left 0.310 in. wc, under the 0.50 in. '
             'wc minimum.', D, 'VAV box'),
    ('error', 'Unit max external static 0.750 in. wc: computed 0.850 in. wc (Supply Air + '
              'Return Air), EXCEEDS IT.', D, 'above the unit\'s maximum'),
    # Fix the model
    ('warn', '3 terminal(s) have Flow = 0: D-1 (id 3650011), D-2 (id 3650012), id 3650013. '
             'Assign CFM values in the model before running Duct Performance.', M,
     '3 diffusers or grilles have no airflow'),
    ('warn', '1 terminal(s) have Flow = 0: D-1 (id 3650011). Assign CFM values.', M,
     '1 diffuser or grille has no airflow'),
    ('warn', '2 terminal(s) are missing the "Flow" parameter entirely: id 3650014, id 3650015.',
     M, 'no Flow parameter'),
    ('warn', '4 duct(s) have no dimension data — will show gray.', M, '4 ducts have no size'),
    ('error', 'No elements found in traversal. Check that the selected element is connected to '
              'a duct system.', M, 'Nothing is connected'),
    ('error', 'Ceiling exhaust fan discharge has 2 ductwork ends (3650016, 3650017); one run to '
              'one wall or roof cap is needed.', M, 'one wall or roof cap'),
    ('warn', 'No air terminals (OST_DuctTerminal) found in traversal. Check that diffusers are '
             'physically connected to the duct system.', M, 'No diffusers or grilles'),
    ('warn', 'No duct segments found in traversal.', M, 'No ductwork'),
    ('warn', 'Equipment id 3650018: inlet connector direction could not be read, so it is '
             'treated as a VAV box.', M, 'treated as a VAV box'),
    ('warn', 'No base equipment (AHU) found in traversal. CFM sums are computed away from the '
             'selected element.', M, 'Pick the unit'),
    ('warn', 'VAV System is ticked but no VAV box was found on this unit\'s supply. Judged as a '
             'constant volume system.', M, 'Choose Rooftop Unit'),
    ('warn', 'Parallel units on one main. Its external static is unreliable.', M,
     'Two units share'),
    ('warn', 'Supply Air: Take-off rows add to 0.512 in. wc, the run totals 0.511; check the '
             'index run detail.', M, 'does not add up'),
    ('warn', 'Supply Air: Path search hit its step limit, so this may not be the true worst '
             'run.', M, 'too large'),
    ('warn', 'Return Air: 1,200 CFM (6 terminals), but no index run was traced - no static '
             'pressure for it.', M, 'no run from the unit'),
    ('warn', 'WARNING: no mechanical floor plan on Level 2 and one could not be created; its '
             'ductwork is colored but not shown on the sheet.', M, 'Create a mechanical'),
    ('error', 'ERROR: exhaust fan 3650019, no connection reaches a grille or diffuser; the fan '
              'is not computed.', M, 'intake'),
    ('warn', 'WARNING: return inlet size unreadable; fan system effect not checked', M,
     'no readable size'),
    ('warn', 'WARNING: elbow 3650020, 1.0 ft from the return inlet: no size could be read; '
             'straight duct not checked.', M, 'no readable size'),
    ('warn', 'WARNING: tap 3650021, 1.0 ft from the supply outlet, but no size could be read; '
             'fan system effect not priced', M, 'no readable size'),
    ('note', 'NOTE: duct 3650022 (Supply Air, 12"x8") not judged: No airflow data.', M,
     'no airflow reaches them'),
    ('note', 'NOTE: duct 3650023 (Supply Air, -) not judged: No duct size data.', M,
     'no size'),
    ('note', 'NOTE: duct 3650024 (Supply Air, 8") not judged: Cannot check: rect duct on round '
             'neck.', M, 'does not match'),
    ('note', 'NOTE: duct 3650025 (Supply Air, 8") not judged: Cannot classify diffuser.', M,
     'could not be identified'),
    ('warn', 'UNPRICED: Rect Elbow 3650026: unrecognised fitting (UNPRICED)', M,
     'not recognised'),
    ('warn', 'UNPRICED: Rect Transition: transition, size change unreadable (UNPRICED)', M,
     'transition'),
    ('warn', 'UNPRICED: Fire Damper: accessory on the run, not priced', M, 'accessory'),
    ('note', 'NOTE: unit 3650027 connection at element 3650028 is a single diffuser main (one '
             'terminal); its run is judged as a branch.', M, 'only one diffuser'),
    # For information
    ('note', 'NOTE: duct 3650029 (Outside Air, 8") not judged: No table for Outside Air.', I,
     'Outside air'),
    ('warn', 'FALLBACK: taps 3641113, 3641127 (duct 3640926): 2001 ED5-2 Cc,s not positive; '
             'RJA worksheet C 0.60 used', I, 'RJA worksheet value'),
    ('warn', 'FALLBACK: Rectangular Tee: tee, priced as a take-off branch, RJA worksheet C 1.00 '
             '(geometry or flows unreadable)', I, 'RJA worksheet value'),
    ('note', 'NOTE: read at the table edge: transition 3640946 SR4-1 angle 7.6 deg at 10.0 deg',
     I, 'outside the ASHRAE table range'),
    ('note', 'NOTE: read next to a blank cell, nearest printed value used: elbow 12 CD3-12', I,
     'blank cell'),
    ('note', 'NOTE: tap positions on duct 3650030 unreadable: every tap priced at the full duct '
             'flow', I, 'tap positions'),
    ('note', 'NOTE: tap 3650031 position on duct 3650032 unreadable, the whole duct length '
             'counted', I, 'tap positions'),
    ('note', 'NOTE: unit inlet unreadable; window D taken from the first duct (14 in.)', I,
     'could not be read'),
    ('note', 'NOTE: unit supply outlet connection size unreadable; the first duct (14x10) is '
             'used for Ao', I, 'could not be read'),
    ('note', 'NOTE: inlet elbow geometry unreadable; D 14 in. taken from the first duct', I,
     'could not be read'),
    ('note', 'NOTE: elbow 3650033, 1.2 ft from the return inlet, is mitered: ER7-1 read on the '
             'r/H 1.0 row.', I, 'could not be read'),
    ('note', 'NOTE: supply outlet at element 3650034, most restrictive run (D-3, id 3650035): '
             'first elbow 3650036 at 3.4 ft, inside Le 3.5 ft; SR7-11 at the 1500 fpm fan '
             'velocity', I, 'worst run'),
    ('note', 'NOTE: return inlet, run 2 (RG-4, id 3650037): no elbow within 5 D (5.9 ft)', I,
     'worst run'),
    ('note', 'NOTE: unit connection to element 3650038 not checked for fan system effect: it '
             'feeds equipment directly.', I, 'not checked for fan loss'),
    ('note', 'NOTE: fan system effect at the supply outlet not checked, no airflow on its '
             'first duct', I, 'not checked for fan loss'),
    ('note', 'NOTE: fan system effect not priced at a VAV box.', I, 'VAV box'),
    ('note', 'NOTE: exhaust fan read as an inline fan: intake at element 1, discharge at '
             'element 2.', I, 'exhaust fan type'),
    ('note', 'NOTE: read as a ceiling exhaust fan (its only duct connection is Exhaust Air, '
             'out): 90 CFM entered', I, 'exhaust fan type'),
    ('note', '2 Outside Air branch(es) not traversed — equipment-level mode is on.', I,
     'Outside air'),
    ('note', '1 upstream Supply Air branch(es) not traversed — equipment-level mode never '
             'travels back.', I, 'upstream'),
    ('note', 'NOTE: oval duct 3650039 read as rectangular W x H for area, velocity and '
             'friction; oval sizing is not supported.', I, 'Oval'),
    ('note', 'NOTE: no mechanical floor plan on Level 2; one was created from the active '
             'view\'s type and template for this sheet.', I, 'was created'),
    ('warn', 'Supply rank 9 is above the 6 diffuser(s)/grille(s) on this system; run 6 (the '
             'least restrictive) is shown.', I, 'ranked run'),
    ('warn', 'UNPRICED: discharge splits at element 3650040: priced along its first leg only',
     I, 'first leg'),
    ('note', 'NOTE: elbow 3650041, 1.0 ft from the return inlet: ED7-2: elbow radius '
             'unreadable, smallest printed r/D 0.50 used (highest C, conservative)', I,
     'ASHRAE table assumption'),
    ('note', 'NOTE: tap 3764706 on main 3764127 (Supply Air 6x12"): ER5-3 is printed for Ab/Ac '
             '0.50 only and is read there', I, 'ASHRAE table assumption'),
    # review 2026-11-23: an inline fan's place carries 'at element N', a roof
    # curb its clause; the sheet names the place only
    ('red', F.inlet_red_text('exhaust inlet at element 3650100', 3650003, 1.0, 20.0, True), D,
     'between the exhaust inlet and the elbow'),
    ('gray', F.high_text(F.outlet_effect('DWDI', 0.5, 4.0, 8000.0, 0.0),
                         'exhaust outlet at element 3650101', 3650005), D,
     'too high at the exhaust outlet.'),
    ('red', F.swirl_text(F.OUTLET, F.curb_where('supply outlet', 14 / 12.0),
                         ((123, 1.0, None), (456, 2.5, None)), 3.0, 6.0), D,
     'Two elbows at the supply outlet turn'),
    ('warn', F.damper_window_text('exhaust inlet at element 3650102', 3650001, 'Balancing Damper',
                                  1.0, 3.5, 'the inlet window, 5 H'), D,
     'fan at the exhaust inlet.'),
    ('note', 'NOTE: duct 3650024 (Supply Air, 8") not judged: Cannot check: 10" neck not in '
             'table.', M, 'not in the RJA sizing table'),
    ('note', 'NOTE: duct 3650024 (Supply Air, 8") not judged: Cannot check: branch duct size '
             'unreadable.', M, 'could not be read'),
    ('note', 'NOTE: duct 5 (-, 8") not judged: No table for None.', M,
     'could not be identified'),
    ('note', 'NOTE: duct 5 (Transfer Air, 8") not judged: No table for Transfer Air.', I,
     'transfer air'),
    ('note', F.run_note('supply outlet', 'D-3, id 3650103',
                        none='unrecognised fitting 7 at 1.0 ft, inside Le (3.5 ft), flagged, '
                             'not priced'), I, 'worst run'),
]


class PlainTests(unittest.TestCase):

    def test_every_template(self):
        for kind, text, group, word in CASES:
            g, sentence, _est = dpc.plain(kind, text)
            self.assertEqual(g, group, text)
            self.assertNotIn(sentence, GENERIC, text)
            self.assertIn(word.lower(), sentence.lower(), text)

    def test_sheet_sentences_have_no_ids_or_table_codes(self):
        for kind, text, group, _w in CASES:
            _g, sentence, _e = dpc.plain(kind, text)
            self.assertIsNone(ID_RUN.search(sentence), sentence)
            self.assertIsNone(TABLE_CODE.search(sentence), sentence)

    def test_unknown_line_keeps_its_group(self):
        self.assertEqual(dpc.plain('red', 'RED: something new')[:2], (D, GENERIC[0]))
        self.assertEqual(dpc.plain('warn', 'WARNING: something new')[:2], (M, GENERIC[0]))
        self.assertEqual(dpc.plain('note', 'NOTE: something new')[:2], (I, GENERIC[1]))

    def _entries(self, unit='RTU-1 (id 5)'):
        c = dpc.correction
        le = F.le_text('supply outlet', 3650002, 1.2, 3.9, 'Eq. 37')
        return [
            c('red', le, unit),
            c('red', le.replace('3650002', '3650099'), unit),
            c('warn', '3 terminal(s) have Flow = 0: D-1 (id 11), D-2 (id 12), id 13.', unit),
            c('warn', 'FALLBACK: tap 1 passed on duct 4, 2001 ED5-2 Cc,s -0.20 is not '
                      'positive; RJA worksheet C 0.60 used', unit),
            c('warn', 'FALLBACK: tap 2 passed on duct 4, 2001 ED5-2 Cc,s -0.20 is not '
                      'positive; RJA worksheet C 0.60 used', unit),
            c('note', 'NOTE: oval duct 8 read as rectangular W x H.'),
        ]

    def test_repeats_are_one_item_with_details(self):
        g = dpc.plain_groups(self._entries())
        self.assertEqual(len(g[D]), 1)
        _k, sentence, details, _e = g[D][0]
        self.assertTrue(sentence.startswith('Elbow too close to the supply outlet'))
        self.assertEqual(len(details), 2)
        self.assertTrue(all('Elbow 36500' in d for d in details))
        self.assertFalse(any(d.startswith('RED') for d in details))
        # the two tap FALLBACKs are one grouped engine line, one item
        info = [it for it in g[I] if 'worksheet' in it[1]]
        self.assertEqual(len(info), 1)
        self.assertEqual(len(info[0][2]), 1)
        self.assertIn('taps 1, 2 (duct 4)', info[0][2][0])

    def test_sheet(self):
        sheet = dpc.corrections_sheet(self._entries())
        self.assertEqual([h for h, _l in sheet],
                         ['FIX THE MODEL', 'FIX THE DESIGN', 'FOR INFORMATION'])
        self.assertEqual(sheet[2][1], [dpc.ESTIMATE_SHEET_LINE])
        for _h, lines in sheet:
            for line in lines:
                self.assertIsNone(ID_RUN.search(line), line)
                self.assertNotIn('RTU-1', line)          # one unit: no unit name
        self.assertEqual(len(sheet[1][1]), 1)            # two elbows, one line

    def test_sheet_names_the_unit_without_its_id_in_a_multi_unit_run(self):
        entries = self._entries('RTU-1 (id 5)') + self._entries('RTU-2 (id 6)')[:1]
        sheet = dict(dpc.corrections_sheet(entries))
        design = sheet['FIX THE DESIGN']
        self.assertEqual(len(design), 2)
        self.assertTrue(design[0].startswith('RTU-1: Elbow too close'))
        self.assertTrue(design[1].startswith('RTU-2: Elbow too close'))
        self.assertNotIn('(id', ' '.join(design))

    def test_no_information_line_without_an_estimate(self):
        c = dpc.correction
        sheet = dpc.corrections_sheet([c('note', 'NOTE: oval duct 8 read as rectangular.')])
        self.assertEqual(sheet, [])

    def test_console_shows_sentence_then_details(self):
        html = dpc.system_corrections(['#'], [['1']], self._entries())
        for title in ('Fix the model', 'Fix the design', 'For information'):
            self.assertIn(title, html)
        self.assertLess(html.index('Fix the model'), html.index('Fix the design'))
        self.assertLess(html.index('Fix the design'), html.index('For information'))
        i = html.index('Elbow too close to the supply outlet')
        self.assertLess(i, html.index('Details: Elbow 3650002'))
        self.assertEqual(html.count('Elbow too close to the supply outlet'), 1)
        text = dpc.corrections_text(self._entries(), any_rows=True)
        self.assertEqual(text[0], '**Fix the model**')
        self.assertIn('    - Details: Elbow 3650002 is 1.2 ft from the supply outlet, L/Le 0.31 '
                      '(Le 3.9 ft, Eq. 37). Recommended: at least 2.0 ft of straight duct; '
                      '3.9 ft (1.0 Le) removes the system effect. Priced as drawn; the '
                      'engineer has final say where space does not allow.', text)


if __name__ == '__main__':
    unittest.main()
