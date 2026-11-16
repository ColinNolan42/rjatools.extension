# -*- coding: utf-8 -*-
"""
Diagnose.pushbutton/script.py

One diagnose button for all three disciplines. Pick the gas meter, the
domestic water RPZ (backflow preventer), or HVAC equipment/a duct/terminal.
The discipline is auto-detected from what was picked, the matching traversal
runs, and a summary of what the model contains is printed. Connectivity
check, not a sizer: no sizes are written. No Revit model changes are made.

The pipeline lives in lib/diagnose_run.py (shared with the headless entry
diagnose_run.run_headless); this button is the pick.

IronPython 2.7 / pyRevit
"""

import os
import sys

from pyrevit import script, forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

import diagnose_run

diagnose_run.run(doc, uidoc, output, forms)
