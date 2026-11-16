# -*- coding: ascii -*-
# Size Gas.pushbutton/script.py
# IFGC Longest Run gas pipe sizing.
# Startup dialog -> user picks the gas meter, or the first new pipe on an
# existing building -> sizing -> writes sizes to Revit.
#
# The pipeline lives in lib/gas_run.py (shared with the headless entry
# gas_run.run_headless); this button is the dialog and the pick.
#
# IronPython 2.7 / PyRevit

import os
import sys

from pyrevit import script, forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

_lib_dir = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

import gas_run

gas_run.run(doc, uidoc, output, forms)
