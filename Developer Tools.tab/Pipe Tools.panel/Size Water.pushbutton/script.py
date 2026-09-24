# -*- coding: ascii -*-
# Size Water.pushbutton/script.py
# Domestic water pipe sizing per the 2024 IPC.
#
# One button. The dialog's two action buttons choose what the run does:
# Create Report puts the WSFU take-off on a drafting view and a sheet and
# leaves the model alone, Size Water writes the sizes into the model.
#
# The pipeline itself lives in lib/water_run.py.
#
# IronPython 2.7 / pyRevit

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

import water_run

water_run.run(doc, uidoc, output, forms)
