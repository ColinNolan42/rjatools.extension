# -*- coding: ascii -*-
# Water Report.pushbutton/script.py
# Produces the domestic water WSFU take-off as a drafting view on a sheet,
# per the 2024 IPC. Does not change the model.
#
# The whole pipeline lives in lib/water_run.py, shared with the Size Water
# button, so the two can never report different numbers for the same system.
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

water_run.run(doc, uidoc, output, forms, water_run.MODE_REPORT)
