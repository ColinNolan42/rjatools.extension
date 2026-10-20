# -*- coding: ascii -*-
# Water Performance.pushbutton/script.py
# Colors domestic water pipes in a copy of the active plan view by drawn size
# against the 2024 IPC firm sizing table. Read only apart from the copied view.
#
# The pipeline lives in lib/water_perf_run.py.
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

import water_perf_run

water_perf_run.run(doc, uidoc, output, forms)
