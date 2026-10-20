# -*- coding: ascii -*-
# Size Duct.pushbutton/script.py
# Sizes every main duct to the RJA velocity and friction limits and writes
# the sizes into the model. Branches, takeoff boots and diffusers are never
# touched.
#
# The pipeline lives in lib/duct_size_run.py, the sizing rules in
# lib/duct_sizing.py, and the traversal is the Duct Performance engine
# (lib/hvac_graph.py).
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

import duct_size_run

duct_size_run.run(doc, uidoc, output, forms)
