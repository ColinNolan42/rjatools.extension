# -*- coding: utf-8 -*-
__title__   = "Grid\nAPI Probe"
__author__  = "MEP Tools"
__version__ = "probe"
__doc__     = "Prints all Grid API methods relevant to bubble separation."

from Autodesk.Revit.DB import FilteredElementCollector, Grid
from pyrevit import script, revit, forms

doc    = revit.doc
output = script.get_output()

grids = FilteredElementCollector(doc).OfClass(Grid).ToElements()
if not grids:
    forms.alert("No grids found in document.")
    script.exit()

g = grids[0]
output.print_md("## Grid: {} (ID {})".format(g.Name, g.Id.IntegerValue))

keywords = [
    'elbow', 'break', 'kink', 'bend', 'split',
    'bubble', 'head', 'datum', 'extent', 'curve',
    'end', 'offset', 'view', 'leader',
]

output.print_md("## Filtered methods (elbow / break / bubble / curve / extent)")
for name in sorted(dir(g)):
    if name.startswith('_'):
        continue
    low = name.lower()
    if any(k in low for k in keywords):
        output.print_md("- `{}`".format(name))

output.print_md("## All public methods and properties")
for name in sorted(dir(g)):
    if not name.startswith('_'):
        output.print_md("- `{}`".format(name))