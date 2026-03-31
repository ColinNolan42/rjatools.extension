# encoding: utf-8
# COMCHECK PDF PLACER - pyRevit Script
# Place Comcheck PDF pages on Revit sheets in a 3x2 grid

import os
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms, revit
import System
from System import Array

doc = revit.doc
uidoc = revit.uidoc

ito_type = clr.GetClrType(ImageTypeOptions)
ctor = ito_type.GetConstructor(
    Array[System.Type]([
        clr.GetClrType(System.String),
        clr.GetClrType(System.Boolean),
        clr.GetClrType(ImageTypeSource)
    ])
)

# 1. User picks PDF
pdf_path = forms.pick_file(file_ext='pdf', title='Select Comcheck PDF')
if not pdf_path:
    script.exit()

# 2. User enters total page count
page_count = forms.ask_for_string(
    prompt='How many pages is your Comcheck PDF?',
    title='Page Count',
    default='6'
)
if not page_count:
    script.exit()
page_count = int(page_count)

# 3. Ask for sheet number prefix
sheet_prefix = forms.ask_for_string(
    prompt='Enter sheet number prefix (e.g. M, E, P)',
    title='Sheet Prefix',
    default='M'
)
if not sheet_prefix:
    script.exit()

# 4. Ask for starting sheet number
sheet_start = forms.ask_for_string(
    prompt='Enter starting sheet number (e.g. 5 will create M005, M006...)',
    title='Starting Sheet Number',
    default='5'
)
if not sheet_start:
    script.exit()
sheet_start = int(sheet_start)

# 5. Ask for sheet name
sheet_name = forms.ask_for_string(
    prompt='Enter sheet name (e.g. COMCHECK, ENERGY COMPLIANCE)',
    title='Sheet Name',
    default='COMCHECK'
)
if not sheet_name:
    script.exit()

# 6. Titleblock picker
tb_collector = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_TitleBlocks)\
    .WhereElementIsElementType()
tb_types = list(tb_collector)
if not tb_types:
    forms.alert("No titleblock types found in project.", exitscript=True)

# Build a dictionary of name -> element for the picker
tb_dict = {}
for tb in tb_types:
    family_name = tb.Family.Name
    type_name = tb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    display_name = "{} : {}".format(family_name, type_name)
    tb_dict[display_name] = tb

# Show picker dialog
selected_tb_name = forms.SelectFromList.show(
    sorted(tb_dict.keys()),
    title='Select Titleblock',
    prompt='Choose a titleblock for the Comcheck sheets:',
    multiselect=False
)
if not selected_tb_name:
    script.exit()

tb_id = tb_dict[selected_tb_name].Id

# 7. Layout Settings
# WARNING: ALL MEASUREMENTS ARE IN FEET
# WARNING: ADJUST CELL_W, CELL_H, ORIGIN TO MATCH YOUR TITLEBLOCK
PAGES_PER_SHEET = 6
COLS = 3
ROWS = 2
SHEET_ORIGIN_X = 0.05
SHEET_ORIGIN_Y = 2.25
CELL_W = 0.725
CELL_H = 0.95
GAP = 0.08

# 8. Calculate Sheet Count
num_sheets = (page_count + PAGES_PER_SHEET - 1) // PAGES_PER_SHEET

# 9. Create Sheets and Place Pages
with revit.Transaction("Place Comcheck PDF Pages"):
    for sheet_idx in range(num_sheets):

        sheet = ViewSheet.Create(doc, tb_id)

        sheet_number = "{}{}".format(
            sheet_prefix,
            str(sheet_start + sheet_idx).zfill(3)
        )
        sheet.SheetNumber = sheet_number
        sheet.Name = sheet_name

        comments_param = sheet.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if comments_param:
            comments_param.Set("MECHANICAL")

        start_page = sheet_idx * PAGES_PER_SHEET
        end_page = min(start_page + PAGES_PER_SHEET, page_count)

        for i, page_num in enumerate(range(start_page, end_page)):
            col = i % COLS
            row = i // COLS

            x = SHEET_ORIGIN_X + col * (CELL_W + GAP)
            y = SHEET_ORIGIN_Y - row * (CELL_H + GAP)
            origin = XYZ(x, y, 0)

            img_opts = ctor.Invoke(
                Array[System.Object]([pdf_path, False, ImageTypeSource.Import])
            )
            img_opts.PageNumber = page_num + 1
            img_opts.Resolution = 150

            img_type = ImageType.Create(doc, img_opts)

            place_opts = ImagePlacementOptions()
            place_opts.PlacementPoint = BoxPlacement.TopLeft
            place_opts.Location = origin

            ImageInstance.Create(doc, sheet, img_type.Id, place_opts)

forms.alert(
    "Done! {} sheet(s) created: {}{} to {}{}".format(
        num_sheets,
        sheet_prefix, str(sheet_start).zfill(3),
        sheet_prefix, str(sheet_start + num_sheets - 1).zfill(3)
    ),
    title="Comcheck Importer"
)