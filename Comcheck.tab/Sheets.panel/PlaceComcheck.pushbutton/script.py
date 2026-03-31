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

# 3. Layout Settings
# WARNING: ALL MEASUREMENTS ARE IN FEET
# WARNING: ADJUST CELL_W, CELL_H, ORIGIN TO MATCH YOUR TITLEBLOCK
PAGES_PER_SHEET = 6
COLS = 3
ROWS = 2
SHEET_ORIGIN_X = 0.05    # close to left edge
SHEET_ORIGIN_Y = 1.85    # pushed higher up the sheet
CELL_W = 0.725
CELL_H = 0.95
GAP = 0.08               # more space between rows

# WARNING: CHANGE THESE SHEET NUMBERS TO MATCH YOUR COMPANY CONVENTION
SHEET_NUMBER_PREFIX = "M"
SHEET_NUMBER_START = 5
SHEET_NAME = "COMCHECK"

# 4. Find Titleblock
tb_collector = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_TitleBlocks)\
    .WhereElementIsElementType()
tb_types = list(tb_collector)
if not tb_types:
    forms.alert("No titleblock types found in project.", exitscript=True)

# WARNING: GRABS FIRST TITLEBLOCK - MAY NOT BE THE RIGHT ONE
tb_id = tb_types[0].Id

# 5. Calculate Sheet Count
num_sheets = (page_count + PAGES_PER_SHEET - 1) // PAGES_PER_SHEET

# 6. Create Sheets and Place Pages
with revit.Transaction("Place Comcheck PDF Pages"):
    for sheet_idx in range(num_sheets):

        sheet = ViewSheet.Create(doc, tb_id)

        sheet_number = "{}{}".format(
            SHEET_NUMBER_PREFIX,
            str(SHEET_NUMBER_START + sheet_idx).zfill(3)
        )
        sheet.SheetNumber = sheet_number
        sheet.Name = SHEET_NAME

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
    "Done! {} sheet(s) created with {} pages placed.".format(num_sheets, page_count),
    title="Comcheck Importer"
)