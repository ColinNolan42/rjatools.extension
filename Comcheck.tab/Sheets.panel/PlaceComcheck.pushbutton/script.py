# encoding: utf-8
# COMCHECK PDF PLACER - pyRevit Script
# Place Comcheck PDF pages on Revit sheets in a 3x2 grid

import os
import re
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

# 3. Ask for full sheet number and name in one box
first_sheet_full = forms.ask_for_string(
    prompt='Enter full sheet number and name (e.g. M005 - COMCHECK)\nAdditional sheets will auto increment the number.',
    title='Sheet Number and Name',
    default='M005 - COMCHECK'
)
if not first_sheet_full:
    script.exit()

match = re.match(r'^([A-Za-z]+)(\d+)\s*[-\s]?\s*(.+)$', first_sheet_full.strip())
if not match:
    forms.alert(
        "Could not parse sheet number. Please use format: M005 - COMCHECK",
        exitscript=True
    )

sheet_prefix  = match.group(1).upper()
sheet_start   = int(match.group(2))
sheet_name    = match.group(3).strip().upper()
zero_pad      = len(match.group(2))

# 4. Titleblock picker
tb_collector = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_TitleBlocks)\
    .WhereElementIsElementType()
tb_types = list(tb_collector)
if not tb_types:
    forms.alert("No titleblock types found in project.", exitscript=True)

tb_dict = {}
for tb in tb_types:
    family_name = tb.Family.Name
    type_name = tb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    display_name = "{} : {}".format(family_name, type_name)
    tb_dict[display_name] = tb

selected_tb_name = forms.SelectFromList.show(
    sorted(tb_dict.keys()),
    title='Select Titleblock',
    prompt='Choose a titleblock for the Comcheck sheets:',
    multiselect=False
)
if not selected_tb_name:
    script.exit()

selected_tb = tb_dict[selected_tb_name]
tb_id = selected_tb.Id

# 5. User picks sheet size
sheet_size = forms.SelectFromList.show(
    ['24 x 36', '30 x 42'],
    title='Select Sheet Size',
    prompt='Choose your sheet size:',
    multiselect=False
)
if not sheet_size:
    script.exit()

PAGES_PER_SHEET = 6
COLS = 3
ROWS = 2

if sheet_size == '24 x 36':
    sheet_w       = 3.0
    sheet_h       = 2.0
    MARGIN_LEFT   = 0.5     # shifted more right
    MARGIN_TOP    = 0.10
    MARGIN_RIGHT  = 0.70
    MARGIN_BOTTOM = 0.20
    GAP_COL       = 0.04
    GAP_ROW       = 0.06
else:
    # 30 x 42
    sheet_w       = 3.5
    sheet_h       = 2.5
    MARGIN_LEFT   = 0.03
    MARGIN_TOP    = 0.15
    MARGIN_RIGHT  = 0.85
    MARGIN_BOTTOM = 0.25
    GAP_COL       = 0.05
    GAP_ROW       = 0.08

# Auto calculate cell sizes
available_w = sheet_w - MARGIN_LEFT - MARGIN_RIGHT - (GAP_COL * (COLS - 1))
available_h = sheet_h - MARGIN_TOP - MARGIN_BOTTOM - (GAP_ROW * (ROWS - 1))

CELL_W = available_w / COLS
CELL_H = available_h / ROWS

SHEET_ORIGIN_X = MARGIN_LEFT
SHEET_ORIGIN_Y = sheet_h - MARGIN_TOP

# 6. Calculate Sheet Count
num_sheets = (page_count + PAGES_PER_SHEET - 1) // PAGES_PER_SHEET

# 7. Create Sheets and Place Pages
with revit.Transaction("Place Comcheck PDF Pages"):
    for sheet_idx in range(num_sheets):

        sheet = ViewSheet.Create(doc, tb_id)

        sheet_number = "{}{}".format(
            sheet_prefix,
            str(sheet_start + sheet_idx).zfill(zero_pad)
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

            x = SHEET_ORIGIN_X + col * (CELL_W + GAP_COL)
            y = SHEET_ORIGIN_Y - row * (CELL_H + GAP_ROW)
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

            img_instance = ImageInstance.Create(doc, sheet, img_type.Id, place_opts)
            img_instance.Width = CELL_W
            img_instance.Height = CELL_H

forms.alert(
    "Done! {} sheet(s) created: {}{} to {}{}\nSheet size: {}".format(
        num_sheets,
        sheet_prefix, str(sheet_start).zfill(zero_pad),
        sheet_prefix, str(sheet_start + num_sheets - 1).zfill(zero_pad),
        sheet_size
    ),
    title="Comcheck Importer"
)