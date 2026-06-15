# encoding: utf-8
# COMCHECK PDF PLACER - pyRevit Script
# Place Comcheck PDF pages on Revit sheets in a 4x2 grid

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

# Check Revit version
app = revit.doc.Application
revit_version = int(app.VersionNumber)

if revit_version < 2021:
    forms.alert(
        "This script requires Revit 2021 or newer.\nYou are running Revit {}.".format(revit_version),
        exitscript=True
    )

ito_type = clr.GetClrType(ImageTypeOptions)

if revit_version >= 2021:
    ctor = ito_type.GetConstructor(
        Array[System.Type]([
            clr.GetClrType(System.String),
            clr.GetClrType(System.Boolean),
            clr.GetClrType(ImageTypeSource)
        ])
    )
else:
    ctor = ito_type.GetConstructor(
        Array[System.Type]([
            clr.GetClrType(System.String)
        ])
    )

def detect_pdf_page_count(pdf_bytes):
    """Best-effort PDF page count from raw bytes (no external libraries).

    Tries the /Type /Pages object's /Count value first (most reliable for the
    root page tree), then falls back to counting /Type /Page object headers.
    Returns None if neither approach finds anything.
    """
    counts = [int(c) for c in re.findall(r'/Type\s*/Pages.*?/Count\s+(\d+)', pdf_bytes, re.DOTALL)]
    if counts:
        return max(counts)

    page_objs = re.findall(r'/Type\s*/Page(?!s)', pdf_bytes)
    if page_objs:
        return len(page_objs)

    return None


def detect_pdf_page_size(pdf_bytes):
    """Best-effort PDF page size (width, height) from the first /MediaBox found.

    Returns None if no MediaBox is found.
    """
    match = re.search(r'/MediaBox\s*\[\s*([\d.\-]+)\s+([\d.\-]+)\s+([\d.\-]+)\s+([\d.\-]+)\s*\]', pdf_bytes)
    if not match:
        return None

    x0, y0, x1, y1 = (float(v) for v in match.groups())
    width, height = abs(x1 - x0), abs(y1 - y0)
    if width <= 0 or height <= 0:
        return None
    return (width, height)


def fit_dimensions(cell_w, cell_h, content_w, content_h):
    """Scale (content_w, content_h) to fit within (cell_w, cell_h), preserving aspect ratio."""
    scale = min(cell_w / content_w, cell_h / content_h)
    return content_w * scale, content_h * scale


# 1. User picks PDF
pdf_path = forms.pick_file(file_ext='pdf', title='Select Comcheck PDF')
if not pdf_path:
    script.exit()

with open(pdf_path, 'rb') as f:
    pdf_bytes = f.read()

# Page size used to fit each placed image within its grid cell without distortion
# or overlap into neighboring cells. Falls back to US Letter portrait (COMcheck default).
PAGE_W, PAGE_H = detect_pdf_page_size(pdf_bytes) or (8.5, 11.0)

# 2. Auto-detect total page count (falls back to manual entry if detection fails)
page_count = detect_pdf_page_count(pdf_bytes)
if not page_count:
    page_count = forms.ask_for_string(
        prompt='Could not auto-detect page count. How many pages is your Comcheck PDF?',
        title='Page Count',
        default='8'
    )
    if not page_count:
        script.exit()
    page_count = int(page_count)

# 3. Ask for sheet prefix
sheet_prefix = forms.ask_for_string(
    prompt='Enter sheet prefix (e.g. M, E, P)',
    title='Sheet Prefix',
    default='M'
)
if not sheet_prefix:
    script.exit()
sheet_prefix = sheet_prefix.upper().strip()

# 4. Ask for sheet number
# Handles any format - M0.4, M04, M005, M4 etc
# Only the LAST number in the sequence will increment
sheet_number_input = forms.ask_for_string(
    prompt='Enter sheet number (e.g. 04, 0.4, 005)\nOnly the last number will increment for additional sheets.',
    title='Sheet Number',
    default='005'
)
if not sheet_number_input:
    script.exit()
sheet_number_input = sheet_number_input.strip()

# Parse the last number in the input so we can increment it
# Examples: 0.4 -> last number is 4, 04 -> last number is 4, 005 -> last number is 5
last_num_match = re.search(r'(\d+)$', sheet_number_input)
if not last_num_match:
    forms.alert("Could not parse sheet number. Please enter a number like 04, 0.4, or 005", exitscript=True)

last_num_str   = last_num_match.group(1)          # e.g. "4" from "0.4"
last_num_int   = int(last_num_str)                 # e.g. 4
last_num_pad   = len(last_num_str)                 # e.g. 1
prefix_part    = sheet_number_input[:last_num_match.start()]  # e.g. "0." from "0.4"

def make_sheet_number(idx):
    # Increment only the last number, preserve everything before it
    new_num = str(last_num_int + idx).zfill(last_num_pad)
    return "{}{}{}".format(sheet_prefix, prefix_part, new_num)

# Sheet name always defaults to COMCHECK
sheet_name = "COMCHECK"

# 5. Titleblock picker
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

# 6. User picks sheet size
sheet_size = forms.SelectFromList.show(
    ['24 x 36', '30 x 42'],
    title='Select Sheet Size',
    prompt='Choose your sheet size:',
    multiselect=False
)
if not sheet_size:
    script.exit()

PAGES_PER_SHEET = 8
COLS = 4
ROWS = 2

if sheet_size == '24 x 36':
    sheet_w       = 3.0
    sheet_h       = 2.0
    MARGIN_LEFT   = 0.18
    MARGIN_TOP    = 0.10
    MARGIN_RIGHT  = 0.70
    MARGIN_BOTTOM = 0.20
    GAP_COL       = 0.05
    GAP_ROW       = 0.06
else:
    sheet_w       = 3.5
    sheet_h       = 2.5
    MARGIN_LEFT   = 0.02
    MARGIN_TOP    = 0.15
    MARGIN_RIGHT  = 0.75
    MARGIN_BOTTOM = 0.25
    GAP_COL       = 0.06
    GAP_ROW       = 0.08

# Auto calculate cell sizes
available_w = sheet_w - MARGIN_LEFT - MARGIN_RIGHT - (GAP_COL * (COLS - 1))
available_h = sheet_h - MARGIN_TOP - MARGIN_BOTTOM - (GAP_ROW * (ROWS - 1))

CELL_W = available_w / COLS
CELL_H = available_h / ROWS

SHEET_ORIGIN_X = MARGIN_LEFT
SHEET_ORIGIN_Y = sheet_h - MARGIN_TOP

# 7. Calculate Sheet Count
num_sheets = (page_count + PAGES_PER_SHEET - 1) // PAGES_PER_SHEET

# 8. Create Sheets and Place Pages
with revit.Transaction("Place Comcheck PDF Pages"):
    for sheet_idx in range(num_sheets):

        sheet = ViewSheet.Create(doc, tb_id)

        sheet.SheetNumber = make_sheet_number(sheet_idx)
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

            if revit_version >= 2021:
                img_opts = ctor.Invoke(
                    Array[System.Object]([pdf_path, False, ImageTypeSource.Import])
                )
            else:
                img_opts = ctor.Invoke(
                    Array[System.Object]([pdf_path])
                )

            img_opts.PageNumber = page_num + 1
            img_opts.Resolution = 150

            img_type = ImageType.Create(doc, img_opts)

            place_opts = ImagePlacementOptions()
            place_opts.PlacementPoint = BoxPlacement.TopLeft
            place_opts.Location = origin

            img_instance = ImageInstance.Create(doc, sheet, img_type.Id, place_opts)
            img_w, img_h = fit_dimensions(CELL_W, CELL_H, PAGE_W, PAGE_H)
            img_instance.Width = img_w
            img_instance.Height = img_h

forms.alert(
    "Done! {} sheet(s) created: {} to {}\nSheet size: {}".format(
        num_sheets,
        make_sheet_number(0),
        make_sheet_number(num_sheets - 1),
        sheet_size
    ),
    title="Comcheck Importer"
)