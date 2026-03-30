# ============================================================
# COMCHECK PDF PLACER — pyRevit Script
# Place Comcheck PDF pages on Revit sheets in a 3x2 grid
# ============================================================

import os
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms, revit

doc   = revit.doc
uidoc = revit.uidoc

# ── 1. User picks PDF ─────────────────────────────────────────────────────────
pdf_path = forms.pick_file(file_ext='pdf', title='Select Comcheck PDF')
if not pdf_path:
    script.exit()

# ── 2. User enters total page count ──────────────────────────────────────────
page_count = forms.ask_for_string(
    prompt='How many pages is your Comcheck PDF?',
    title='Page Count',
    default='6'
)
if not page_count:
    script.exit()
page_count = int(page_count)

# ── 3. Layout Settings ────────────────────────────────────────────────────────
# ⚠️ ALL MEASUREMENTS ARE IN FEET — Revit always uses feet internally
# ⚠️ THESE WILL LIKELY NEED ADJUSTMENT to match your titleblock size
# ⚠️ If pages are overlapping or misaligned, tweak CELL_W, CELL_H, and ORIGIN values

PAGES_PER_SHEET = 6       # 3 columns x 2 rows
COLS            = 3
ROWS            = 2

# ⚠️ SHEET_ORIGIN = where the TOP-LEFT page starts on the sheet
# ⚠️ Depends entirely on your titleblock layout — may need significant adjustment
SHEET_ORIGIN_X  = 0.5    # ~6 inches from left edge
SHEET_ORIGIN_Y  = 0.4    # ~4.8 inches from bottom edge

# ⚠️ CELL SIZE = how much space each PDF page gets on the sheet
# ⚠️ Based on a 24x36 sheet with 6 pages — adjust if your sheet size is different
CELL_W          = 0.725  # ~8.7 inches wide per page
CELL_H          = 0.95   # ~11.4 inches tall per page

GAP             = 0.04   # ~0.5 inch gap between pages

# ── 4. Find Titleblock ────────────────────────────────────────────────────────
tb_collector = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_TitleBlocks)\
    .WhereElementIsElementType()
tb_types = list(tb_collector)
if not tb_types:
    forms.alert("No titleblock types found in project.", exitscript=True)

# ⚠️ THIS GRABS THE FIRST TITLEBLOCK IN THE PROJECT
# ⚠️ If your project has multiple titleblock types (24x36, 8.5x11, etc.)
# ⚠️ this may grab the wrong one — we should add a picker here eventually
tb_id = tb_types[0].Id

# ── 5. Calculate Sheet Count ──────────────────────────────────────────────────
num_sheets = (page_count + PAGES_PER_SHEET - 1) // PAGES_PER_SHEET

# ── 6. Create Sheets and Place Pages ─────────────────────────────────────────
with revit.Transaction("Place Comcheck PDF Pages"):
    for sheet_idx in range(num_sheets):

        # Create the sheet
        sheet = ViewSheet.Create(doc, tb_id)

        # ⚠️ SHEET NUMBER FORMAT — change "COMcheck-1" to match your
        # ⚠️ company's sheet numbering convention (e.g. "E0.1", "M-001", etc.)
        sheet.SheetNumber = "COMcheck-{}".format(sheet_idx + 1)
        sheet.Name = "COMcheck Energy Compliance ({} of {})".format(
            sheet_idx + 1, num_sheets)

        # Pages for this sheet
        start_page = sheet_idx * PAGES_PER_SHEET
        end_page   = min(start_page + PAGES_PER_SHEET, page_count)

        for i, page_num in enumerate(range(start_page, end_page)):
            col = i % COLS
            row = i // COLS

            x = SHEET_ORIGIN_X + col * (CELL_W + GAP)
            # Revit Y goes bottom-up so we flip the row order
            y = SHEET_ORIGIN_Y + (ROWS - 1 - row) * (CELL_H + GAP)
            origin = XYZ(x, y, 0)

            # ⚠️ NOT 100% SURE — PdfImportOptions may not exist in Revit 2021 or earlier
            # ⚠️ Confirmed to exist in Revit 2022+ but exact class name needs verification
            opts = PdfImportOptions()

            # ⚠️ NOT 100% SURE — PageNumber may be 0-based not 1-based
            # ⚠️ If pages are offset by 1 (wrong pages importing), change to: page_num
            opts.PageNumber = page_num + 1

            # ⚠️ NOT 100% SURE — PdfResolution.MediumDefinition is my best guess
            # ⚠️ at the enum name. Other possible values may be:
            # ⚠️ PdfResolution.Low, PdfResolution.Medium, PdfResolution.High
            opts.Resolution = PdfResolution.MediumDefinition

            # ⚠️ NOT 100% SURE — doc.Import() signature for PDFs on sheets
            # ⚠️ The last argument (origin) may need to be passed differently
            # ⚠️ Some versions of the API use a separate Move() call after import
            # ⚠️ If pages all land in the same spot, this is why
            doc.Import(pdf_path, opts, sheet, origin)

forms.alert(
    "Done! {} sheet(s) created with {} pages placed.".format(num_sheets, page_count),
    title="Comcheck Importer"
)