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
from pyrevit import forms, revit, script
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


# Per-sheet-size layout defaults. Advanced overrides in the form fall back to
# these values when left blank.
LAYOUT_DEFAULTS = {
    '24 x 36': dict(sheet_w=3.0, sheet_h=2.0, margin_left=0.20, margin_top=0.06, margin_right=0.39, margin_bottom=0.20, gap_col=0.05, gap_row=0.06, cols=4, rows=2),
    '30 x 42': dict(sheet_w=3.5, sheet_h=2.5, margin_left=0.00, margin_top=0.15, margin_right=0.65, margin_bottom=0.25, gap_col=0.06, gap_row=0.08, cols=4, rows=2),
}
DEFAULT_RESOLUTION = 600


# 1. User picks PDF
pdf_path = forms.pick_file(file_ext='pdf', title='Select Comcheck PDF')
if not pdf_path:
    script.exit()

with open(pdf_path, 'rb') as f:
    pdf_bytes = f.read()

# Page size used to fit each placed image within its grid cell without distortion
# or overlap into neighboring cells. Falls back to US Letter portrait (COMcheck default).
PAGE_W, PAGE_H = detect_pdf_page_size(pdf_bytes) or (8.5, 11.0)

# 2. Titleblock collection (needed to populate the form's combobox)
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

# 3. Consolidated form: sheet prefix/number, titleblock, sheet size, and
# optional advanced layout overrides.
#
# This pyRevit version's pyrevit.forms has no FlexForm/Label/TextBox/etc.
# Build the dialog directly on top of forms.WPFWindow with inline XAML.
COMCHECK_FORM_XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Comcheck Sheet Placement"
    Height="600" Width="440"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    ShowInTaskbar="False">
    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="12">
        <StackPanel>
            <Label Content="Sheet Prefix (e.g. M, E, P)"/>
            <TextBox Name="prefix_box" Text="M" Margin="0,0,0,8"/>

            <Label Content="Sheet Number (e.g. 04, 0.4, 005)"/>
            <TextBox Name="number_box" Text="005" Margin="0,0,0,8"/>

            <Label Content="Titleblock"/>
            <ComboBox Name="titleblock_combo" Margin="0,0,0,8"/>

            <Label Content="Sheet Size"/>
            <ComboBox Name="sheetsize_combo" Margin="0,0,0,8"/>

            <Expander Header="Advanced (optional overrides - leave blank to use defaults for the selected sheet size)" Margin="0,4,0,0">
                <StackPanel Margin="12,8,0,0">
                    <Label Content="Margin Left (ft)"/>
                    <TextBox Name="margin_left_box" Margin="0,0,0,6"/>
                    <Label Content="Margin Top (ft)"/>
                    <TextBox Name="margin_top_box" Margin="0,0,0,6"/>
                    <Label Content="Margin Right (ft)"/>
                    <TextBox Name="margin_right_box" Margin="0,0,0,6"/>
                    <Label Content="Margin Bottom (ft)"/>
                    <TextBox Name="margin_bottom_box" Margin="0,0,0,6"/>
                    <Label Content="Gap Col (ft)"/>
                    <TextBox Name="gap_col_box" Margin="0,0,0,6"/>
                    <Label Content="Gap Row (ft)"/>
                    <TextBox Name="gap_row_box" Margin="0,0,0,6"/>
                    <Label Content="Columns"/>
                    <TextBox Name="cols_box" Margin="0,0,0,6"/>
                    <Label Content="Rows"/>
                    <TextBox Name="rows_box" Margin="0,0,0,6"/>
                    <Label Content="Resolution (DPI)"/>
                    <TextBox Name="resolution_box" Margin="0,0,0,6"/>
                    <Label Content="Page Count Override"/>
                    <TextBox Name="page_count_box" Margin="0,0,0,6"/>
                </StackPanel>
            </Expander>

            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
                <Button Name="cancel_button" Content="Cancel" Width="80" Margin="0,0,8,0" Click="cancel_click"/>
                <Button Name="place_button" Content="Place Comcheck Sheets" Width="170" Click="place_click"/>
            </StackPanel>
        </StackPanel>
    </ScrollViewer>
</Window>
"""


class ComcheckPlacementForm(forms.WPFWindow):
    def __init__(self, xaml_source, titleblock_names):
        forms.WPFWindow.__init__(self, xaml_source, literal_string=True)
        self.titleblock_combo.ItemsSource = titleblock_names
        if titleblock_names:
            self.titleblock_combo.SelectedIndex = 0
        self.sheetsize_combo.ItemsSource = ['24 x 36', '30 x 42']
        self.sheetsize_combo.SelectedIndex = 0
        self.values = None

    def place_click(self, sender, args):
        self.values = {
            'sheet_prefix': self.prefix_box.Text,
            'sheet_number': self.number_box.Text,
            'titleblock': self.titleblock_combo.SelectedItem,
            'sheet_size': self.sheetsize_combo.SelectedItem,
            'margin_left': self.margin_left_box.Text,
            'margin_top': self.margin_top_box.Text,
            'margin_right': self.margin_right_box.Text,
            'margin_bottom': self.margin_bottom_box.Text,
            'gap_col': self.gap_col_box.Text,
            'gap_row': self.gap_row_box.Text,
            'cols': self.cols_box.Text,
            'rows': self.rows_box.Text,
            'resolution': self.resolution_box.Text,
            'page_count': self.page_count_box.Text,
        }
        self.Close()

    def cancel_click(self, sender, args):
        self.values = None
        self.Close()


form = ComcheckPlacementForm(COMCHECK_FORM_XAML, sorted(tb_dict.keys()))
form.ShowDialog()

if not form.values:
    script.exit()

values = form.values

# Sheet prefix
sheet_prefix = values.get('sheet_prefix', '')
if not sheet_prefix:
    script.exit()
sheet_prefix = sheet_prefix.upper().strip()

# Sheet number
sheet_number_input = values.get('sheet_number', '')
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

# Titleblock
selected_tb_name = values.get('titleblock')
if not selected_tb_name or selected_tb_name not in tb_dict:
    script.exit()

selected_tb = tb_dict[selected_tb_name]
tb_id = selected_tb.Id

# Sheet size
sheet_size = values.get('sheet_size')
if not sheet_size:
    script.exit()

defaults = LAYOUT_DEFAULTS[sheet_size]

def _override_float(key):
    raw = values.get(key)
    if raw:
        raw = raw.strip()
    if raw:
        return float(raw)
    return defaults[key]

def _override_int(key, default_value):
    raw = values.get(key)
    if raw:
        raw = raw.strip()
    if raw:
        return int(raw)
    return default_value

sheet_w       = defaults['sheet_w']
sheet_h       = defaults['sheet_h']
MARGIN_LEFT   = _override_float('margin_left')
MARGIN_TOP    = _override_float('margin_top')
MARGIN_RIGHT  = _override_float('margin_right')
MARGIN_BOTTOM = _override_float('margin_bottom')
GAP_COL       = _override_float('gap_col')
GAP_ROW       = _override_float('gap_row')
COLS          = _override_int('cols', defaults['cols'])
ROWS          = _override_int('rows', defaults['rows'])
RESOLUTION    = _override_int('resolution', DEFAULT_RESOLUTION)

PAGES_PER_SHEET = COLS * ROWS

# Page count: use override if provided, otherwise auto-detect
page_count_override = values.get('page_count')
if page_count_override:
    page_count_override = page_count_override.strip()
if page_count_override:
    page_count = int(page_count_override)
else:
    page_count = detect_pdf_page_count(pdf_bytes)
    if not page_count:
        forms.alert(
            "Could not auto-detect the page count from this PDF.\n"
            "Please enter a value in the 'Page Count Override' field and rerun.",
            exitscript=True
        )

# Auto calculate cell sizes
available_w = sheet_w - MARGIN_LEFT - MARGIN_RIGHT - (GAP_COL * (COLS - 1))
available_h = sheet_h - MARGIN_TOP - MARGIN_BOTTOM - (GAP_ROW * (ROWS - 1))

CELL_W = available_w / COLS
CELL_H = available_h / ROWS

SHEET_ORIGIN_X = MARGIN_LEFT
SHEET_ORIGIN_Y = sheet_h - MARGIN_TOP

# 4. Calculate Sheet Count
num_sheets = (page_count + PAGES_PER_SHEET - 1) // PAGES_PER_SHEET

# 5. Create Sheets and Place Pages
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
            img_opts.Resolution = RESOLUTION

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