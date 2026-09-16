"""Shared JSON schema for the extracted BOM & cost estimation payload.

This is the single contract between the AI extraction step, the frontend
editor, and the Excel export — every part of the app reads/writes this shape.
"""

EMPTY_PAYLOAD = {
    "project": {
        "projectName": "",
        "documentRef": "",
        "client": "",
        "manufacturer": "",
        "date": "",
        "currency": "SAR",
        "vatRate": 0.15,
        "notes": "",
    },
    "items": [],
    "bom": [],
    "hardware": [],
    "assumptions": [],
}

# Column order used by both the frontend table renderer and the Excel export,
# so the two never drift apart.
ITEM_COLUMNS = [
    ("itemNo", "Item #"),
    ("drawingRef", "Drawing / Item Ref"),
    ("description", "Description & Specification"),
    ("dimensions", "Dimensions (mm)"),
    ("unit", "Unit"),
    ("qty", "Qty"),
    ("unitCost", "Unit Cost"),
]

BOM_COLUMNS = [
    ("itemNo", "Part #"),
    ("parentRef", "Parent Item Ref"),
    ("assembly", "Assembly / Sub-System"),
    ("partDescription", "Part Description"),
    ("materialSpec", "Material & Substrate Spec"),
    ("dimensions", "Finished Dimensions (mm)"),
    ("qtyPerUnit", "Qty / Unit"),
    ("totalQty", "Total Qty"),
    ("unit", "Unit"),
    ("notes", "Notes"),
]

HARDWARE_COLUMNS = [
    ("code", "Item Code"),
    ("category", "Category"),
    ("itemName", "Item Name / Model"),
    ("brand", "Brand / Manufacturer"),
    ("material", "Material & Finish"),
    ("applicable", "Applicable Items"),
    ("qtyPerUnit", "Qty / Unit"),
    ("totalQty", "Total Qty"),
    ("unit", "Unit"),
    ("unitCost", "Unit Cost"),
    ("scope", "Scope / Function"),
]
