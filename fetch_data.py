import os
import re
import json
import smartsheet

# Load token from GitHub Secret environment variable
TOKEN = os.environ.get("SMARTSHEET_TOKEN")
SHEET_ID = 4059040942542724  # Extracted from your sheet URL: 9mRVF6gf9Vr4Mqhx6hm3wvh6pj6FCqcH2G29Rg41

smart = smartsheet.Smartsheet(TOKEN)
sheet = smart.Sheets.get_sheet(SHEET_ID)

cols = {col.title: col.id for col in sheet.columns}

projects = []
current_project = None

OWNERS = ["NORMAN", "NORLENE"]
AC_PATTERN = re.compile(r'\bAC\d+\b', re.IGNORECASE)

def determine_status_color(item):
    comments = str(item.get("comments", "")).upper()
    received = item.get("received", False)
    po_nr = str(item.get("po_nr", "")).strip()

    if "ON HOLD" in comments:
        return "red"
    if "QUERY" in comments or "CHECK" in comments or item.get("followed_up", False):
        return "purple"
    if received:
        return "green"
    if po_nr:
        return "orange"
    return "white"

for row in sheet.rows:
    cell_values = {cell.column_id: (cell.value or cell.display_value or "") for cell in row.cells}

    primary_val = str(cell_values.get(cols.get("Primary Column"), "")).strip()
    supplier_val = str(cell_values.get(cols.get("SUPPLIER"), "")).strip()

    is_owner_row = supplier_val.upper() in OWNERS or any(o in primary_val.upper() for o in OWNERS)
    has_ac = bool(AC_PATTERN.search(primary_val))

    if is_owner_row or has_ac:
        if current_project and current_project["items"]:
            projects.append(current_project)
            current_project = None

        owner = "UNASSIGNED"
        if supplier_val.upper() in OWNERS:
            owner = supplier_val.upper()
        elif "NORMAN" in primary_val.upper():
            owner = "NORMAN"
        elif "NORLENE" in primary_val.upper():
            owner = "NORLENE"

        ac_match = AC_PATTERN.search(primary_val)
        ac_num = ac_match.group(0) if ac_match else ""

        current_project = {
            "owner": owner,
            "project_name": primary_val if not has_ac else "",
            "ac_number": ac_num,
            "items": []
        }
    elif current_project and primary_val:
        item = {
            "component": primary_val,
            "supplier": supplier_val,
            "supplier_contact": cell_values.get(cols.get("Supplier Contact"), ""),
            "po_nr": cell_values.get(cols.get("PO NR"), ""),
            "date_po": cell_values.get(cols.get("Date of PO"), ""),
            "received": bool(cell_values.get(cols.get("Received"), False)),
            "date_received": cell_values.get(cols.get("Date Received"), ""),
            "followed_up": bool(cell_values.get(cols.get("Followed up"), False)),
            "comments": cell_values.get(cols.get("Follow up comments / ETA"), "")
        }
        item["status_color"] = determine_status_color(item)
        current_project["items"].append(item)

if current_project and current_project["items"]:
    projects.append(current_project)

with open("projects_data.json", "w") as f:
    json.dump(projects, f, indent=2)
