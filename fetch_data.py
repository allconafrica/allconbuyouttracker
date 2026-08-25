import os
import re
import json
import smartsheet

TOKEN = os.environ.get("SMARTSHEET_TOKEN")
if not TOKEN:
    raise ValueError("SMARTSHEET_TOKEN secret is missing or empty.")

smart = smartsheet.Smartsheet(TOKEN)

target_sheet_name = "ALLCON PROJECTS BUY OUT LIST"
response = smart.Sheets.list_sheets(include_all=True)

sheet_id = None
for s in response.data:
    if s.name.strip().upper() == target_sheet_name.upper():
        sheet_id = s.id
        break

if not sheet_id and response.data:
    sheet_id = response.data[0].id

sheet = smart.Sheets.get_sheet(sheet_id)

# Map column names dynamically
cols = {col.title: col.id for col in sheet.columns}

projects = []
current_project = None
active_owner = "UNASSIGNED"

OWNERS = ["NORMAN", "NORLENE"]

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
    row_text = " ".join([str(v) for v in cell_values.values()]).upper()

    # Track owner if detected in any column of the row
    for o in OWNERS:
        if o in row_text:
            active_owner = o
            break

    primary_val = str(cell_values.get(cols.get("Primary Column"), "")).strip()
    supplier_val = str(cell_values.get(cols.get("SUPPLIER"), "")).strip()

    if not primary_val and not supplier_val:
        continue

    # Header Detection: Contains "AC" in Primary Column or Supplier Column
    is_ac_header = "AC" in primary_val.upper() or "AC" in supplier_val.upper()

    if is_ac_header:
        if current_project and current_project["items"]:
            projects.append(current_project)

        # Preserve exact full string without regex splitting
        raw_ac = primary_val if "AC" in primary_val.upper() else supplier_val
        raw_title = supplier_val if "AC" in primary_val.upper() else primary_val

        # Clean title if it matches owner name
        if raw_title.upper() in OWNERS:
            raw_title = ""

        current_project = {
            "owner": active_owner,
            "ac_number": raw_ac,
            "project_name": raw_title,
            "items": []
        }
    elif current_project:
        # If header had no project name, capture first descriptive text row
        if not current_project["project_name"] and primary_val and "AC" not in primary_val.upper():
            # Check if line looks like a project title rather than a component item
            if not cell_values.get(cols.get("PO NR"), "") and not cell_values.get(cols.get("Date of PO"), ""):
                current_project["project_name"] = primary_val
                continue

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

# Fallback cleanups
for p in projects:
    if not p["project_name"]:
        p["project_name"] = "Project " + p["ac_number"]

with open("projects_data.json", "w") as f:
    json.dump(projects, f, indent=2)

print(f"Successfully exported {len(projects)} grouped project blocks.")
