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
cols = {col.title: col.id for col in sheet.columns}

projects = []
current_project = None
last_known_owner = "UNASSIGNED"

OWNERS = ["NORMAN", "NORLENE"]
AC_PATTERN = re.compile(r'\bAC\d+', re.IGNORECASE)

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

    if not primary_val and not supplier_val:
        continue

    # Check for Owner
    detected_owner = None
    if supplier_val.upper() in OWNERS:
        detected_owner = supplier_val.upper()
    else:
        for o in OWNERS:
            if o in primary_val.upper():
                detected_owner = o
                break

    if detected_owner:
        last_known_owner = detected_owner

    has_ac = bool(AC_PATTERN.search(primary_val))
    is_header = detected_owner is not None or has_ac

    if is_header:
        # Save previous project block if populated
        if current_project and current_project["items"]:
            projects.append(current_project)
            current_project = None

        ac_matches = AC_PATTERN.findall(primary_val)
        ac_str = primary_val if has_ac else ""

        # Distinguish project name from job number
        proj_name = primary_val if not has_ac else ""
        if not proj_name and detected_owner and supplier_val.upper() in OWNERS:
            proj_name = primary_val

        current_project = {
            "owner": detected_owner or last_known_owner,
            "project_name": proj_name,
            "ac_number": ac_str,
            "items": []
        }
    elif current_project:
        # Fill missing project name/AC number from adjacent header rows
        if not current_project["project_name"] and not has_ac and primary_val:
            current_project["project_name"] = primary_val
            continue
        elif not current_project["ac_number"] and has_ac:
            current_project["ac_number"] = primary_val
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

# Post-processing cleanup for headers
for p in projects:
    if not p["project_name"] and p["ac_number"]:
        p["project_name"] = "Project " + p["ac_number"]
    elif not p["ac_number"]:
        p["ac_number"] = "NO AC"

with open("projects_data.json", "w") as f:
    json.dump(projects, f, indent=2)

print(f"Successfully processed {len(projects)} projects.")
