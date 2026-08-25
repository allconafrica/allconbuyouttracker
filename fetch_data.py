import os
import re
import json
import smartsheet

# 1. Load secret token from GitHub Actions environment
TOKEN = os.environ.get("SMARTSHEET_TOKEN")

if not TOKEN:
    raise ValueError("SMARTSHEET_TOKEN secret is missing or empty.")

smart = smartsheet.Smartsheet(TOKEN)

# 2. Dynamically find the sheet ID by its exact name
target_sheet_name = "ALLCON PROJECTS BUY OUT LIST"
response = smart.Sheets.list_sheets(include_all=True)

sheet_id = None
for s in response.data:
    if s.name.strip().upper() == target_sheet_name.upper():
        sheet_id = s.id
        break

# Fallback: take the first sheet available if name matching fails
if not sheet_id and response.data:
    sheet_id = response.data[0].id

if not sheet_id:
    raise ValueError("No sheets found for this API token. Verify token access in Smartsheet.")

sheet = smart.Sheets.get_sheet(sheet_id)

# 3. Map sheet columns dynamically
cols = {col.title: col.id for col in sheet.columns}

projects = []
current_project = None

OWNERS = ["NORMAN", "NORLENE"]
AC_PATTERN = re.compile(r'\bAC\d+\b', re.IGNORECASE)

# 4. Color status evaluation function
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

# 5. Process rows sequentially
for row in sheet.rows:
    cell_values = {cell.column_id: (cell.value or cell.display_value or "") for cell in row.cells}

    primary_val = str(cell_values.get(cols.get("Primary Column"), "")).strip()
    supplier_val = str(cell_values.get(cols.get("SUPPLIER"), "")).strip()

    is_owner_row = supplier_val.upper() in OWNERS or any(o in primary_val.upper() for o in OWNERS)
    has_ac = bool(AC_PATTERN.search(primary_val))

    # Detect project headers
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

    # Detect component rows under current project
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

# Save final project block
if current_project and current_project["items"]:
    projects.append(current_project)

# Export processed JSON for GitHub Pages frontend
with open("projects_data.json", "w") as f:
    json.dump(projects, f, indent=2)

print(f"Successfully exported {len(projects)} projects to projects_data.json")
