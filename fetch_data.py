import os
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
pending_project_name = ""

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

    primary_val = str(cell_values.get(cols.get("Primary Column"), "")).strip()
    supplier_val = str(cell_values.get(cols.get("SUPPLIER"), "")).strip()

    if not primary_val and not supplier_val:
        continue

    # Detect Owner
    detected_owner = "UNASSIGNED"
    if supplier_val.upper() in OWNERS:
        detected_owner = supplier_val.upper()
    elif primary_val.upper() in OWNERS:
        detected_owner = primary_val.upper()

    # Case 1: Row above AC containing Project Name (e.g. MARMATO - SM TROMMEL FRAME)
    if detected_owner != "UNASSIGNED" or (primary_val and "AC" not in primary_val.upper() and not supplier_val and not cell_values.get(cols.get("PO NR"), "")):
        pending_project_name = primary_val
        if detected_owner != "UNASSIGNED":
            active_owner = detected_owner
        else:
            active_owner = "UNASSIGNED"
        continue

    # Case 2: Row containing AC Number (e.g. AC2942)
    if "AC" in primary_val.upper() or "AC" in supplier_val.upper():
        if current_project and current_project["items"]:
            projects.append(current_project)

        ac_str = primary_val if "AC" in primary_val.upper() else supplier_val

        current_project = {
            "owner": active_owner if 'active_owner' in locals() else "UNASSIGNED",
            "project_name": pending_project_name if pending_project_name else ac_str,
            "ac_number": ac_str,
            "items": []
        }
        pending_project_name = "" # Reset for next project
        continue

    # Case 3: Standard Component Rows (e.g. PLATE MATERIAL, STRUCTURAL MATERIAL)
    if current_project:
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

# Output preserved sheet order directly
with open("projects_data.json", "w") as f:
    json.dump(projects, f, indent=2)

print(f"Exported {len(projects)} projects in original Smartsheet order.")
