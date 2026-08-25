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

# Dynamic column mapping (handles any case/whitespace variation)
cols = {}
for col in sheet.columns:
    cols[col.title.strip().upper()] = col.id

def get_val(cell_map, col_name):
    cid = cols.get(col_name.upper())
    return str(cell_map.get(cid, "")).strip() if cid else ""

OWNERS = ["NORMAN", "NORLENE"]
# Strictly matches AC numbers like AC2942, AC 2942, or multi-AC like AC2904 & AC2905
AC_REGEX = re.compile(r'\bAC\s*\d+', re.IGNORECASE)

def determine_status_color(item):
    comments = str(item.get("comments", "")).upper()
    received = item.get("received", False)
    po_nr = str(item.get("po_nr", "")).strip()

    if "ON HOLD" in comments:
        return "red"
    if "QUOTE" in comments or "WAITING" in comments or "FEEDBACK" in comments:
        return "yellow"
    if "QUERY" in comments or "CHECK" in comments or item.get("followed_up", False):
        return "purple"
    if received:
        return "green"
    if po_nr:
        return "orange"
    return "white"

all_rows = list(sheet.rows)
projects = []

for idx, row in enumerate(all_rows):
    cell_values = {cell.column_id: (cell.value or cell.display_value or "") for cell in row.cells}
    
    # 1. Search ENTIRE row across all columns for AC pattern
    row_ac_match = None
    for val in cell_values.values():
        val_str = str(val).strip()
        match = AC_REGEX.search(val_str)
        if match:
            row_ac_match = val_str
            break

    # ANCHOR 1: FOUND AN AC ROW
    if row_ac_match:
        ac_number = row_ac_match

        # --- ANCHOR 2: SCAN UPWARD FOR PROJECT NAME & OWNER ---
        project_name = ""
        project_owner = "UNASSIGNED"

        up_idx = idx - 1
        while up_idx >= 0:
            up_row = all_rows[up_idx]
            up_cells = {cell.column_id: (cell.value or cell.display_value or "") for cell in up_row.cells}
            
            # Check if upward row contains another AC
            up_has_ac = any(AC_REGEX.search(str(v)) for v in up_cells.values())
            if up_has_ac:
                break

            # Search upward row for Owner
            for v in up_cells.values():
                val_upper = str(v).strip().upper()
                for o in OWNERS:
                    if o in val_upper:
                        project_owner = o
                        break

            # Get text from primary/first non-empty column in the upward header row
            up_texts = [str(v).strip() for v in up_cells.values() if str(v).strip()]
            if up_texts:
                first_text = up_texts[0]
                if first_text.upper() not in OWNERS:
                    project_name = first_text
                    break # Captured grey project name row

            up_idx -= 1

        if not project_name:
            project_name = ac_number

        # --- ANCHOR 3: SCAN DOWNWARD FOR COMPONENTS ---
        items = []
        down_idx = idx + 1

        while down_idx < len(all_rows):
            down_row = all_rows[down_idx]
            down_cells = {cell.column_id: (cell.value or cell.display_value or "") for cell in down_row.cells}
            
            # STOP: Hit next project's AC row
            if any(AC_REGEX.search(str(v)) for v in down_cells.values()):
                break

            # STOP: Hit next project's Grey Header row (No PO/Date, sits right above an AC row)
            down_texts = [str(v).strip() for v in down_cells.values() if str(v).strip()]
            has_po = bool(get_val(down_cells, "PO NR") or get_val(down_cells, "DATE OF PO"))
            
            if down_texts and not has_po:
                if down_idx + 1 < len(all_rows):
                    next_cells = {cell.column_id: (cell.value or cell.display_value or "") for cell in all_rows[down_idx + 1].cells}
                    if any(AC_REGEX.search(str(v)) for v in next_cells.values()):
                        break # Next row is AC, so this line is the next project header!

            # Grab primary component description
            primary_text = down_texts[0] if down_texts else ""
            if primary_text:
                item = {
                    "component": primary_text,
                    "supplier": get_val(down_cells, "SUPPLIER"),
                    "supplier_contact": get_val(down_cells, "SUPPLIER CONTACT"),
                    "po_nr": get_val(down_cells, "PO NR"),
                    "date_po": get_val(down_cells, "DATE OF PO"),
                    "received": bool(down_cells.get(cols.get("RECEIVED"), False)),
                    "date_received": get_val(down_cells, "DATE RECEIVED"),
                    "followed_up": bool(down_cells.get(cols.get("FOLLOWED UP"), False)),
                    "comments": get_val(down_cells, "FOLLOW UP COMMENTS / ETA")
                }
                item["status_color"] = determine_status_color(item)
                items.append(item)

            down_idx += 1

        projects.append({
            "owner": project_owner,
            "project_name": project_name,
            "ac_number": ac_number,
            "items": items
        })

with open("projects_data.json", "w") as f:
    json.dump(projects, f, indent=2)

print(f"Successfully processed {len(projects)} projects.")
