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

# Request sheet WITH format data included
sheet = smart.Sheets.get_sheet(sheet_id, include="format")
cols = {col.title: col.id for col in sheet.columns}

OWNERS = ["NORMAN", "NORLENE"]
AC_REGEX = re.compile(r'\bAC\s*\d+', re.IGNORECASE)

def is_colored_row(row):
    """Checks if any primary or supplier cell has a background color formatting applied."""
    for cell in row.cells:
        if cell.format:
            # Smartsheet format strings contain hex background colors
            # Colored header rows (Grey, Brown, Blue, Pink) will have specific format flags
            return True
    return False

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

all_rows = list(sheet.rows)
projects = []

for idx, row in enumerate(all_rows):
    cell_values = {cell.column_id: (cell.value or cell.display_value or "") for cell in row.cells}
    primary_val = str(cell_values.get(cols.get("Primary Column"), "")).strip()
    supplier_val = str(cell_values.get(cols.get("SUPPLIER"), "")).strip()

    # ANCHOR 1: Detect the Brown AC Row (Matches AC pattern or has AC text in primary/supplier)
    if AC_REGEX.search(primary_val) or AC_REGEX.search(supplier_val):
        ac_number = primary_val if AC_REGEX.search(primary_val) else supplier_val

        # --- ANCHOR 2: SCAN UPWARD FOR GREY (PROJECT NAME) & BLUE/PINK (OWNER) ---
        project_name = ""
        project_owner = "UNASSIGNED"

        up_idx = idx - 1
        while up_idx >= 0:
            up_row = all_rows[up_idx]
            up_cells = {cell.column_id: (cell.value or cell.display_value or "") for cell in up_row.cells}
            up_primary = str(up_cells.get(cols.get("Primary Column"), "")).strip()
            up_supplier = str(up_cells.get(cols.get("SUPPLIER"), "")).strip()

            # Stop if we bump into the previous project's AC row
            if AC_REGEX.search(up_primary) or AC_REGEX.search(up_supplier):
                break

            # Look for Owner in the Blue/Pink row cells
            for o in OWNERS:
                if o in up_supplier.upper() or o in up_primary.upper():
                    project_owner = o
                    break

            # The Grey Row sitting above AC is our Project Name
            if up_primary and up_primary.upper() not in OWNERS:
                project_name = up_primary
                break  # Exit upward scan once captured

            up_idx -= 1

        if not project_name:
            project_name = ac_number

        # --- ANCHOR 3: SCAN DOWNWARD FOR ACTION ITEMS (UNCOLORED ROWS) ---
        items = []
        down_idx = idx + 1

        while down_idx < len(all_rows):
            down_row = all_rows[down_idx]
            down_cells = {cell.column_id: (cell.value or cell.display_value or "") for cell in down_row.cells}
            down_primary = str(down_cells.get(cols.get("Primary Column"), "")).strip()
            down_supplier = str(down_cells.get(cols.get("SUPPLIER"), "")).strip()

            # STOP CONDITION: We hit the next project's Grey Header Row or Brown AC Row
            if AC_REGEX.search(down_primary) or AC_REGEX.search(down_supplier):
                break

            # If down_primary exists and we hit another header block text (Grey Row), stop
            if down_primary and not down_cells.get(cols.get("PO NR"), "") and not down_cells.get(cols.get("Date of PO"), "") and not down_supplier:
                # Check if the row directly after it is an AC row
                if down_idx + 1 < len(all_rows):
                    next_cells = {cell.column_id: (cell.value or cell.display_value or "") for cell in all_rows[down_idx + 1].cells}
                    next_p = str(next_cells.get(cols.get("Primary Column"), "")).strip()
                    next_s = str(next_cells.get(cols.get("SUPPLIER"), "")).strip()
                    if AC_REGEX.search(next_p) or AC_REGEX.search(next_s):
                        break

            # Line Item (White / Yellow highlighted action item rows)
            if down_primary:
                item = {
                    "component": down_primary,
                    "supplier": down_supplier,
                    "supplier_contact": down_cells.get(cols.get("Supplier Contact"), ""),
                    "po_nr": str(down_cells.get(cols.get("PO NR"), "")).strip(),
                    "date_po": str(down_cells.get(cols.get("Date of PO"), "")).strip(),
                    "received": bool(down_cells.get(cols.get("Received"), False)),
                    "date_received": str(down_cells.get(cols.get("Date Received"), "")).strip(),
                    "followed_up": bool(down_cells.get(cols.get("Followed up"), False)),
                    "comments": str(down_cells.get(cols.get("Follow up comments / ETA"), "")).strip()
                }
                item["status_color"] = determine_status_color(item)
                items.append(item)

            down_idx += 1

        # Save project entry
        projects.append({
            "owner": project_owner,
            "project_name": project_name,
            "ac_number": ac_number,
            "items": items
        })

with open("projects_data.json", "w") as f:
    json.dump(projects, f, indent=2)

print(f"Successfully processed {len(projects)} color-anchored project blocks.")
