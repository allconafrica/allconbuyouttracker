import os
import re
import smtplib
from datetime import datetime, timezone, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smartsheet

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = os.environ.get("SMTP_EMAIL")
SENDER_PASSWORD = os.environ.get("SMTP_PASSWORD")
NORMAN_EMAIL = "norman@allconafrica.co.za"

smart = smartsheet.Smartsheet(os.environ.get("SMARTSHEET_TOKEN"))
target_sheet_name = "ALLCON PROJECTS BUY OUT LIST"

response = smart.Sheets.list_sheets(include_all=True)
sheet_id = next((s.id for s in response.data if s.name.strip().upper() == target_sheet_name.upper()), response.data[0].id)

sheet = smart.Sheets.get_sheet(sheet_id)
cols = {col.title.strip().upper(): col.id for col in sheet.columns}

def get_val(cell_map, col_name):
    cid = cols.get(col_name.upper())
    return str(cell_map.get(cid, "")).strip() if cid else ""

AC_REGEX = re.compile(r'\bAC\s*\d+', re.IGNORECASE)
OWNERS = ["NORMAN", "NORLENE"]

now = datetime.now(timezone.utc)
# Look back window: modified within the last 30 minutes
LOOKBACK_WINDOW = timedelta(minutes=30)

all_rows = list(sheet.rows)
received_projects = []

for idx, row in enumerate(all_rows):
    cell_values = {cell.column_id: (cell.value or cell.display_value or "") for cell in row.cells}
    row_ac_match = next((str(v).strip() for v in cell_values.values() if AC_REGEX.search(str(v))), None)

    if row_ac_match:
        ac_number = row_ac_match
        project_name = ""
        project_owner = "UNASSIGNED"

        # Scan Upward for Owner & Project Name
        up_idx = idx - 1
        while up_idx >= 0:
            up_row = all_rows[up_idx]
            up_cells = {cell.column_id: (cell.value or cell.display_value or "") for cell in up_row.cells}
            
            if any(AC_REGEX.search(str(v)) for v in up_cells.values()):
                break

            for v in up_cells.values():
                val_upper = str(v).strip().upper()
                for o in OWNERS:
                    if o in val_upper:
                        project_owner = o
                        break

            up_texts = [str(v).strip() for v in up_cells.values() if str(v).strip()]
            if up_texts and up_texts[0].upper() not in OWNERS:
                project_name = up_texts[0]
                break
            up_idx -= 1

        if not project_name:
            project_name = ac_number

        if project_owner.upper() != "NORMAN":
            continue

        # Scan Downward for newly received items
        items = []
        down_idx = idx + 1
        while down_idx < len(all_rows):
            down_row = all_rows[down_idx]
            down_cells = {cell.column_id: (cell.value or cell.display_value or "") for cell in down_row.cells}

            if any(AC_REGEX.search(str(v)) for v in down_cells.values()):
                break

            down_texts = [str(v).strip() for v in down_cells.values() if str(v).strip()]
            has_po = bool(get_val(down_cells, "PO NR") or get_val(down_cells, "DATE OF PO"))
            
            if down_texts and not has_po and down_idx + 1 < len(all_rows):
                next_cells = {cell.column_id: (cell.value or cell.display_value or "") for cell in all_rows[down_idx + 1].cells}
                if any(AC_REGEX.search(str(v)) for v in next_cells.values()):
                    break

            if down_texts:
                component = down_texts[0]
                received = bool(down_cells.get(cols.get("RECEIVED"), False))
                date_received = get_val(down_cells, "DATE RECEIVED")
                po_nr = get_val(down_cells, "PO NR")

                row_mod_date = down_row.modified_at or row.modified_at

                # Target items marked RECEIVED recently
                if received and row_mod_date and (now - row_mod_date) <= LOOKBACK_WINDOW:
                    items.append({
                        "component": component,
                        "po_nr": po_nr or "N/A",
                        "date_received": date_received or "Today"
                    })

            down_idx += 1

        if items:
            received_projects.append({
                "ac_number": ac_number,
                "project_name": project_name,
                "items": items
            })

# --- SEND RECEIVED NOTIFICATION ---
if received_projects:
    html_body = """
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; background-color: #f4f6f9; color: #333; margin: 0; padding: 20px; }
        .container { max-width: 680px; margin: 0 auto; background: #ffffff; border-radius: 8px; padding: 25px; border: 1px solid #e1e4e8; }
        .header { border-bottom: 2px solid #2ecc71; padding-bottom: 12px; margin-bottom: 20px; }
        .header h2 { margin: 0; color: #2ecc71; font-size: 20px; }
        .project-block { margin-bottom: 25px; border: 1px solid #e9ecef; border-radius: 6px; overflow: hidden; }
        .project-header { background: #1e1e1e; color: #ffffff; padding: 10px 15px; }
        .ac-title { font-size: 16px; font-weight: bold; }
        .project-title { font-size: 13px; color: #bbb; }
        table { width: 100%; border-collapse: collapse; }
        th { background: #f8f9fa; text-align: left; padding: 10px; font-size: 11px; text-transform: uppercase; color: #6c757d; border-bottom: 1px solid #dee2e6; }
        td { padding: 12px 10px; border-bottom: 1px solid #edf2f7; font-size: 13px; }
        .badge { display: inline-block; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 11px; color: #fff; background-color: #2ecc71; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h2>GOODS RECEIVED NOTICE</h2>
          <p style="font-size: 12px; color: #666; margin-top: 5px;">Stores Update Notification</p>
        </div>
    """

    for p in received_projects:
        html_body += f"""
        <div class="project-block">
          <div class="project-header">
            <div class="ac-title">{p['ac_number']}</div>
            <div class="project-title">PROJECT: {p['project_name']}</div>
          </div>
          <table>
            <thead>
              <tr>
                <th>Item</th>
                <th>PO Number</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
        """
        for item in p['items']:
            html_body += f"""
              <tr>
                <td><strong>{item['component']}</strong></td>
                <td>{item['po_nr']}</td>
                <td><span class="badge">RECEIVED ({item['date_received']})</span></td>
              </tr>
            """
        html_body += "</tbody></table></div>"

    html_body += "</div></body></html>"

    msg = MIMEMultipart("alternative")
    msg['Subject'] = "ITEMS RECEIVED IN STORES"
    msg['From'] = SENDER_EMAIL
    msg['To'] = NORMAN_EMAIL
    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)

    print(f"Received alert email sent to {NORMAN_EMAIL}.")
