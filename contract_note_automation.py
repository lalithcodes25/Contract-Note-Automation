from datetime import datetime, timedelta
import email
from email.header import decode_header
import imaplib
import json
import os
import re
import sys
import camelot
import openpyxl
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo
import pandas as pd
from pypdf import PdfReader, PdfWriter
from dotenv import load_dotenv

# Load variables from .env file
load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
IMAP_SERVER = os.getenv("IMAP_SERVER")
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR")
PDF_PASSWORD = os.getenv("PDF_PASSWORD")
EXCEL_PATH = os.getenv("EXCEL_PATH")
STATUS_FILE = "run_status.json"

excel_path = (
    EXCEL_PATH
    if EXCEL_PATH
    else (
        "/Users/lchatala/Desktop/Lalith"
        " Chatala/Personal/Vscode/contract_note_automation/trade_book.xlsx"
    )
)


def get_last_run_date():
  """Reads and returns the latest successfully processed date from JSON."""
  if not os.path.exists(STATUS_FILE):
    return None
  try:
    with open(STATUS_FILE, "r") as f:
      data = json.load(f)
      if not data:
        return None
      dates = []
      for d in data.keys():
        clean_d = d.split()[0]
        dates.append(datetime.strptime(clean_d, "%Y-%m-%d").date())
      return max(dates) if dates else None
  except (ValueError, IOError, json.JSONDecodeError):
    return None


def update_status(target_date, status="Success"):
  """Updates JSON status file with date, day of the week, and status."""
  data = {}
  if os.path.exists(STATUS_FILE):
    try:
      with open(STATUS_FILE, "r") as f:
        data = json.load(f)
    except (IOError, json.JSONDecodeError):
      data = {}

  date_str = (
      f"{target_date.strftime('%Y-%m-%d')} ({target_date.strftime('%A')})"
  )
  data[date_str] = status

  sorted_data = dict(
      sorted(
          data.items(),
          key=lambda x: datetime.strptime(x[0].split()[0], "%Y-%m-%d"),
      )
  )

  with open(STATUS_FILE, "w") as f:
    json.dump(sorted_data, f, indent=4)


def get_client_suffix(download_dir):
  if not os.path.exists(download_dir):
    return "XNY058"
  for f in os.listdir(download_dir):
    if f.endswith(".pdf") and "_unlocked" not in f:
      match = re.search(r"-contract-notes_(.*?)\.pdf", f)
      if match:
        return match.group(1)
  return "XNY058"


def clean_val(val):
  s = str(val).strip()
  if not s or s == "nan":
    return 0.0
  s = s.replace("₹", "").replace(",", "").strip()
  if s.startswith("(") and s.endswith(")"):
    s = "-" + s[1:-1]
  try:
    return float(s)
  except ValueError:
    return 0.0


def extract_pdf_recursively(msg, target_filename, download_dir):
  downloaded = False
  for part in msg.walk():
    content_type = part.get_content_type()
    if content_type == "message/rfc822":
      payload = part.get_payload()
      if isinstance(payload, list):
        for sub_msg in payload:
          if downloaded:
            break
          if extract_pdf_recursively(sub_msg, target_filename, download_dir):
            downloaded = True
      elif isinstance(payload, email.message.Message):
        if extract_pdf_recursively(payload, target_filename, download_dir):
          downloaded = True
      if downloaded:
        break

    filename = part.get_filename()
    if filename and filename.lower() == target_filename.lower():
      filepath = os.path.join(download_dir, filename)
      if not os.path.exists(filepath):
        payload_data = part.get_payload(decode=True)
        if payload_data:
          with open(filepath, "wb") as f:
            f.write(payload_data)
          print(f"Downloaded: {filename} to {download_dir}")
          downloaded = True
      else:
        print(f"Skipped download (already exists): {filename}")
        downloaded = True
  return downloaded


def cleanup_downloads():
  """Scans download directory and removes all remaining PDF files."""
  if not DOWNLOAD_DIR or not os.path.exists(DOWNLOAD_DIR):
    return

  print(f"Scanning directory for PDF cleanup: {DOWNLOAD_DIR}")
  deleted_count = 0

  for filename in os.listdir(DOWNLOAD_DIR):
    if filename.lower().endswith(".pdf"):
      file_path = os.path.join(DOWNLOAD_DIR, filename)
      try:
        os.remove(file_path)
        print(f"Deleted: {filename}")
        deleted_count += 1
      except Exception as e:
        print(f"Failed to delete {filename}: {e}")

  print(
      f"Cleanup complete. Successfully removed {deleted_count} PDF file(s)."
  )


def process_date(target_date, mail):
  target_date_str = target_date.strftime("%Y-%m-%d")
  client_suffix = get_client_suffix(DOWNLOAD_DIR)
  date_str_formatted = target_date.strftime("%d-%m-%Y")
  target_filename = f"{date_str_formatted}-contract-notes_{client_suffix}.pdf"
  subject_date_str = (
      f"{target_date.strftime('%B')} {target_date.day}, {target_date.year}"
  )
  target_subject_phrase = (
      f"Combined Equity Contract Note for {client_suffix} - {subject_date_str}"
  )

  print(
      f"==> Processing contract note for {target_date_str}"
      f" ({target_date.strftime('%A')}) | Target file: {target_filename}"
  )

  if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

  # 1. Fetch via IMAP (Filename first, then Subject fallback)
  fallback_query = f"X-GM-RAW filename:{target_filename}"
  status, messages = mail.search(None, fallback_query)
  email_ids = messages[0].split() if messages and messages[0] else []

  if not email_ids:
    print("    -> Filename search failed. Trying fallback search by subject...")
    search_query = f'X-GM-RAW "subject:\\"{target_subject_phrase}\\""'
    status, messages = mail.search(None, search_query)
    email_ids = messages[0].split() if messages and messages[0] else []

  if email_ids:
    for e_id in email_ids:
      res, msg_data = mail.fetch(e_id, "(RFC822)")
      for response_part in msg_data:
        if isinstance(response_part, tuple):
          msg = email.message_from_bytes(response_part[1])
          extract_pdf_recursively(msg, target_filename, DOWNLOAD_DIR)

  file_path = os.path.join(DOWNLOAD_DIR, target_filename)
  unlocked_path = os.path.join(
      DOWNLOAD_DIR, target_filename.replace(".pdf", "_unlocked.pdf")
  )

  if not os.path.exists(file_path):
    print(
        f"    -> No contract note found locally or via email for"
        f" {target_date_str}."
    )
    update_status(target_date, "Not Found")
    return False

  # 2. Unlock PDF & Parse Tables
  reader = PdfReader(file_path)
  if reader.is_encrypted:
    reader.decrypt(PDF_PASSWORD)
  writer = PdfWriter()
  for page in reader.pages:
    writer.add_page(page)
  with open(unlocked_path, "wb") as f:
    writer.write(f)

  tables = camelot.read_pdf(unlocked_path, pages="all", flavor="lattice")
  trade_date = target_date.strftime("%Y%m%d")

  # Parse Trades
  trade_rows = []
  annexure_df = None
  for t in tables:
    df = t.df
    header_text = " ".join(str(val) for val in df.iloc[0].values)
    if "Security / Contract Description" in header_text or "Order No." in header_text:
      annexure_df = df
      break
  if annexure_df is None and len(tables) >= 6:
    annexure_df = tables[5].df

  if annexure_df is not None:
    header_row_idx = 0
    for r_idx in range(len(annexure_df)):
      row_str = " ".join(str(val) for val in annexure_df.iloc[r_idx].values)
      if "Security / Contract Description" in row_str and "Buy(B)" in row_str:
        header_row_idx = r_idx
        break

    for r_idx in range(header_row_idx + 1, len(annexure_df)):
      row = annexure_df.iloc[r_idx]
      sec_desc_raw = str(row.iloc[4]).strip()
      if (
          not sec_desc_raw
          or "ISIN" in sec_desc_raw.upper()
          or "TOTAL" in sec_desc_raw.upper()
          or "EQUITY" in sec_desc_raw.upper()
      ):
        continue
      stock_name = (
          sec_desc_raw.split("/")[0].replace("-A", "").replace("\n", " ").strip()
      )
      bs = str(row.iloc[5]).strip().upper()
      qty_raw = clean_val(row.iloc[7])
      brokerage_raw = clean_val(row.iloc[8])
      net_rate = clean_val(row.iloc[9])
      if qty_raw == 0:
        continue
      qty = int(qty_raw) if qty_raw.is_integer() else qty_raw
      trade_rows.append({
          "date": trade_date,
          "stock name": stock_name,
          "B/S": "B" if bs.startswith("B") else "S",
          "qty": -qty if bs.startswith("S") else qty,
          "price of share": net_rate,
          "brokerage per share": brokerage_raw,
          "total value": round(abs(qty) * net_rate, 2),
      })

  new_trades_df = pd.DataFrame(trade_rows)

  # Parse Charges
  brokerage_val, stt_val, other_charges_val = 0.0, 0.0, 0.0
  for t in tables:
    df_t = t.df
    for _, row in df_t.iterrows():
      row_str = " ".join(str(val) for val in row.values).lower()
      if "pay in" in row_str or "net amount" in row_str:
        continue
      if any(
          kw in row_str
          for kw in [
              "taxable value",
              "securities transaction tax",
              "exchange transaction",
              "clearing",
              "cgst",
              "sgst",
              "igst",
              "sebi",
              "stamp",
          ]
      ):
        net_col_idx = len(df_t.columns) - 1
        desc = str(row.iloc[0]).strip().lower()
        val = abs(clean_val(row.iloc[net_col_idx]))
        if "taxable value" in desc or "brokerage" in desc:
          brokerage_val += val
        elif "securities transaction tax" in desc:
          stt_val += val
        elif any(
            kw in desc
            for kw in [
                "exchange transaction",
                "clearing",
                "cgst",
                "sgst",
                "igst",
                "sebi",
                "stamp",
            ]
        ):
          other_charges_val += val

  new_charges_df = pd.DataFrame([{
      "date": trade_date,
      "brokerage": round(brokerage_val, 2),
      "stt": round(stt_val, 2),
      "other charges (excep brokerage, stt)": round(other_charges_val, 2),
      "total charges": round(
          brokerage_val + stt_val + other_charges_val, 2
      ),
  }])

  # 3. Update Excel
  existing_trades_df, existing_charges_df = pd.DataFrame(), pd.DataFrame()
  if os.path.exists(excel_path):
    try:
      existing_trades_df = pd.read_excel(excel_path, sheet_name="Trades")
      existing_trades_df["date"] = existing_trades_df["date"].astype(str)
    except ValueError:
      pass
    try:
      existing_charges_df = pd.read_excel(excel_path, sheet_name="Charges")
      existing_charges_df["date"] = existing_charges_df["date"].astype(str)
    except ValueError:
      pass

  final_trades_df = pd.concat(
      [existing_trades_df, new_trades_df], ignore_index=True
  )
  final_charges_df = pd.concat(
      [existing_charges_df, new_charges_df], ignore_index=True
  )

  with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
    final_trades_df.to_excel(writer, sheet_name="Trades", index=False)
    final_charges_df.to_excel(writer, sheet_name="Charges", index=False)

  # Format Excel Tables
  wb = openpyxl.load_workbook(excel_path)
  if "Trades" in wb.sheetnames:
    ws = wb["Trades"]
    if ws.max_row > 1:
      ws.tables.clear()
      tab = Table(
          displayName="TradesTable",
          ref=f"A1:{openpyxl.utils.get_column_letter(ws.max_column)}{ws.max_row}",
      )
      tab.tableStyleInfo = TableStyleInfo(
          name="TableStyleMedium9", showRowStripes=True
      )
      ws.add_table(tab)

  if "Charges" in wb.sheetnames:
    ws = wb["Charges"]
    if ws.max_row > 1:
      ws.tables.clear()
      tab = Table(
          displayName="ChargesTable",
          ref=f"A1:{openpyxl.utils.get_column_letter(ws.max_column)}{ws.max_row}",
      )
      tab.tableStyleInfo = TableStyleInfo(
          name="TableStyleMedium3", showRowStripes=True
      )
      tab.showTotalsRow = True
      cols = []
      for i in range(len(ws[1])):
        tc = TableColumn(
            id=i + 1, name=str(ws.cell(row=1, column=i + 1).value)
        )
        tc.totalsRowLabel = "Total" if i == 0 else None
        tc.totalsRowFunction = None if i == 0 else "sum"
        cols.append(tc)
      tab.tableColumns = cols
      ws.add_table(tab)
  wb.save(excel_path)

  update_status(target_date, "Success")
  print(f"Successfully processed and updated status for {target_date_str}")
  return True

def download_all_contract_notes(mail):
  """Downloads all contract note PDFs from the inbox into DOWNLOAD_DIR."""
  if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

  print("Searching inbox for all contract-note emails...")
  status, messages = mail.search(None, 'X-GM-RAW "contract-notes"')
  email_ids = messages[0].split() if messages and messages[0] else []
  print(f"Found {len(email_ids)} matching email(s). Downloading attachments...")

  for e_id in email_ids:
    res, msg_data = mail.fetch(e_id, "(RFC822)")
    for response_part in msg_data:
      if isinstance(response_part, tuple):
        msg = email.message_from_bytes(response_part[1])
        for part in msg.walk():
          filename = part.get_filename()
          if filename and "contract-notes" in filename.lower() and not filename.endswith("_unlocked.pdf"):
            filepath = os.path.join(DOWNLOAD_DIR, filename)
            if not os.path.exists(filepath):
              payload_data = part.get_payload(decode=True)
              if payload_data:
                with open(filepath, "wb") as f:
                  f.write(payload_data)
                print(f"Downloaded: {filename}")

def main():
  today = datetime.now().date()
  target_end_date = today - timedelta(days=1)

  manual_date = None
  is_all_flag = False

  if len(sys.argv) > 1:
    arg = sys.argv[1].strip()
    if arg.lower() == "--all":
      is_all_flag = True
    else:
      try:
        if re.match(r"^\d{8}$", arg):
          manual_date = datetime.strptime(arg, "%Y%m%d").date()
        elif ".pdf" in arg.lower():
          match = re.search(r"(\d{2})-(\d{2})-(\d{4})", arg)
          if match:
            d, m, y = match.groups()
            manual_date = datetime(int(y), int(m), int(d)).date()
        else:
          manual_date = datetime.strptime(arg, "%Y-%m-%d").date()
      except ValueError:
        pass

  mail = imaplib.IMAP4_SSL(IMAP_SERVER)
  try:
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")

    if is_all_flag:
      print("==> Full historical sync active (--all).")
      download_all_contract_notes(mail)

      # Get all downloaded contract note PDFs and sort them chronologically
      pdf_files = [
          f
          for f in os.listdir(DOWNLOAD_DIR)
          if f.lower().endswith(".pdf") and "_unlocked" not in f
      ]
      pdf_files.sort(
          key=lambda x: datetime.strptime(
              re.search(r"(\d{2})-(\d{2})-(\d{4})", x).group(0), "%d-%m-%Y"
          )
          if re.search(r"(\d{2})-(\d{2})-(\d{4})", x)
          else datetime.min
      )

      for filename in pdf_files:
        match = re.search(r"(\d{2})-(\d{2})-(\d{4})-contract-notes", filename)
        if match:
          d, m, y = match.groups()
          target_date = datetime(int(y), int(m), int(d)).date()
          process_date(target_date, mail)
      return

    if manual_date:
      print(
          f"==> Manual override active. Processing single date:"
          f" {manual_date} ({manual_date.strftime('%A')})"
      )
      process_date(manual_date, mail)
    else:
      last_run = get_last_run_date()
      if last_run is None:
        print(
            "No previous run record found. Initializing for yesterday:"
            f" {target_end_date}"
        )
        process_date(target_end_date, mail)
      elif last_run >= target_end_date:
        print(f"Fetch already up to date. Last run was on {last_run}.")
      else:
        current_date = last_run + timedelta(days=1)
        print(f"Catching up from {current_date} to {target_end_date}...")
        while current_date <= target_end_date:
          process_date(current_date, mail)
          current_date += timedelta(days=1)

  finally:
    mail.logout()
    cleanup_downloads()


if __name__ == "__main__":
  main()