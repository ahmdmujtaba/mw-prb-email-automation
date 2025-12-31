# =========================
# IMPORTS
# =========================

# Regular expressions → used to extract PRB#######
import re

# win32com → allows Python to control Outlook Desktop via COM
import win32com.client

# openpyxl → read/write Excel (.xlsx)
from openpyxl import load_workbook


# =========================
# CONFIGURATION SECTION
# =========================

# Base working directory where everything lives
BASE_DIR = r"E:\Jazz\Automation\MW Tracker Email Automation"

# Path to the Excel file that stores PRBs
EXCEL_FILE = rf"{BASE_DIR}\MW_Tickets_Raw.xlsx"

# Excel sheet name
SHEET_NAME = "Sheet1"

# Column number where PRBs are stored (Column A = 1)
PRB_COLUMN = 1

# Outlook mailbox name (must match exactly what Outlook shows)
MAILBOX_NAME = "ahmed.mujtaba@jazz.com.pk"

# Target Outlook folder under the mailbox root
TARGET_FOLDER = "PRB TT"

# Only accept emails from this sender
ALLOWED_SENDER = "tec-noc-analytics@jazz.com.pk"


# =========================
# REGEX DEFINITIONS
# =========================

# Regex to match PRB followed by exactly 7 digits
# Example: PRB0499824
PRB_REGEX = re.compile(r"(PRB\d{7})", re.IGNORECASE)


# =========================
# LOAD EXCEL WORKBOOK
# =========================

# Open the Excel workbook
wb = load_workbook(EXCEL_FILE)

# Select the target worksheet
ws = wb[SHEET_NAME]

# Read all existing PRBs from Column A into a Python set
# Why set?
# - Fast lookup
# - Prevents duplicates
existing_prbs = {
    ws.cell(row=r, column=PRB_COLUMN).value
    for r in range(2, ws.max_row + 1)  # Skip header row
    if ws.cell(row=r, column=PRB_COLUMN).value
}


# =========================
# CONNECT TO OUTLOOK
# =========================

# Create Outlook COM object
outlook = win32com.client.Dispatch("Outlook.Application")

# Get MAPI namespace (entry point to all mailboxes)
namespace = outlook.GetNamespace("MAPI")

# Access mailbox root (NOT Inbox)
root = namespace.Folders[MAILBOX_NAME]

# Access the specific folder where PRB emails are stored
prb_folder = root.Folders[TARGET_FOLDER]

# Get all messages in the folder
messages = prb_folder.Items

# Sort messages so newest emails are processed first
messages.Sort("[ReceivedTime]", True)


# =========================
# PROCESS EMAILS
# =========================

for msg in messages:

    # ---- Safely read subject ----
    try:
        subject = msg.Subject or ""
    except Exception:
        # Skip corrupt or non-mail items
        continue


    # =========================
    # SENDER FILTER
    # =========================

    sender_email = None

    try:
        # Sender object (Exchange user)
        sender = msg.Sender

        if sender:
            # Convert Exchange user → real SMTP address
            exch_user = sender.GetExchangeUser()
            if exch_user:
                sender_email = exch_user.PrimarySmtpAddress
    except Exception:
        # If anything fails, treat sender as invalid
        pass

    # Skip email if sender is not exactly what we want
    if not sender_email or sender_email.lower() != ALLOWED_SENDER:
        continue


    # =========================
    # SUBJECT KEYWORD FILTER
    # =========================

    # Convert subject to lowercase for safe comparisons
    subject_lower = subject.lower()

    # Require BOTH words to be present (order does not matter)
    if "high" not in subject_lower or "utilization" not in subject_lower:
        continue


    # =========================
    # PRB EXTRACTION
    # =========================

    # Search subject for PRB#######
    match = PRB_REGEX.search(subject)

    # If PRB pattern not found, skip
    if not match:
        continue

    # Extract PRB and normalize to uppercase
    prb = match.group(1).upper()


    # =========================
    # EXCEL DEDUPLICATION
    # =========================

    # Skip if PRB already exists in Excel
    if prb in existing_prbs:
        continue


    # =========================
    # APPEND TO EXCEL
    # =========================

    # Append PRB to the next empty row in Column A
    ws.cell(row=ws.max_row + 1, column=PRB_COLUMN).value = prb

    # Track PRB to prevent duplicates in same run
    existing_prbs.add(prb)

    print(f"✅ Added {prb}")


# =========================
# SAVE EXCEL FILE
# =========================

# Save changes to Excel
wb.save(EXCEL_FILE)

print("✔ Done. Outlook PRBs synced safely.")
