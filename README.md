# Contract Note Automation

An end-to-end automated pipeline designed to fetch, unlock, parse, and log stock equity contract notes directly from Gmail into a structured, professionally formatted multi-sheet Excel workbook (`trade_book.xlsx`)[cite: 3]. 

---

## 📂 Project Architecture

```text
contract_note_automation/
│
├── .env                      # Configuration & credentials (ignored in git)
├── requirements.txt          # Python package dependencies
├── contract_automation.py    # Unified fetch, unlock, parse, Excel merge, and cleanup script
├── run_automation.sh         # Master Unix/macOS shell runner & automation wrapper
├── run_status.json           # Execution tracking log (stores dates with day of the week & status)
└── trade_book.xlsx           # Output workbook containing 'Trades' and 'Charges' sheets

```

---

## ⚙️ Prerequisites & System Dependencies

Because **Camelot** relies on image processing for table grids, **Ghostscript** is required on your system level.

### **macOS**

```bash
brew install ghostscript

```

### **Windows**

Choose one of the following installation methods:

* **Chocolatey:** `choco install ghostscript`
* **Scoop:** `scoop install ghostscript`
* **Manual:** Download the installer from the [Ghostscript Releases Page](https://ghostscript.com/) and ensure the `bin` path is added to your System Environment Variables (`PATH`).



---

## 🛠️ Installation & Setup

1. Clone or Navigate to the Project Directory:

```bash
cd contract_note_automation

```

2. Create and Activate a Virtual Environment:

* macOS / Linux:
```bash
python3 -m venv .venv
source .venv/bin/activate

```


* Windows:
```cmd
python -m venv .venv
.venv\Scripts\activate

```



3. Install Python Dependencies:

```bash
pip install -r requirements.txt

```

4. Configure Your Environment Variables (`.env`):
Create a `.env` file in the root directory and populate it with your specific configuration:

```env
EMAIL_USER=your_email@gmail.com
EMAIL_PASS=your_gmail_app_password
IMAP_SERVER=imap.gmail.com
DOWNLOAD_DIR=/path/to/your/project/folder
EXCEL_PATH=/path/to/your/project/folder/trade_book.xlsx
PDF_PASSWORD=your_pdf_password

```

---

## 🚀 Usage

You can run the full automation pipeline manually via the master runner script, which handles environment loading, email fetching, table extraction, Excel updates, and download cleanup automatically.

### **1. Run for Yesterday (Default)**

If no arguments are provided, the script targets yesterday's date:

```bash
./run_automation.sh

```

### **2. Run for a Specific Date (`YYYYMMDD`)**

```bash
./run_automation.sh 20260806

```

### **3. Run for a Specific PDF Filename**

```bash
(Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned) ; (& d:\Lalith\contract_note_automation\.venv\Scripts\Activate.ps1)
./run_automation.sh 06-08-2026-contract-notes_XNY058.pdf

```

---

## ⏰ Automation & Scheduling

### **Weekend & Catch-up Behavior**

* **No Weekend Skipping:** The automation attempts to fetch and process contract notes every single day (including weekends). If markets were closed and no contract note exists, it cleanly records `"Not Found"` in the status log and proceeds.


* **Smart Catch-up:** If left unrun for multiple days, the script automatically catches up sequentially from your last successful run date up to yesterday.



### **macOS (launchd / LaunchAgents)**

To automate the script on macOS using native background agents (`launchd`), set up a property list (`.plist`) file.

1. Create and edit the plist file in your LaunchAgents directory:

```bash
nano ~/Library/LaunchAgents/com.lchatala.contractautomation.plist

```

2. Paste the following configuration:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "[http://apple.com/DTDs/PropertyList-1.0.dtd](http://apple.com/DTDs/PropertyList-1.0.dtd)">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.lchatala.contractautomation</string>

    <key>ProgramArguments</key>
    <array>
        <string>/bin/bash</string>
        <string>/Users/lchatala/Desktop/Lalith Chatala/Personal/Vscode/contract_note_automation/run_automation.sh</string>
    </array>

    <key>StartCalendarInterval</key>
    <dict>
        <key>Hour</key>
        <integer>9</integer>
        <key>Minute</key>
        <integer>0</integer>
    </dict>

    <key>StandardOutPath</key>
    <string>/Users/lchatala/Desktop/Lalith Chatala/Personal/Vscode/contract_note_automation/cron.log</string>

    <key>StandardErrorPath</key>
    <string>/Users/lchatala/Desktop/Lalith Chatala/Personal/Vscode/contract_note_automation/cron.log</string>
</dict>
</plist>

```

3. Load the background agent using `launchctl`:

```bash
launchctl load ~/Library/LaunchAgents/com.lchatala.contractautomation.plist

```

---

## 📊 Output Features

* **Trades Sheet:** Appends buy/sell records with structured pricing, net quantities, and a clean Blue Table style.


* **Charges Sheet:** Aggregates statutory levies, GST, brokerage, and STT, complete with an automated Excel totals row and styling.


* **Run Status Log (`run_status.json`):** Tracks history cleanly with entries recorded in `YYYY-MM-DD (Day)` format alongside statuses like `Success` or `Not Found`.


* **Automated Cleanup:** Safely purges downloaded source PDFs and temporary unlocked files upon run completion.

# Initial setup
cd <to the project path>
python -m venv .venv
.\.venv\Scripts\Activate
pip install -r requirements.txt (if not installed)
python Tests/validate_pkg.py
python contract_note_automation.py


# important notes while changing .env file and .gitignore
.env file variable add/change make change in secrets in github repo to work properly
add sufficient files in .gitignore too

# Notes
automation can break if gmail account passoword or auth settings is changed, then we need to create new app password and store it in .env file and also in github secrets