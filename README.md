# Digii_Monday_Bot 🚀
> Production-ready, modular Python automation tool for Windows to automate recurring administrative tasks, report extractions, and compliance archiving on the **DIGII ERP** web portal (`https://poornima.digiicampus.com/V2/`).

---

## 📋 Table of Contents
1. [Overview & Architecture](#-overview--architecture)
2. [Safe Zero-Write-Risk Demo Tasks](#-safe-zero-write-risk-demo-tasks)
3. [Project Layout](#-project-layout)
4. [Prerequisites & Quickstart](#-prerequisites--quickstart)
5. [Configuring URLs & Actions (`urls_config.xlsx`)](#-configuring-urls--actions-urls_configxlsx)
6. [Using `playwright codegen` to Capture Selectors](#-using-playwright-codegen-to-capture-selectors)
7. [Scheduling Weekly Monday Runs via Windows Task Scheduler](#-scheduling-weekly-monday-runs-via-windows-task-scheduler)
8. [Run Outputs & Executive Excel Audit Report](#-run-outputs--executive-excel-audit-report)
9. [Advanced Configurations & Troubleshooting](#-advanced-configurations--troubleshooting)

---

## 🌟 Overview & Architecture

`Digii_Monday_Bot` is designed for institutional administrators and automation engineers. It eliminates repetitive Monday morning portal logins, report extractions, and compliance archiving:

- **Dynamic Multi-URL Configuration:** Modify target modules, export triggers, and enable/disable tasks directly via Excel (`config/urls_config.xlsx`) without touching code.
- **Hybrid Authenticated Portal Session:** Reuses session credentials via cached Playwright `auth_state.json`. If a session expires or `auth_state.json` is missing, it cleanly falls back to fresh `.env` authentication and refreshes the cache.
- **Resilient SPA Extraction & Proof (`keep ss`):** Navigates to dynamic single-page application (SPA) views, waits for dynamic data rendering/network settlement (`networkidle`), triggers export downloads, and captures full-page screenshot proof after login, after downloads, and upon any error.
- **Zero-Write-Risk Compliance:** Supports form dry-run testing with aggressive browser autofill clearing and automatic cancellation so zero records are committed to the live database.
- **Automated Run Archiving & Styled Executive Audit Report:** Every execution creates `Runs/YYYY-MM-DD_HHMM/` containing `raw_downloads/`, `screenshots/`, and an executive `Audit_Report.xlsx` featuring color-coded statuses, direct clickable hyperlinks to downloaded files and screenshots, and auto-opens upon completion.
- **One-Click Windows Launcher:** `Run_Monday_Bot.bat` provides instant non-technical execution and validation.

---

## 🛡️ Safe Zero-Write-Risk Demo Tasks

The bot comes pre-configured with three live demo tasks targeted at `https://poornima.digiicampus.com/V2/`:

### 1. Task 1 (`REPORT_MODAL_DOWNLOAD`) — Batch Report Extraction
- **Route:** `https://poornima.digiicampus.com/V2/#/staffManagement`
- **Action:** Clicks the `Reports` button to open the "Select a Report to Download" modal.
- **Downloads:**
  - Row 1: `Faculty & Administrator Login Report`
  - Row 2: `Faculty & Administrator Not Logged In Report`
- **Output:** Saves exported spreadsheets to `raw_downloads/` and captures full-page screenshot proofs.

### 2. Task 2 (`DIRECT_DOWNLOAD`) — Master Data Export
- **Route:** `https://poornima.digiicampus.com/V2/#/staffManagement`
- **Action:** Clicks the `Download Staff Data` button.
- **Output:** Saves exported staff spreadsheet into `raw_downloads/` and captures full-page screenshot proof.

### 3. Task 3 (`FORM_DRY_RUN`) — Create New Staff Record Dry Run
- **Route:** `https://poornima.digiicampus.com/V2/#/staffManagement/add`
- **Autofill Clearing:** Chrome autofill pre-populates `Employee ID` and `Password`. The bot explicitly clears all autofill fields before populating.
- **Safe Test Values:**
  - `Employee ID`: `DEMO_TEST_001`
  - `User Type`: `Faculty`
  - `Department`: First available option in dropdown
  - `First Name`: `Demo`
  - `Last name`: `Faculty`
  - `Gender`: `Male`
  - `Email ID`: `demo.test@poornima.org`
  - `Phone Number`: `9999999999`
  - `Password`: `DemoPass@123`
- **Proof:** Captures full-page screenshot `04_Form_Dry_Run_Populated.png`.
- **CRITICAL SAFETY RULE:** The bot **never** clicks `Create User`. It clicks `Cancel` to guarantee zero records are written to the live database, and records `DRY_RUN_SUCCESS` in `Audit_Report.xlsx`.

---

## 📁 Project Layout

```text
Digii_Monday_Bot/
├── config/
│   └── urls_config.xlsx      # Pre-populated target URL & export action matrix
├── Runs/                     # Git-ignored run archives (YYYY-MM-DD_HHMM/)
│   └── 2026-11-20_0915/
│       ├── raw_downloads/    # Extracted Excel/CSV report files
│       ├── screenshots/      # Full-page PNG visual audit proofs
│       └── Audit_Report.xlsx # Formatted Executive Audit Report with hyperlinks
├── .env.example              # Template for credentials & portal configuration
├── .env                      # Real credentials (Git-ignored)
├── .gitignore                # Git ignore rules for safety and security
├── requirements.txt          # Python dependencies (playwright, pandas, openpyxl, python-dotenv)
├── digii_bot.py              # Modular Playwright + OpenPyXL automation engine
├── Run_Monday_Bot.bat        # 1-click Windows batch launcher
└── README.md                 # Documentation, selector recording, and scheduler guide
```

---

## ⚡ Prerequisites & Quickstart

### 1. Prerequisites
- **Windows 10 / 11 / Windows Server**
- **Python 3.10+** (Make sure **"Add Python to PATH"** was checked during Python installation)

### 2. Initial Setup
1. **Clone or download** this repository to `C:\Digii_Monday_Bot`.
2. **Create your `.env` file** from the template:
   ```cmd
   copy .env.example .env
   ```
3. **Edit `.env`** with your credentials:
   ```ini
   LOGIN_URL=https://poornima.digiicampus.com/V2/#/login
   DIGII_USER=your_username_or_email
   DIGII_PASS=your_secret_password
   HEADLESS=false
   ```
4. **Install dependencies & Playwright browser:**
   ```cmd
   pip install -r requirements.txt
   playwright install chromium
   ```

### 3. Launching the Bot
Simply double-click **`Run_Monday_Bot.bat`** or run from PowerShell / Command Prompt:
```cmd
Run_Monday_Bot.bat
```
*(By default, `HEADLESS=false` allows you to watch the browser work in real-time during live demos.)*

---

## 📊 Configuring URLs & Actions (`urls_config.xlsx`)

Open `config/urls_config.xlsx` in Microsoft Excel to manage tasks:

| Task_Name | Mode | Target_URL | Action_Target | Active |
| :--- | :--- | :--- | :--- | :--- |
| `Faculty_Admin_Login_Report` | `REPORT_MODAL_DOWNLOAD` | `https://poornima.digiicampus.com/V2/#/staffManagement` | `button:has-text("Reports") -> text="Faculty & Administrator Login Report"` | `TRUE` |
| `Faculty_Admin_Not_Logged_In_Report` | `REPORT_MODAL_DOWNLOAD` | `https://poornima.digiicampus.com/V2/#/staffManagement` | `button:has-text("Reports") -> text="Faculty & Administrator Not Logged In Report"` | `TRUE` |
| `Download_Staff_Data_Master` | `DIRECT_DOWNLOAD` | `https://poornima.digiicampus.com/V2/#/staffManagement` | `button:has-text("Download Staff Data"), button:has-text("Download Staff")` | `TRUE` |
| `Create_Staff_Record_Dry_Run` | `FORM_DRY_RUN` | `https://poornima.digiicampus.com/V2/#/staffManagement/add` | `button:has-text("Cancel")` | `TRUE` |

---

## 🔍 Using `playwright codegen` to Capture Selectors

To record new selectors or verify portal UI elements:
```cmd
playwright codegen https://poornima.digiicampus.com/V2/#/login
```
Copy any target locator into `urls_config.xlsx`.

---

## ⏰ Scheduling Weekly Monday Runs via Windows Task Scheduler

To execute `Digii_Monday_Bot` automatically every Monday at 8:00 AM:

### Option A: Via Graphical Task Scheduler UI (`taskschd.msc`)
1. Press `Win + R`, type **`taskschd.msc`**, and press Enter.
2. In the right pane, click **Create Task...**:
   - **General:** Name: `DIGII_Monday_Bot_Weekly_Run`. Run with highest privileges.
   - **Triggers:** Weekly, recur every 1 week on **Monday** at `08:00:00 AM`.
   - **Actions:** Start a program: `C:\Digii_Monday_Bot\Run_Monday_Bot.bat`
     - **Start in:** `C:\Digii_Monday_Bot\` *(Mandatory for batch files to locate files)*
   - **Settings:** Stop the task if it runs longer than 1 hour. Click OK.

### Option B: Via One-Line PowerShell Command (Admin)
```powershell
$action = New-ScheduledTaskAction -Execute "C:\Digii_Monday_Bot\Run_Monday_Bot.bat" -WorkingDirectory "C:\Digii_Monday_Bot"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 8:00am
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)
Register-ScheduledTask -TaskName "DIGII_Monday_Bot" -Action $action -Trigger $trigger -Settings $settings -Description "Recurring Monday DIGII ERP Bot"
```

---

**Digii_Monday_Bot** | Automated, Reliable, Audited.
#   D i g g i _ E R P _ B o t  
 