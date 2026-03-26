# IAW Data Formatting

Small utilities for preparing spreadsheets for Neon One data migration.

---

## What does this script do?

`merge_grant_contacts.py` is a Python script that takes two Excel spreadsheets you already have and combines them into one formatted spreadsheet that is ready to be imported into Neon One.

Specifically, it:

1. Reads your **grants spreadsheet** (File 1), which contains grant records like funder name, award amounts, dates, and notes.
2. Reads your **contacts spreadsheet** (File 2), which contains contact people associated with each foundation.
3. Matches them together by comparing the **Funder** column in File 1 against the **Foundation** column in File 2 (the match is not case-sensitive and ignores extra spaces).
4. Outputs a single merged spreadsheet with one row per matched contact/grant pair, formatted to the Neon One import template.

If a grant's funder has no matching foundation in the contacts file, the grant is still included in the output, just without contact details.

---

## Before you start: Install Python

This script requires **Python 3.9 or newer**. If you are not sure whether Python is installed on your computer, follow these steps.

### Check if Python is already installed

**On Windows:**
1. Press the Windows key, type `cmd`, and open **Command Prompt**.
2. Type `python --version` and press Enter.
3. If you see something like `Python 3.11.2`, you are good to go.
4. If you see an error or a version starting with `2.`, you need to install Python.

**On Mac:**
1. Press Cmd + Space, type `Terminal`, and open it.
2. Type `python3 --version` and press Enter.
3. If you see something like `Python 3.11.2`, you are good to go.

### Install Python (if needed)

Go to [https://www.python.org/downloads/](https://www.python.org/downloads/) and download the latest version. During installation on Windows, **make sure to check the box that says "Add Python to PATH"** before clicking Install.

---

## Setup (one-time only)

You only need to do this once. It installs the script's dependencies in an isolated environment so nothing interferes with other software on your computer.

Open a terminal (Command Prompt on Windows, Terminal on Mac) and navigate to the folder where you downloaded this project. You can do this by typing `cd ` (with a space after it) and then dragging the folder into the terminal window, then pressing Enter.

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

### Windows (Command Prompt)

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
pip install --upgrade pip
pip install -r requirements.txt
```

### Windows (PowerShell)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install -r requirements.txt
```

You will know setup worked when you see something like `Successfully installed pandas openpyxl ...` with no red error messages.

To exit the virtual environment when you are done working:

```bash
deactivate
```

---

## How to run the script

Every time you want to run the script, you first need to activate the virtual environment (the setup step above only needs to happen once, but the activation needs to happen each session).

**Mac/Linux:**
```bash
source .venv/bin/activate
```

**Windows (Command Prompt):**
```cmd
.venv\Scripts\activate.bat
```

**Windows (PowerShell):**
```powershell
.\.venv\Scripts\Activate.ps1
```

Then run the merge:

```bash
python merge_grant_contacts.py path/to/grants.xlsx path/to/contacts.xlsx -o path/to/output.xlsx
```

Replace the placeholders with your actual file paths. For example:

```bash
python merge_grant_contacts.py Input/foundation_grants.xlsx Input/contacts.xlsx -o Output/merged_export.xlsx
```

If you leave out `-o output_file.xlsx`, the script will automatically save the result as `merged_export.xlsx` in the same folder you are running the script from.

### Checking what columns the script detected (debug mode)

If the output looks wrong or columns are missing, you can run the script with `--debug-headers` to print out every column name it detected in both input files:

```bash
python merge_grant_contacts.py path/to/grants.xlsx path/to/contacts.xlsx --debug-headers -o path/to/output.xlsx
```

---

## Input file requirements

The script looks for specific column names in each file. Column names are matched loosely (case-insensitive, extra spaces ignored), but the names must be recognizable. Below is what each file needs.

### File 1: Grants spreadsheet

| Column name | What it maps to in the output |
|---|---|
| Funder | Company Name (used for matching) |
| Purpose | Grant Purpose |
| Amount | Ask Amount |
| Received | Funded Amount |
| Remaining | Grant Remaining |
| Award Date | Grant Award Date |
| Grant Start Date | Funded Date |
| Grant End Date | Close Date |
| Report Date | Grant Report Date |
| Report Notes | Grant Report Notes (combined with Notes) |
| Notes | Grant Report Notes (combined with Report Notes) |

### File 2: Contacts spreadsheet

| Column name | What it maps to in the output |
|---|---|
| Foundation | Used to match against Funder (required) |
| First Name | First Name |
| Last Name | Last Name |
| Title | Job Title |
| Email | Email 1 |
| Phone | Phone 1 |
| Address | Address Line 1 |
| (next column after Address) | Address Line 2 |
| (2 columns after Address) | City |
| (3 columns after Address) | State/Province |
| (4 columns after Address) | Zip |

> **Note on address columns:** The city, state, and zip columns in your contacts file may not have header labels in Excel. That is fine. The script automatically looks for them by their position right after the "Address" column.

### A note on dates and amounts

The script will automatically format dates as `mm/dd/yyyy` and dollar amounts to two decimal places (e.g., `50000` becomes `50000.00`). You do not need to reformat these manually beforehand.

---

## Troubleshooting

**"python is not recognized" or "command not found"**
Python is either not installed or not added to your PATH. Reinstall Python from [python.org](https://www.python.org/downloads/) and check "Add Python to PATH" during installation. On Mac, try `python3` instead of `python`.

**"No module named pandas" or similar**
The virtual environment is not activated or dependencies were not installed. Make sure you ran the activation step for your OS before running the script.

**The output file is empty or has very few rows**
The Funder and Foundation values in your two files may not match exactly. Run the script with `--debug-headers` to confirm the column names were detected correctly. Then check that the funder names in File 1 are spelled the same way as the foundation names in File 2.

**"Not found: path/to/file.xlsx"**
The file path you typed is incorrect. Double-check the spelling and make sure the file actually exists at that location. On Windows, you can drag the file directly into the Command Prompt window to get its exact path.

**Script ran but amounts or dates look wrong**
Check that the source columns in your input files are named correctly (see the Input file requirements section above). Use `--debug-headers` to see exactly what column names the script detected.

**Permission error on Windows when activating the virtual environment (PowerShell)**
Run this command first: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, then try activating again.

---

## Output

The merged file will contain the following columns, matching the Neon One import template. Columns that have no corresponding data in either input file will be left blank.

`Prefix`, `First Name`, `Middle Name`, `Last Name`, `Suffix`, `Individual/Company Type`, `IsCompany?`, `Company Name`, `Department`, `Job Title`, `Preferred Name`, `Salutation`, `Deceased`, `Deceased Date`, `Do Not Contact`, `Email Opt Out`, `Email 1`, `Email 2`, `Email 3`, `SMS/MMS Number`, `SMS/MMS Consent`, `Phone 1`, `Phone 1 Type`, `Phone 2`, `Phone 2 Type`, `Phone 3`, `Phone 3 Type`, `Website`, `Fax`, `Address Line 1`, `Address Line 2`, `Address Line 3`, `Address Line 4`, `Address Type`, `City`, `State/Province`, `Territory`, `County`, `Country`, `Zip`, `Birthday`, `Gender`, `Login Name`, `Login Password`, `Note`, `Note Title`, `Note Type`, `Pinned Note?`, `Account Source`, `Volunteer Role(s)`, `Volunteer Group(s)`, `Created By`, `Created Date`, `Last Updated By`, `Last Updated Date`, `Custom Fields`, `Grant Status`, `Grant System User`, `Grant Name`, `Ask Date`, `Ask Amount`, `Funded Date`, `Funded Amount`, `Close Date`, `Grant Note`, `Grant Campaign`, `Grant Fund`, `Grant Purpose`, `Grant Remaining`, `Grant Award Date`, `Grant Report Date`, `Grant Report Notes`

All rows will have `IsCompany?` set to `Yes` and `Country` set to `United States`.
