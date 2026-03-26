# IAW Data Formatting

Small utilities for preparing spreadsheets (for example, grant and contact data migration).

## Merge grants and contacts (`merge_grant_contacts.py`)

Combines two Excel workbooks—grants (file 1) and contacts (file 2)—into one sheet with a fixed column layout. Rows are matched where **Funder** equals **Foundation** (case-insensitive, trimmed).

### Requirements

- Python 3.9 or newer (3.10+ recommended)

### Setup with a virtual environment

From the project root:

**macOS / Linux**

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

**Windows (Command Prompt)**

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
pip install --upgrade pip
pip install -r requirements.txt
```

**Windows (PowerShell)**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install -r requirements.txt
```

The `.venv` directory is listed in `.gitignore` and should not be committed.

To leave the virtual environment when you are done:

```bash
deactivate
```

### Run the merge

With the venv activated:

```bash
python merge_grant_contacts.py path/to/grants.xlsx path/to/contacts.xlsx -o path/to/output.xlsx
```

If you omit `-o`, the script writes `merged_export.xlsx` in the current working directory.

### Input expectations

- **File 1 (grants):** columns such as Funder, Purpose, Amount, Received, Remaining, Award Date, Grant Start Date, Grant End Date, Report Date, Report Notes, Notes.
- **File 2 (contacts):** First Name, Last Name, Foundation, Title, email, Phone, plus Address and the following columns (line 2, city, state, zip may have blank headers in Excel).

See the script docstring and comments in `merge_grant_contacts.py` for the full output column list and field mapping.
