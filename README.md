# CMCS Valid License Updater

A Python script that automatically downloads and updates valid mining license data from the Mongolian [CMCS (Cadastral Management and Control System)](https://cmcs.mrpam.gov.mn) into local Excel files.

---

## Requirements

- Python 3.8 or higher
- Internet connection to access `https://cmcs.mrpam.gov.mn`

> **Note:** Required Python packages are installed automatically on first run.

### Required Input Files
The following files must be present in the **same folder** as the script before running:

| File | Description |
|------|-------------|
| `old_valid_licences.xlsx` | Previous run's license data (used for comparison) |
| `old_valid_licence_coordinates.xlsx` | Previous run's coordinate data (used for comparison) |

---

## Usage

Simply **double-click** the script file to run it.

Or run it from the terminal:
```bash
python "UpdateCMCS_ValidLicense v2.py"
```

---

## What It Does

1. Checks and installs any missing Python packages
2. Loads existing license data from the input Excel files
3. Logs in to the CMCS system
4. Retrieves all current valid licenses
5. Compares with existing data to identify new licenses
6. Downloads coordinate data for any new licenses found
7. Saves updated data to output Excel files

---

## Output Files

| File | Description |
|------|-------------|
| `valid_licences.xlsx` | Updated license information |
| `valid_licence_coordinates.xlsx` | Updated coordinate data |
| `old_valid_licences.xlsx` | Overwritten as backup for next run |
| `old_valid_licence_coordinates.xlsx` | Overwritten as backup for next run |

---

## Notes

- If no new licenses are found since the last run, the script will exit without making any changes
- Licenses that are no longer valid in the CMCS system are marked as `NotValid` rather than deleted
- Close all output Excel files before running the script to avoid save errors

---

## Author
Byambadorj Mendbayar