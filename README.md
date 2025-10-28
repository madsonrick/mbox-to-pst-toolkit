# MBOX → EML → PST Toolkit

Two small Python scripts to migrate large mailboxes:
- **MBOX → EML**: exports `.mbox` into structured `.eml` folders (year/month/flat), with filters and per-folder limits.
- **EML → PST**: imports `.eml` into Outlook **PST** using MAPI/COM (pywin32) — supports per-year split, even split by size, PST size cap, live counts, and periodic flush so Windows updates file size.

## Requirements
- Windows with **Microsoft Outlook** installed (M365/2016+).
- Python 3.9+ and `pywin32`.

```bash
pip install -r requirements.txt
# or
pip install pywin32
