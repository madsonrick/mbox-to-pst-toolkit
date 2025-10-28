# MBOX → EML → PST Toolkit

Export large `.mbox` mailboxes into filesystem `.eml` files and import them into Outlook **.pst** archives — with year/month folder layouts, filters, size limits, progress logs, and safe Outlook integration.

- **Step 1:** `mbox_to_eml_exporter.py` → MBOX → EML  
- **Step 2:** `eml_to_pst_import.py` → EML → PST (Outlook / pywin32)

> Designed for multi-tens-of-GB mailboxes. Avoids the “Drafts” pitfall by creating items **directly** in the PST folder.

---

## Features

**MBOX → EML**
- Layout: **year**, **month** (YYYY/MM), or **flat**
- Year range filters (`--start-year`, `--end-year`)
- Per-directory caps: **max files** or **max GB**
- Filename sanitization for Windows/macOS
- Periodic progress logs

**EML → PST (Outlook)**
- Split by **year** (one PST per year) or even-split into **N PSTs by bytes**
- **Max PST size** (e.g., 15–20 GB) with auto-rotate to `…_part2.pst`, `…_part3.pst`
- Periodic **flush** (detach/reattach) so Windows shows file growth
- Live folder **item counts** for validation
- Creates items **directly in the PST folder** (avoids default Drafts)

---

## Requirements

- **Windows** with **Microsoft Outlook** (M365/2019/2016+)
- **Python 3.9+**
- Python package: `pywin32`

```bash
pip install pywin32

---

## Outlook tips
- Put Outlook in Work Offline during imports.
- Ensure a default Outlook profile opens without prompts.
- Don’t browse/move items inside the target PST while the script runs.
- If Windows Explorer shows PST ~256 KB while attached, that’s normal; the script detaches PSTs at the end (and on --flush-every) so the OS updates size.

---

## Quick Start

D:\Mail\inbox.mbox          # your source mailbox
D:\Export_EML               # where .eml files will be written
D:\PSTs                     # where .pst files will be created

---

## 1) Export MBOX → EML
Script: mbox_to_eml_exporter.py

python mbox_to_eml_exporter.py ^
  --mbox "D:\Mail\inbox.mbox" ^
  --out-dir "D:\Export_EML" ^
  --layout month ^
  --sanitize-filenames

##Common layouts
--layout year → Export_EML\YYYY\*.eml
--layout month → Export_EML\YYYY\MM\*.eml
--layout flat → Export_EML\*.eml

##Year filters
:: Only 2007–2016
python mbox_to_eml_exporte.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout year --start-year 2007 --end-year 2016

##Folder limits
:: Cap folders at 50k files OR ~9 GB
python mbox_to_eml_exporter.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout month --max-per-dir 50000 --max-dir-bytes 9

##Flags (summary)
--layout {year,month,flat}
--start-year N, --end-year N
--max-per-dir N (files)
--max-dir-bytes GB (gigabytes)
--sanitize-filenames
--progress-every N

##2) Import EML → PST (Outlook)

Script: eml_to_pst_import.py

Split by year, cap 15 GB per PST

python eml_to_pst_import.py ^
  --src "D:\Export_EML" ^
  --out-dir "D:\PSTs" ^
  --base-name emails ^
  --split-by year ^
  --max-pst-gb 15 ^
  --pst-root "Imported (EML)" ^
  --flush-every 5000 ^
  --count-every 200

##Even-split into N PSTs by total size (no per-year split)

python eml_to_pst_import.py ^
  --src "D:\Export_EML" ^
  --out-dir "D:\PSTs" ^
  --base-name emails ^
  --splits 6 ^
  --max-pst-gb 18 ^
  --pst-root "Imported (EML)"

Useful flags (summary)
--split-by year or --splits N
--max-pst-gb 15 (rotate to part2 when exceeded)
--pst-root "Imported (EML)" (folder inside each PST)
--flush-every N (forces Explorer to update file size periodically)
--count-every N (prints Outlook folder item counts)

Why PST size looks “stuck” at ~256 KB?
While Outlook holds the PST open, Windows Explorer may not refresh its size. This script detaches PSTs at the end (and optionally during --flush-every), forcing size updates.
You can also check inside Outlook → folder Properties → Folder Size….

##FAQ
Q: Can I import only one year at a time?
A: Yes. Point --src directly to a year folder, e.g., D:\Export_EML\2007.

Q: What PST size should I use?
A: Keep --max-pst-gb around 15–20 GB for stability; the script will auto-rotate to …_part2.pst.

Q: I see items in Drafts instead of the PST.
A: Use this importer v3.3 — it creates items directly in the PST folder (Items.Add(0)), not via CreateItem.

Q: Outlook shows “Server not available” or prompts for profile.
A: Ensure a default profile opens without UI. Use Work Offline during imports.

#Troubleshooting
Items appear in “Drafts”
You’re likely on an older importer version; (direct creation in PST folder).

Explorer file size doesn’t change
Normal while PST is attached. Increase --flush-every for more frequent detach/reattach, or just rely on the final close.

##Performance tips
Use SSDs for both source and destination if possible.
Increase --count-every to reduce console overhead (e.g., --count-every 1000).
Lower --flush-every only if you need to see file growth during the run; flushing too often slows things down.
