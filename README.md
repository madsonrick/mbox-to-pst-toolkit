# MBOX → EML → PST Toolkit

Export large `.mbox` mailboxes into filesystem `.eml` files and import them into Outlook **.pst** archives — with year/month folder layouts, filters, size limits, progress logs, and safe Outlook integration.

- **Step 1:** `mbox_to_eml_exporter_en_v2.py` → MBOX → EML  
- **Step 2:** `eml_to_pst_import_en_v3_3.py` → EML → PST (Outlook / pywin32)

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
