# eml_to_pst_import.py
# -----------------------------------------------------------------------------
# Import .EML files into Outlook PST files via MAPI/COM (pywin32), with:
#   - split by YEAR (one PST per year) OR even-split into N PSTs by bytes
#   - max PST size limit (e.g., 15–20 GB) -> auto-rotate to part2, part3...
#   - periodic flush (detach/reattach PST) so Explorer shows file growth
#   - live folder item counts for progress validation
#
# WHY this version works (Drafts fix):
#   We create each MailItem **directly in the PST target folder** using
#   `folder.Items.Add(0)` (0 = olMailItem). This avoids Outlook saving new
#   items into the default profile Drafts folder (which happens with CreateItem).
#
# Tested on: Windows + Outlook (M365 / 2016+). Requires Outlook installed.
#
# Dependencies (install on CMD/PowerShell):
#   pip install pywin32
#
# Outlook prerequisites:
#   1) Have Outlook installed and a default mail profile configured.
#   2) Put Outlook in "Work Offline" while importing (recommended).
#   3) Do NOT keep the target PST open in other apps.
#   4) If you have multiple profiles, ensure the default one opens without
#      prompts; this script calls `ns.Logon("", "", False, True)`.
#
# Usage examples (PowerShell / CMD):
#   # Year-based, limit ~15 GB per PST, flush every 5000 items, progress every 200
#   python eml_to_pst_import_en_v3_3.py ^
#     --src "D:\Export_EML" ^
#     --out-dir "D:\PSTs" ^
#     --base-name emails ^
#     --split-by year ^
#     --max-pst-gb 15 ^
#     --pst-root "Imported (EML)" ^
#     --flush-every 5000 ^
#     --count-every 200
#
#   # Even-split into 6 PSTs by total bytes (no per-year folders):
#   python eml_to_pst_import_en_v3_3.py ^
#     --src "D:\Export_EML" ^
#     --out-dir "D:\PSTs" ^
#     --base-name emails ^
#     --splits 6 ^
#     --max-pst-gb 18 ^
#     --pst-root "Imported (EML)"
#
#   # Use Gmail labels to organize into folders:
#   python eml_to_pst_import_en_v3_3.py ^
#     --src "D:\Export_EML" ^
#     --out-dir "D:\PSTs" ^
#     --base-name emails ^
#     --pst-root "Imported" ^
#     --use-gmail-labels
#
# Notes:
#   - Windows Explorer may show PST as ~256 KB while Outlook has it open.
#     The script closes PSTs at the end (and optionally does periodic flush).
#   - You can point --src either to the root containing many .eml subfolders
#     (e.g., by year/month) or to a specific folder for a single batch.
# -----------------------------------------------------------------------------

import os, sys, time, argparse, mimetypes
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime, getaddresses

try:
    import win32com.client as win32
except Exception:
    print("Error: pywin32 is not installed. Run:  pip install pywin32", file=sys.stderr)
    sys.exit(1)

# --- MAPI named property tags we set for best fidelity -----------------------
PR_TRANSPORT_MESSAGE_HEADERS_A = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
PR_TRANSPORT_MESSAGE_HEADERS_W = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
PR_MESSAGE_DELIVERY_TIME       = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
PR_CLIENT_SUBMIT_TIME          = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
PR_INTERNET_MESSAGE_ID         = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
PR_SENDER_NAME                 = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001E"
PR_SENDER_EMAIL_ADDRESS        = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E"
PR_SENDER_ADDRTYPE             = "http://schemas.microsoft.com/mapi/proptag/0x0C1E001E"
PR_SENT_REPRESENTING_NAME      = "http://schemas.microsoft.com/mapi/proptag/0x0042001E"
PR_SENT_REPRESENTING_EMAIL     = "http://schemas.microsoft.com/mapi/proptag/0x0065001E"
PR_ATTACH_CONTENT_ID           = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
PR_ATTACH_FLAGS                = "http://schemas.microsoft.com/mapi/proptag/0x37140003"

def set_prop(pa, tag, value):
    """Safely set a MAPI property; swallow COM-specific errors."""
    try:
        pa.SetProperty(tag, value)
        return True
    except Exception:
        return False

def addresses_to_str(value):
    """Normalize RFC822 addresses into Outlook-friendly 'Name <addr>; ...' string."""
    if not value:
        return ""
    parts = []
    for name, addr in getaddresses([value]):
        name = (name or "").strip()
        addr = (addr or "").strip()
        if addr and name:
            parts.append(f"{name} <{addr}>")
        elif addr:
            parts.append(addr)
        elif name:
            parts.append(name)
    return "; ".join(parts)

def pick_body(msg):
    """Pick HTML if available; otherwise text/plain; ignore attachment bodies."""
    html = None
    text = None
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp  = (part.get("Content-Disposition") or "").lower()
            if ctype == "text/html" and "attachment" not in disp:
                try:
                    html = part.get_content()
                except:
                    try:
                        html = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="ignore")
                    except:
                        pass
            elif ctype == "text/plain" and "attachment" not in disp:
                try:
                    text = part.get_content()
                except:
                    try:
                        text = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="ignore")
                    except:
                        pass
    else:
        ctype = msg.get_content_type()
        if ctype == "text/html":
            try:
                html = msg.get_content()
            except:
                try:
                    html = msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8", errors="ignore")
                except:
                    pass
        else:
            try:
                text = msg.get_content()
            except:
                try:
                    text = msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8", errors="ignore")
                except:
                    pass
    return html, text

def add_attachments(mail, msg, tmpdir):
    """Save each attachment to a temp file and add via Outlook Attachments.Add."""
    count = 0
    for part in msg.walk():
        disp = (part.get("Content-Disposition") or "").lower()
        if "attachment" in disp or (part.get_filename() and part.get_content_maintype() != "text"):
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            fn = part.get_filename()
            if not fn:
                ext = mimetypes.guess_extension(part.get_content_type() or "") or ".bin"
                fn = f"attachment{count+1}{ext}"
            p = os.path.join(tmpdir, f"att_{time.time_ns()}_{fn}")
            try:
                with open(p, "wb") as f:
                    f.write(payload)
                a = mail.Attachments.Add(p)
                cid = part.get("Content-ID")
                if cid:
                    cid = cid.strip("<>")
                    try:
                        a.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)
                        a.PropertyAccessor.SetProperty(PR_ATTACH_FLAGS, 0)
                    except Exception:
                        pass
                count += 1
            except Exception:
                pass
            finally:
                try:
                    os.remove(p)
                except Exception:
                    pass
    return count

def build_headers_text(raw_bytes):
    """Extract raw header block (up to the first blank line) for PR_TRANSPORT_MESSAGE_HEADERS."""
    try:
        end = raw_bytes.find(b"\n\n")
        if end == -1:
            end = min(len(raw_bytes), 1024 * 128)
        hdr = raw_bytes[:end]
        return hdr.decode("utf-8", errors="ignore")
    except Exception:
        return ""

def extract_read_status_from_filename(filename):
    """Extract read/unread status from EML filename suffix. Returns True if READ, False if UNREAD, None if unknown."""
    filename_upper = filename.upper()
    if '_READ.EML' in filename_upper:
        return True
    elif '_UNREAD.EML' in filename_upper:
        return False
    return None

def parse_gmail_labels(msg):
    """
    Extract Gmail labels from X-Gmail-Labels header.
    Returns a list of label strings (empty list if not present).
    """
    labels_header = msg.get("X-Gmail-Labels", "")
    if not labels_header:
        return []
    # Gmail labels are comma-separated
    labels = [lbl.strip() for lbl in labels_header.split(",") if lbl.strip()]
    return labels

def ensure_folder_path(root_folder, folder_path):
    """
    Ensure a folder path exists under root_folder.
    folder_path can be a simple name or a slash-separated path like "Parent/Child".
    Returns the deepest folder.
    """
    parts = [p.strip() for p in folder_path.split("/") if p.strip()]
    current = root_folder
    for part in parts:
        try:
            current = current.Folders.Item(part)
        except Exception:
            current = current.Folders.Add(part)
    return current

def create_mail_in_dest(dest_folder, msg, raw_headers_text, is_read=None):
    """
    Create MailItem directly in the target PST folder (avoids default Drafts).
    Return unsent MailItem (not submitted), already associated with `dest_folder`.
    """
    mail = dest_folder.Items.Add(0)  # 0 = olMailItem
    pa   = mail.PropertyAccessor
    
    # Set read/unread status if provided
    if is_read is not None:
        try:
            mail.UnRead = not is_read  # UnRead=True means unread, UnRead=False means read
        except Exception:
            pass

    # Raw headers first (helps with PR_TRANSPORT_MESSAGE_HEADERS)
    if raw_headers_text:
        if not set_prop(pa, PR_TRANSPORT_MESSAGE_HEADERS_W, raw_headers_text):
            set_prop(pa, PR_TRANSPORT_MESSAGE_HEADERS_A, raw_headers_text)

    # Basic fields
    mail.Subject = msg.get("Subject", "") or ""
    mail.To  = addresses_to_str(msg.get("To", ""))
    mail.CC  = addresses_to_str(msg.get("Cc", ""))
    mail.BCC = addresses_to_str(msg.get("Bcc", ""))

    # Sender metadata (improves fidelity in Outlook item inspector)
    from_list = getaddresses([msg.get("From", "")])
    if from_list:
        name, addr = (from_list[0][0] or ""), (from_list[0][1] or "")
        set_prop(pa, PR_SENDER_NAME, name or addr or "")
        set_prop(pa, PR_SENDER_EMAIL_ADDRESS, addr or "")
        set_prop(pa, PR_SENDER_ADDRTYPE, "SMTP")
        set_prop(pa, PR_SENT_REPRESENTING_NAME, name or addr or "")
        set_prop(pa, PR_SENT_REPRESENTING_EMAIL, addr or "")

    mid = msg.get("Message-ID")
    if mid:
        set_prop(pa, PR_INTERNET_MESSAGE_ID, mid)

    # Dates
    sent_dt = None
    if msg.get("Date"):
        try:
            sent_dt = parsedate_to_datetime(msg.get("Date"))
        except Exception:
            pass
    if sent_dt:
        try:
            mail.SentOn = sent_dt
        except Exception:
            pass
        set_prop(pa, PR_MESSAGE_DELIVERY_TIME, sent_dt)
        set_prop(pa, PR_CLIENT_SUBMIT_TIME, sent_dt)

    # Body
    html, text = pick_body(msg)
    if html:
        mail.HTMLBody = html
    elif text:
        mail.Body = text
    else:
        mail.Body = ""

    return mail

def ensure_outlook_with_logon():
    """Start Outlook.Application and log on to the default MAPI profile."""
    app = win32.Dispatch("Outlook.Application")
    ns  = app.GetNamespace("MAPI")
    ns.Logon("", "", False, True)  # default profile (no UI)
    return app, ns

def normcasepath(p): 
    return os.path.normcase(os.path.normpath(p))

def find_store_by_path(ns, pst_path):
    """Return an Outlook Store matched by FilePath, else None."""
    target = normcasepath(pst_path)
    for store in ns.Stores:
        try:
            current = normcasepath(store.FilePath)
            if current == target:
                return store
        except Exception:
            pass
    return None

def create_or_attach_pst(ns, desired_path):
    """
    Create and attach a PST at desired_path; return (store, root_folder, actual_path).
    Outlook sometimes creates PST elsewhere (e.g., "Outlook Files"); we detect that
    and return the actual path.
    """
    import time as _time
    desired_path = os.path.normpath(desired_path)
    out_dir = os.path.dirname(desired_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    # Remove if already attached or file exists
    exist = find_store_by_path(ns, desired_path)
    if exist:
        try:
            ns.RemoveStore(exist.GetRootFolder())
        except Exception:
            pass
    if os.path.exists(desired_path):
        try:
            os.remove(desired_path)
        except Exception:
            pass

    actual = desired_path
    try:
        ns.AddStoreEx(actual, 2)  # 2 = Unicode PST
    except Exception:
        ns.AddStore(actual)

    # Find the Store that matches the actual path we requested
    for _ in range(20):
        for store in ns.Stores:
            try:
                if normcasepath(store.FilePath) == normcasepath(actual):
                    return store, store.GetRootFolder(), actual
            except Exception:
                pass
        _time.sleep(0.25)

    # Fallback: if Outlook created the PST in a default location, return the last added
    last = None
    for s in ns.Stores:
        last = s
    if last:
        return last, last.GetRootFolder(), last.FilePath

    raise RuntimeError(f"Failed to create/attach PST: {desired_path}")

class PstRouter:
    """
    Decide which PST (Store/Root) should receive the next item:
      - by year: one PST per year (auto-rotate if > max size)
      - by even split: target_bytes computed from total size / splits
      - always tracks approximate bytes to trigger rotation
    """
    def __init__(self, ns, out_dir, base_name, max_gb, splits, split_by_year):
        self.ns = ns
        self.out_dir = os.path.normpath(out_dir)
        self.base = base_name
        self.max_bytes = int(max_gb * (1024**3))
        self.splits = splits
        self.by_year = split_by_year
        self.cur = {"store": None, "root": None, "path": None, "bytes": 0, "part": 0}
        self.year = {}
        self.target_bytes = None
        self.used_stores = []

    def set_total_bytes(self, total_bytes):
        # For even-split mode: divide total payload by N
        if self.splits and self.splits > 0:
            self.target_bytes = max(1, total_bytes // self.splits)

    def _new_path(self, suffix):
        return os.path.join(self.out_dir, f"{self.base}_{suffix}.pst")

    def _rotate_new(self):
        self.cur["part"] += 1
        desired = self._new_path(f"part{self.cur['part']}")
        store, root, actual = create_or_attach_pst(self.ns, desired)
        self.cur.update({"store": store, "root": root, "path": actual, "bytes": 0})
        self.used_stores.append(store)
        print(f"\nUSING PST: {actual}")
        return store, root, actual

    def _need_rotate_seq(self, incoming):
        if self.cur["store"] and (self.cur["bytes"] + incoming > self.max_bytes):
            return True
        if self.target_bytes and self.cur["store"] and (self.cur["bytes"] >= self.target_bytes):
            return True
        return False

    def _ensure_year(self, year, incoming):
        s = self.year.get(year)
        if s is None:
            desired = self._new_path(f"{year}_part1")
            store, root, actual = create_or_attach_pst(self.ns, desired)
            s = {"path": actual, "store": store, "root": root, "bytes": 0, "part": 1}
            self.year[year] = s
            self.used_stores.append(store)
            print(f"\nUSING PST ({year}): {actual}")
        elif s["bytes"] + incoming > self.max_bytes:
            s["part"] += 1
            desired = self._new_path(f"{year}_part{s['part']}")
            store, root, actual = create_or_attach_pst(self.ns, desired)
            s.update({"path": actual, "store": store, "root": root, "bytes": 0})
            self.used_stores.append(store)
            print(f"\nROTATED PST ({year}) → {actual}")
        return s

    def route(self, approx_size, year_hint=None):
        """Return (store, root, path) to drop the next item."""
        size = approx_size
        if self.by_year:
            y = year_hint or 1970
            s = self._ensure_year(y, size)
            s["bytes"] += size
            return s["store"], s["root"], s["path"]
        # even-split (no year)
        if self.cur["store"] is None or self._need_rotate_seq(size):
            self._rotate_new()
        self.cur["bytes"] += size
        return self.cur["store"], self.cur["root"], self.cur["path"]

def list_eml_files(root_dir):
    """Return [(path, size)], total_bytes – recursively collects *.eml."""
    files = []
    total_bytes = 0
    for base, _, fnames in os.walk(root_dir):
        for fn in fnames:
            if fn.lower().endswith(".eml"):
                p = os.path.join(base, fn)
                try:
                    sz = os.path.getsize(p)
                except OSError:
                    sz = 0
                files.append((p, sz))
                total_bytes += sz
    files.sort()
    return files, total_bytes

def main():
    ap = argparse.ArgumentParser(description="Import EML → PST (Outlook/MAPI) with split & progress")
    ap.add_argument("--src", required=True, help="Root folder containing .eml files (recursively)")
    ap.add_argument("--out-dir", required=True, help="Directory to write .pst files")
    ap.add_argument("--base-name", default="emails", help="Base name for PST files")
    ap.add_argument("--split-by", choices=["year"], default=None, help="Split by mail year (one PST per year)")
    ap.add_argument("--splits", type=int, default=None, help="Even-split into N PSTs by total bytes")
    ap.add_argument("--max-pst-gb", type=float, default=15.0, help="Max ~GB per PST before rotating part2/part3...")
    ap.add_argument("--pst-root", default="Imported (EML)", help="Folder name inside each PST (or root when using Gmail labels)")
    ap.add_argument("--flush-every", type=int, default=0, help="Detach/reattach PST every N items (0=off)")
    ap.add_argument("--count-every", type=int, default=200, help="Print folder Items.Count every N items (0=off)")
    ap.add_argument("--use-gmail-labels", action="store_true", help="Use X-Gmail-Labels header to organize into folders; duplicates email if multiple labels")
    args = ap.parse_args()

    src = os.path.normpath(args.src)
    out_dir = os.path.normpath(args.out_dir)
    os.makedirs(out_dir, exist_ok=True)

    files, total_bytes = list_eml_files(src)
    if not files:
        print("No .eml files found.", file=sys.stderr)
        sys.exit(1)
    print(f"EML files: {len(files):,} | Total size: {total_bytes/1e9:,.2f} GB")

    app, ns = ensure_outlook_with_logon()

    router = PstRouter(
        ns=ns,
        out_dir=out_dir,
        base_name=args.base_name,
        max_gb=args.max_pst_gb,
        splits=args.splits,
        split_by_year=(args.split_by == "year"),
    )
    router.set_total_bytes(total_bytes)

    parser = BytesParser(policy=policy.default)
    start = time.perf_counter()
    processed = 0
    done_bytes = 0
    tmpdir = os.path.join(out_dir, "_tmp_eml_to_pst")
    os.makedirs(tmpdir, exist_ok=True)
    current_pst = None

    try:
        for (path, sz) in files:
            # Read raw bytes & parse message
            try:
                with open(path, "rb") as f:
                    raw = f.read()
            except Exception as e:
                print(f"\n[WARN] Could not read {path}: {e}")
                continue

            try:
                msg = parser.parsebytes(raw)
            except Exception:
                # Minimal fallback if parsing fails
                msg = parser.parsebytes(b"Subject: (no subject)\r\n\r\n")

            # Year hint (used when --split-by year)
            year = 1970
            try:
                if msg.get("Date"):
                    year = parsedate_to_datetime(msg.get("Date")).year
            except Exception:
                pass

            # Choose PST destination
            store, root, pst_path = router.route(sz or len(raw), year)
            if pst_path != current_pst:
                current_pst = pst_path
                print(f"\nCURRENT PST: {current_pst}")

            # Extract read status from filename
            is_read = extract_read_status_from_filename(os.path.basename(path))
            
            # Determine target folders based on Gmail labels or default pst-root
            target_folders = []
            if args.use_gmail_labels:
                gmail_labels = parse_gmail_labels(msg)
                if gmail_labels:
                    # Create folders based on Gmail labels (under pst_root)
                    for label in gmail_labels:
                        try:
                            # Ensure base folder exists
                            try:
                                base_folder = root.Folders.Item(args.pst_root)
                            except Exception:
                                base_folder = root.Folders.Add(args.pst_root)
                            
                            # Create/get the label folder (supports nested paths like "Parent/Child")
                            label_folder = ensure_folder_path(base_folder, label)
                            target_folders.append(label_folder)
                        except Exception as e:
                            print(f"\n[WARN] Failed to create folder for label '{label}': {e}")
                else:
                    # No Gmail labels found, use default pst-root
                    try:
                        dest = root.Folders.Item(args.pst_root)
                    except Exception:
                        dest = root.Folders.Add(args.pst_root)
                    target_folders.append(dest)
            else:
                # Standard mode: use pst-root
                try:
                    dest = root.Folders.Item(args.pst_root)
                except Exception:
                    dest = root.Folders.Add(args.pst_root)
                target_folders.append(dest)

            # Create item in first target folder, then copy references to others
            if target_folders:
                try:
                    # Create the email in the first folder
                    first_dest = target_folders[0]
                    item = create_mail_in_dest(first_dest, msg, build_headers_text(raw), is_read)
                    add_attachments(item, msg, tmpdir)
                    item.Save()

                    # Safety: ensure the item actually resides in the target folder
                    try:
                        if item.Parent and item.Parent.EntryID != first_dest.EntryID:
                            item = item.Move(first_dest)
                            item.Save()
                    except Exception:
                        pass

                    # For additional folders, copy the item (creates a reference, not a duplicate)
                    for idx in range(1, len(target_folders)):
                        try:
                            dest = target_folders[idx]
                            item.Copy().Move(dest)
                        except Exception as e:
                            print(f"\n[WARN] Failed to copy item from {path} to folder #{idx+1}: {e}")

                except Exception as e:
                    print(f"\n[WARN] Failed to save item from {path}: {e}")

            processed += 1
            done_bytes += (sz or len(raw))

            # Periodic live count
            if args.count_every and processed % args.count_every == 0:
                try:
                    fresh_dest = root.Folders.Item(args.pst_root)
                    print(f"\n  Items in '{args.pst_root}': {fresh_dest.Items.Count}")
                except Exception:
                    pass

            # Periodic flush: close & reopen the current PST so Explorer shows growth
            if args.flush_every and processed % args.flush_every == 0:
                try:
                    ns.RemoveStore(root)
                    store, root, pst_path = create_or_attach_pst(ns, pst_path)
                    print(f"\n[FLUSH] Re-attached PST: {pst_path}")
                except Exception as e:
                    print(f"\n[WARN] Periodic flush failed: {e}")

            # Console progress & ETA
            if processed % 100 == 0 or processed == len(files):
                pct = (done_bytes / total_bytes) * 100.0
                elapsed = time.perf_counter() - start
                rate = done_bytes / max(elapsed, 1e-9)
                eta = (total_bytes - done_bytes) / max(rate, 1e-9)
                h, m, s = int(eta // 3600), int(eta % 3600 // 60), int(eta % 60)
                print(f"\r{processed:,}/{len(files):,} ({pct:7.3f}%) | ETA {h:02d}:{m:02d}:{s:02d} | PST: {current_pst}", end="")

        print("\nDone.")
        print(f"Imported items: {processed:,}")
        print("PST(s) written to:", out_dir)

    finally:
        # Final flush: remove all open stores so Outlook writes PSTs to disk
        try:
            for s in getattr(router, "year", {}).values():
                try:
                    ns.RemoveStore(s["root"])
                except Exception:
                    pass
            if getattr(router, "cur", {}).get("root"):
                try:
                    ns.RemoveStore(router.cur["root"])
                except Exception:
                    pass
        except Exception as e:
            print("\n[WARN] Failed to close stores:", e)
        try:
            app.Quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()
