# mbox_to_eml_exporter.py
# Export messages from a .mbox file to .eml files with flexible folder layouts and filters.
# Features:
#   - --layout {year,month,flat}: choose directory structure (YYYY, YYYY/MM, or flat)
#   - --start-year / --end-year: export only messages within a year range (based on Date header)
#   - --max-per-dir N: cap number of files per directory (0 = no limit)
#   - --max-dir-bytes GB: cap total size per directory in GB (0 = no limit)
#   - --sanitize-filenames: make file names safe for Windows/macOS
#   - periodic progress logs and final stats
#   - --mbox can be a single .mbox file, a directory (processes all .mbox files), or a .zip file
#
# Example:
#   python mbox_to_eml_exporter.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout month --start-year 2005 --end-year 2016 --max-per-dir 50000 --sanitize-filenames
#   python mbox_to_eml_exporter.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout month --sanitize-filenames
#   python mbox_to_eml_exporter.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout year --start-year 2007 --end-year 2016
#   python mbox_to_eml_exporter.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout month --max-per-dir 50000 --max-dir-bytes 9
#   python mbox_to_eml_exporter.py --mbox "D:\Mail\" --out-dir "D:\Export_EML" --layout month
#   python mbox_to_eml_exporter.py --mbox "D:\Mail\archive.zip" --out-dir "D:\Export_EML" --layout month
#
import argparse
import hashlib
import mailbox
import os
import re
import sys
import tempfile
import time
import zipfile
from email.utils import parsedate_to_datetime

def safe_name(s: str) -> str:
    """Return a filesystem-friendly string (safe for Windows/macOS)."""
    s = re.sub(r'[\\/:*?"<>|]+', '_', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s[:120] or 'msg'

def ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path

def pick_year_month(msg):
    """Extract (year, month) from the Date header; default to (1970, 1) if missing/invalid."""
    year = 1970
    month = 1
    try:
        d = msg.get('Date')
        if d:
            dt = parsedate_to_datetime(d)
            year = dt.year
            month = dt.month
    except Exception:
        pass
    return year, month

def eml_bytes(msg):
    """Return raw EML bytes for a mailbox.mboxMessage."""
    try:
        return msg.as_bytes()
    except Exception:
        # Fallback: encode string representation if needed
        return msg.as_string().encode('utf-8', errors='ignore')

def is_message_read(msg):
    """Check if message has been read based on mbox flags and Gmail labels."""
    try:
        # Check Gmail labels first (most common in Google Takeout)
        gmail_labels = msg.get('X-Gmail-Labels', '')
        if gmail_labels:
            # Gmail "Unread" label in various languages
            # Source: Gmail interface translations
            unread_labels = {
                'unread',           # English
                'non lus',          # French
                'non lues',         # French (plural)
                'da leggere',       # Italian
                'ungelesen',        # German
                'no leído',         # Spanish
                'não lido',         # Portuguese
                'oläst',            # Swedish
                'ulæst',            # Danish
                'ongelezen',        # Dutch
                'непрочитанные',    # Russian
                '未読',             # Japanese
                '未读',             # Chinese Simplified
                '未讀',             # Chinese Traditional
                '읽지 않음',         # Korean
                'okunmamış',        # Turkish
                'غير مقروء',        # Arabic
                'nie przeczytane',  # Polish
                'nelidos',          # Spanish (alt)
                'μη αναγνωσμένα',   # Greek
            }
            
            labels_lower = gmail_labels.lower()
            
            # Check if any unread label variant is present
            for unread_label in unread_labels:
                if unread_label in labels_lower:
                    return False  # Message is unread
            
            # If no unread label found, consider as read
            return True
        
        # mboxMessage has get_flags() method that returns flags like 'RS' (Read, Seen)
        if hasattr(msg, 'get_flags'):
            flags = msg.get_flags()
            # 'S' flag means 'Seen' (read)
            if flags and 'S' in flags:
                return True
        
        # Check Status header (some mbox formats)
        status = msg.get('Status', '')
        if 'R' in status or 'O' in status:  # R=Read, O=Old (read)
            return True
        
        # Check X-Status header
        x_status = msg.get('X-Status', '')
        if 'R' in x_status:
            return True
    except Exception:
        pass
    
    # Default: if no indicators found, assume read (most messages are read)
    return True

def unique_eml_name(idx, msg, is_read=None):
    """Create a unique file name using Message-ID + index + time-based salt + read status."""
    mid = (msg.get('Message-ID') or '').encode('utf-8', errors='ignore')
    h = hashlib.sha1(mid + str(idx).encode() + str(time.time_ns()).encode()).hexdigest()[:12]
    subj = safe_name(msg.get('Subject') or 'no_subject')
    read_suffix = '_READ' if is_read else '_UNREAD'
    return f"{subj}__{h}{read_suffix}.eml"

def fits_limits(count_in_dir, bytes_in_dir, max_per_dir, max_dir_bytes):
    if max_per_dir and count_in_dir >= max_per_dir:
        return False
    if max_dir_bytes and bytes_in_dir >= max_dir_bytes:
        return False
    return True

def find_mbox_files_in_dir(directory):
    """Find all mbox files in a directory (non-recursive)."""
    mbox_files = []
    try:
        for entry in os.listdir(directory):
            full_path = os.path.join(directory, entry)
            if os.path.isfile(full_path):
                # Check for mbox files by extension or common naming
                if entry.lower().endswith('.mbox') or entry.lower().endswith('.mbx') or 'mbox' in entry.lower():
                    mbox_files.append(full_path)
    except Exception as e:
        print(f"[WARN] Error scanning directory {directory}: {e}")
    return mbox_files

def find_mbox_files_in_zip(zip_path):
    """Find all mbox files in a zip archive and return list of (name, is_mbox) tuples."""
    mbox_entries = []
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            for info in zf.infolist():
                if not info.is_dir():
                    name = info.filename
                    # Check for mbox files by extension or common naming
                    if name.lower().endswith('.mbox') or name.lower().endswith('.mbx') or 'mbox' in name.lower():
                        mbox_entries.append(name)
    except Exception as e:
        print(f"[WARN] Error reading zip file {zip_path}: {e}")
    return mbox_entries

def extract_mbox_from_zip(zip_path, mbox_name, temp_dir):
    """Extract a single mbox file from zip to temporary directory."""
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            # Extract to temp directory
            extracted_path = zf.extract(mbox_name, temp_dir)
            return extracted_path
    except Exception as e:
        print(f"[ERROR] Failed to extract {mbox_name} from {zip_path}: {e}")
        return None

def process_single_mbox(mbox_path, out_root, args, dir_counts, dir_bytes, max_dir_bytes):
    """Process a single mbox file and return statistics."""
    print(f"\n{'='*60}")
    print(f"Processing MBOX: {mbox_path}")
    print(f"{'='*60}")
    
    try:
        mb = mailbox.mbox(mbox_path, factory=mailbox.mboxMessage)
        total = len(mb)
        print(f"Messages in MBOX: {total:,}")
    except Exception as e:
        print(f"[ERROR] Failed to open mbox file {mbox_path}: {e}")
        return 0, 0

    exported = 0
    skipped = 0
    started = time.time()
    read_count = 0
    unread_count = 0

    for i, msg in enumerate(mb, 1):
        year, month = pick_year_month(msg)

        # Year filters
        if args.start_year and year < args.start_year:
            skipped += 1
            continue
        if args.end_year and year > args.end_year:
            skipped += 1
            continue

        # Destination directory
        if args.layout == 'flat':
            dest_dir = out_root
        elif args.layout == 'year':
            dest_dir = os.path.join(out_root, f"{year:04d}")
        else:  # 'month'
            dest_dir = os.path.join(out_root, f"{year:04d}", f"{month:02d}")
        ensure_dir(dest_dir)

        # Per-dir limits
        c = dir_counts.get(dest_dir, 0)
        b = dir_bytes.get(dest_dir, 0)

        raw = eml_bytes(msg)
        size = len(raw)

        if not fits_limits(c, b, args.max_per_dir, max_dir_bytes):
            # When limits are reached, create incremental subfolders: e.g., \2007\01__part2, part3, ...
            part = 2
            base_dir = dest_dir
            while True:
                alt = base_dir + f"__part{part}"
                ensure_dir(alt)
                c_alt = dir_counts.get(alt, 0)
                b_alt = dir_bytes.get(alt, 0)
                if fits_limits(c_alt, b_alt, args.max_per_dir, max_dir_bytes):
                    dest_dir = alt
                    c = c_alt
                    b = b_alt
                    break
                part += 1

        # File name with read status
        is_read = is_message_read(msg)
        
        # Debug output for first 10 messages
        if args.debug_read_status and i <= 10:
            print(f"\n--- Message {i} Debug ---")
            print(f"Subject: {msg.get('Subject', 'N/A')[:50]}")
            print(f"X-Gmail-Labels: {msg.get('X-Gmail-Labels', 'N/A')}")
            print(f"Status: {msg.get('Status', 'N/A')}")
            print(f"X-Status: {msg.get('X-Status', 'N/A')}")
            if hasattr(msg, 'get_flags'):
                print(f"Flags: {msg.get_flags()}")
            print(f"Detected as: {'READ' if is_read else 'UNREAD'}")
        
        if is_read:
            read_count += 1
        else:
            unread_count += 1
        
        fname = unique_eml_name(i, msg, is_read)
        if args.sanitize_filenames:
            fname = safe_name(fname)
        dest_path = os.path.join(dest_dir, fname)

        try:
            with open(dest_path, 'wb') as f:
                f.write(raw)
            dir_counts[dest_dir] = c + 1
            dir_bytes[dest_dir]  = b + size
            exported += 1
        except Exception as e:
            print(f"[WARN] Failed to save {dest_path}: {e}")
            skipped += 1

        if args.progress_every and args.progress_every > 0 and i % args.progress_every == 0:
            elapsed = time.time() - started
            rate = exported / elapsed if elapsed > 0 else 0
            print(f"{i:,}/{total:,} | exported: {exported:,} | skipped: {skipped:,} | read: {read_count:,} | unread: {unread_count:,} | {rate:.1f} msg/s")

    elapsed = time.time() - started
    print(f"\nCompleted processing {os.path.basename(mbox_path)}")
    print(f"Exported: {exported:,} | Skipped: {skipped:,}")
    print(f"Read: {read_count:,} | Unread: {unread_count:,}")
    if exported:
        print(f"Average speed: {exported/elapsed:.1f} msg/s")
    
    return exported, skipped

def main():
    ap = argparse.ArgumentParser(description="Export .mbox → .eml with layout and filters")
    ap.add_argument('--mbox', required=True, help='Path to .mbox file, directory containing .mbox files, or .zip file')
    ap.add_argument('--out-dir', required=True, help='Destination directory for .eml files')
    ap.add_argument('--layout', choices=['year','month','flat'], default='year', help='Directory structure (year/month/flat)')
    ap.add_argument('--start-year', type=int, default=None, help='Export from this year (inclusive)')
    ap.add_argument('--end-year', type=int, default=None, help='Export up to this year (inclusive)')
    ap.add_argument('--max-per-dir', type=int, default=0, help='Max files per directory (0 = no limit)')
    ap.add_argument('--max-dir-bytes', type=float, default=0.0, help='Max total size per directory in GB (0 = no limit)')
    ap.add_argument('--sanitize-filenames', action='store_true', help='Sanitize filenames')
    ap.add_argument('--progress-every', type=int, default=1000, help='Log progress every N messages (0 = disabled)')
    ap.add_argument('--debug-read-status', action='store_true', help='Show detailed read/unread detection info for first 10 messages')
    args = ap.parse_args()

    src_path = os.path.normpath(args.mbox)
    out_root = os.path.normpath(args.out_dir)
    
    if not os.path.exists(src_path):
        print(f"Error: Source not found: {src_path}", file=sys.stderr)
        sys.exit(1)
    
    os.makedirs(out_root, exist_ok=True)
    max_dir_bytes = int(args.max_dir_bytes * (1024**3)) if args.max_dir_bytes and args.max_dir_bytes > 0 else 0

    # Determine what we're processing
    mbox_files = []
    temp_dir = None
    
    if os.path.isfile(src_path):
        if src_path.lower().endswith('.zip'):
            # Process zip file
            print(f"Source is a ZIP file: {src_path}")
            mbox_entries = find_mbox_files_in_zip(src_path)
            if not mbox_entries:
                print(f"Error: No mbox files found in zip: {src_path}", file=sys.stderr)
                sys.exit(1)
            print(f"Found {len(mbox_entries)} mbox file(s) in zip")
            
            # Create temporary directory for extraction
            temp_dir = tempfile.mkdtemp(prefix='mbox_export_')
            print(f"Using temporary directory: {temp_dir}")
            
            # Extract all mbox files
            for mbox_name in mbox_entries:
                print(f"Extracting: {mbox_name}")
                extracted = extract_mbox_from_zip(src_path, mbox_name, temp_dir)
                if extracted:
                    mbox_files.append(extracted)
        else:
            # Single mbox file
            print(f"Source is a single mbox file: {src_path}")
            mbox_files = [src_path]
    elif os.path.isdir(src_path):
        # Directory with mbox files
        print(f"Source is a directory: {src_path}")
        mbox_files = find_mbox_files_in_dir(src_path)
        if not mbox_files:
            print(f"Error: No mbox files found in directory: {src_path}", file=sys.stderr)
            sys.exit(1)
        print(f"Found {len(mbox_files)} mbox file(s) in directory")
    else:
        print(f"Error: Invalid source path: {src_path}", file=sys.stderr)
        sys.exit(1)

    # Global directory trackers (shared across all mbox files)
    dir_counts = {}   # path → int
    dir_bytes  = {}   # path → int
    
    total_exported = 0
    total_skipped = 0
    overall_start = time.time()

    try:
        # Process each mbox file
        for idx, mbox_path in enumerate(mbox_files, 1):
            print(f"\n[{idx}/{len(mbox_files)}] Processing: {os.path.basename(mbox_path)}")
            exported, skipped = process_single_mbox(mbox_path, out_root, args, dir_counts, dir_bytes, max_dir_bytes)
            total_exported += exported
            total_skipped += skipped
    finally:
        # Clean up temporary directory if created
        if temp_dir and os.path.exists(temp_dir):
            print(f"\nCleaning up temporary directory: {temp_dir}")
            try:
                import shutil
                shutil.rmtree(temp_dir)
            except Exception as e:
                print(f"[WARN] Failed to remove temporary directory: {e}")

    # Final summary
    overall_elapsed = time.time() - overall_start
    print("\n" + "="*60)
    print("FINAL SUMMARY")
    print("="*60)
    print(f"Processed {len(mbox_files)} mbox file(s)")
    print(f"Total Exported: {total_exported:,} | Total Skipped: {total_skipped:,}")
    print(f"Distinct directories created: {len(dir_counts):,}")
    if total_exported:
        print(f"Overall average speed: {total_exported/overall_elapsed:.1f} msg/s")
        # Top 10 directories by total size
        top = sorted(dir_bytes.items(), key=lambda kv: kv[1], reverse=True)[:10]
        print("\nTop directories by size:")
        for p, bytes_ in top:
            print(f" - {p} → {bytes_/1e9:.2f} GB, {dir_counts.get(p,0)} files")

if __name__ == "__main__":
    main()
