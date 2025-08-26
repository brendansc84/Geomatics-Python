import os, re, json, zipfile, sys, time
from pathlib import Path
import win32com.client
import pythoncom
from contextlib import contextmanager

# ---------- Runtime paths (EXE-safe) ----------
def runtime_dir() -> Path:
    # When frozen (PyInstaller), write beside the EXE; otherwise beside the .py
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

RUNTIME_DIR = runtime_dir()
LOCK_PATH = RUNTIME_DIR / "process_sr.lock"                   # lock beside exe/py
LOG_PATH  = RUNTIME_DIR / "processed_sr_entryids.json"        # seen cache beside exe/py

LOCK_MAX_AGE_SECS = 15 * 60  # auto-clear stale lock after 15 minutes

@contextmanager
def single_instance_lock():
    # Clear stale lock if the previous run crashed or was killed
    try:
        if LOCK_PATH.exists():
            age = time.time() - LOCK_PATH.stat().st_mtime
            if age > LOCK_MAX_AGE_SECS:
                print(f"Stale lock ({int(age)}s). Removing {LOCK_PATH.name}.")
                LOCK_PATH.unlink(missing_ok=True)
    except Exception:
        pass

    try:
        fd = os.open(LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_RDWR)
    except FileExistsError:
        print("Another run is in progress. Exiting.")
        raise SystemExit(0)
    try:
        yield
    finally:
        try:
            os.close(fd)
            LOCK_PATH.unlink(missing_ok=True)
        except Exception:
            pass


# ---------- CONFIG ----------
BASE_DIR = r"\\CA0002-PPFSS01\workgroup\1566\active\156660046\Kearl\Survey Requests"
OL_FOLDER_INBOX = 6
MAX_NAME = 150

# SR number anywhere in subject
SR_SUBJECT_PATTERN = re.compile(r"\bSR\d{8,}\b", re.I)
# Replies only (you asked to exclude replies, not forwards)
RE_PREFIX = re.compile(r'^\s*RE\s*(\[\d+\])?\s*:', re.I)


# ---------- Utilities ----------
INVALID_FS_CHARS = r'<>:"/\|?*'
TRANS = str.maketrans({c: " " for c in INVALID_FS_CHARS})

def safe_folder_name(text: str) -> str:
    clean = (text or "").translate(TRANS).strip()
    clean = re.sub(r"\s+", " ", clean)
    # Windows path safety: trim length and trailing dots/spaces
    return clean[:MAX_NAME].rstrip(" .") or "message"

def safe_file_name(name: str) -> str:
    base = Path(name or "attachment").name.translate(TRANS).strip().rstrip(" .")
    base = re.sub(r"\s+", " ", base)
    return (base or "attachment")[:MAX_NAME]

def ensure_dir(path: Path):
    path.mkdir(parents=True, exist_ok=True)

def unique_path(p: Path) -> Path:
    if not p.exists():
        return p
    stem, suffix = p.stem, p.suffix
    i = 2
    while True:
        cand = p.with_name(f"{stem}_{i}{suffix}")
        if not cand.exists():
            return cand
        i += 1

def secure_extract_member(zf: zipfile.ZipFile, member: zipfile.ZipInfo, dest: Path):
    if member.is_dir():
        return None
    fname = Path(member.filename).name  # drop folders in zip
    if not fname:
        return None
    out_path = unique_path(dest / safe_file_name(fname))
    try:
        out_path_resolved = out_path.resolve()
        dest_resolved = dest.resolve()
        if dest_resolved not in out_path_resolved.parents and out_path_resolved != dest_resolved:
            return None
    except Exception:
        pass
    with zf.open(member) as src, open(out_path, "wb") as dst:
        dst.write(src.read())
    return out_path

def load_log() -> set[str]:
    if LOG_PATH.exists():
        try:
            return set(json.loads(LOG_PATH.read_text(encoding="utf-8")))
        except Exception:
            return set()
    return set()

def save_log(seen: set[str]):
    LOG_PATH.write_text(json.dumps(sorted(seen)), encoding="utf-8")

def save_msg(mail, dest_folder: Path, subject: str):
    fname = safe_folder_name(subject) or "message"
    msg_path = unique_path(dest_folder / f"{fname}.msg")
    mail.SaveAs(str(msg_path), 3)  # 3 = .msg
    return msg_path

def save_attachments(mail, dest_folder: Path):
    """
    Save all attachments. If .zip: extract contents to dest and delete the zip.
    Returns list of saved non-zip attachment paths (strings).
    """
    saved = []
    for att in mail.Attachments:
        raw_name = str(att.FileName) or "attachment"
        clean_name = safe_file_name(raw_name)
        save_path = unique_path(dest_folder / clean_name)
        try:
            att.SaveAsFile(str(save_path))
            if clean_name.lower().endswith(".zip"):
                try:
                    with zipfile.ZipFile(save_path, "r") as zf:
                        for member in zf.infolist():
                            secure_extract_member(zf, member, dest_folder)
                except zipfile.BadZipFile:
                    print(f"  ! Corrupt zip (saved but not extracted): {save_path.name}")
                finally:
                    try:
                        save_path.unlink(missing_ok=True)
                    except Exception as e:
                        print(f"  ! Failed to delete zip {save_path.name}: {e}")
            else:
                saved.append(str(save_path))
        except Exception as e:
            print(f"  ! Failed to save attachment {clean_name}: {e}")
    return saved


# ---------- MAIN ----------
def process_folder():
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        folder = outlook.GetDefaultFolder(OL_FOLDER_INBOX)

        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        items = items.Restrict("[Unread] = True")  # minor speed-up

        seen = load_log()
        processed = 0

        for mail in items:
            if getattr(mail, "Class", None) != 43:  # 43 = MailItem
                continue

            try:
                subject = str(mail.Subject)
                entry_id = str(mail.EntryID)
            except Exception:
                continue

            # Must contain an SR####..., and must NOT be a reply (Re:)
            if not SR_SUBJECT_PATTERN.search(subject):
                continue
            if RE_PREFIX.search(subject):
                # Skip replies as requested
                continue

            if entry_id in seen:
                continue

            # Use the FULL subject for the folder name
            dest_name = safe_folder_name(subject)
            dest = Path(BASE_DIR) / dest_name
            ensure_dir(dest)

            print(f"Processing: {subject} -> [{dest_name}]")

            try:
                save_msg(mail, dest, subject)
            except Exception as e:
                print(f"  ! Failed to save .msg: {e}")

            try:
                saved_non_zips = save_attachments(mail, dest)
                if saved_non_zips:
                    print(f"  Saved attachments: {', '.join(Path(p).name for p in saved_non_zips)}")
                else:
                    print("  No non-zip attachments saved (zips extracted & deleted, or no attachments).")
            except Exception as e:
                print(f"  ! Failed to handle attachments: {e}")

            try:
                mail.UnRead = False
                mail.Save()
            except Exception:
                pass

            seen.add(entry_id)
            processed += 1

        save_log(seen)
        print(f"Done. Processed {processed} message(s).")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    with single_instance_lock():
        process_folder()
