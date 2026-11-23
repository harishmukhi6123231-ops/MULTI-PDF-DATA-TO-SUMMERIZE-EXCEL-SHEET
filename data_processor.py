"""
Transit Pass Retriever
----------------------
Give it a list of required Pass Nos and a folder of PDFs; it COPIES the
matching passes into a destination folder (originals stay untouched) and
keeps ONE report in the DESTINATION folder: retrieval_report.xlsx

  Sheet "All Required" - every required pass no, checked against
                         pass_summary.xlsx (date, quantity, Used / Not used),
                         whether its PDF was found, where, and remarks
  Sheet "Not Found"    - the required passes whose PDF is still not found
  Sheet "Run History"  - one line for every time the program was run

SEARCHING MORE THAN ONE FOLDER: when some passes are not found, the program asks
whether to search ANOTHER folder for just those missing passes, and keeps asking
until everything is found or you say no. Everything found goes to the same
destination and into the same report.

Run it again later (e.g. on another source folder) with the SAME destination
and the same report is UPDATED, not duplicated:
  - a pass found earlier stays "found" even if it is not in the new folder
  - a pass that was "not found" before is updated when this run finds it
  - passes not seen before are added; if a pass is found again in another
    folder, the extra file names are added to its row

Pass list can be (or, when the destination already has a report with passes
still not found, you are offered those passes as the list):
  - an Excel file (uses the column headed "Pass No", else the first column;
    rows hidden by a filter are skipped, so you can filter pass_summary.xlsx
    and use it directly)
  - a CSV / TXT file (one pass no per line, or separated by commas)
  - typed / pasted in the window if you skip the file selection

Matching: first use the filename recorded in pass_summary.xlsx. If a required
pass is listed there, the program locates that PDF directly and does not read
its contents. Only passes without a usable summary filename use filename
matching and, as a last resort, PDF text reading.

Install:  pip install pdfplumber openpyxl
Run:      python data_processor.py
"""

import os
import re
import shutil
from datetime import datetime
from pathlib import Path

import pdfplumber
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PASS_RE = re.compile(r"Pass\s*No\.?\s*[:\-]?\s*([A-Z0-9/]+)", re.I)
USED_RE = re.compile(r"[\s_\-.(\[](?:U|USED)[)\]]?\s*$", re.I)
DEFAULT_DEST = "Retrieved Passes"
SUMMARY_NAME = "TP Summary.xlsx"
REPORT_NAME = "retrieval_report.xlsx"

RED = PatternFill("solid", fgColor="FFC7CE")      # PDF not found
ORANGE = PatternFill("solid", fgColor="F8CBAD")   # already used
YELLOW = PatternFill("solid", fgColor="FFEB9C")   # needs checking


# ---------- small helpers ----------
def norm(pass_no: str) -> str:
    """A2603070037/2/539  ->  A2603070037-2-539 (uppercase, no spaces)."""
    return re.sub(r"\s+", "", str(pass_no)).upper().replace("/", "-")


def choose_folder(title: str):
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        folder = filedialog.askdirectory(title=title)
        root.destroy()
        return Path(folder) if folder else None
    except Exception:
        typed = input(f"{title}\nEnter path (blank to skip): ").strip().strip('"')
        return Path(typed) if typed else None


def choose_file(title: str, patterns="*.xlsx *.xlsm *.csv *.txt"):
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        f = filedialog.askopenfilename(title=title, filetypes=[("Files", patterns), ("All files", "*.*")])
        root.destroy()
        return Path(f) if f else None
    except Exception:
        typed = input(f"{title}\nEnter path (blank to skip): ").strip().strip('"')
        return Path(typed) if typed else None


def clean_list(values):
    """Keep entries containing a digit, remove duplicates, keep order."""
    seen, out = set(), []
    for v in values:
        v = str(v).strip()
        if not v or not re.search(r"\d", v):
            continue
        if norm(v) not in seen:
            seen.add(norm(v))
            out.append(v)
    return out


def read_pass_list(path: Path):
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        ws = load_workbook(path, data_only=True).active
        col = 0
        for row in ws.iter_rows(min_row=1, max_row=5):
            for i, c in enumerate(row):
                if isinstance(c.value, str) and re.search(r"pass\s*no", c.value, re.I):
                    col = i
        values = []
        for row in ws.iter_rows():
            if len(row) <= col:
                continue
            cell = row[col]
            if ws.row_dimensions[cell.row].hidden or cell.value is None:
                continue
            values.append(cell.value)
        return clean_list(values)

    try:
        text = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        text = path.read_text(encoding="cp1252")
    return clean_list(re.split(r"[,;\t\r\n]+", text))


def type_pass_list():
    print("Type or paste the pass numbers (one per line, or separated by commas).")
    print("Press Enter on an empty line when finished.\n")
    lines = []
    while True:
        line = input("> ").strip()
        if not line:
            break
        lines.append(line)
    return clean_list(re.split(r"[,;\t\r\n]+", "\n".join(lines)))


# ---------- pass summary ----------
def read_summary(path: Path):
    """Read pass_summary.xlsx without modifying it."""
    ws = load_workbook(path, data_only=True, read_only=True).active
    header = [str(c.value or "").strip().lower() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    def find_any(keys, default=None):
        return next((i for i, h in enumerate(header) if any(k in h for k in keys)), default)
    c_date = find_any(("date",), 0)
    c_qty = find_any(("quant", "qty"), 1)
    c_pass = find_any(("pass no", "pass number", "pass"), 2)
    c_status = find_any(("status", "used"), 3)
    c_filename = find_any(("file name", "filename", "pdf name", "document name", "source file"))
    def cell(row, i):
        return row[i] if i is not None and len(row) > i else None
    data = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        pn = cell(row, c_pass)
        if pn is None or not str(pn).strip():
            continue
        filename = cell(row, c_filename)
        data.setdefault(norm(pn), []).append({
            "date": cell(row, c_date), "qty": cell(row, c_qty),
            "status": str(cell(row, c_status) or "").strip(),
            "filename": str(filename).strip() if filename else "",
        })
    ws.parent.close()
    return data


# ---------- matching ----------
def find_by_name(pass_no: str, pdfs):
    pat = re.compile(r"(?<![A-Z0-9])" + re.escape(norm(pass_no)) + r"(?![A-Z0-9])")
    # keep spaces in the file name so "...-539 U" is not read as "...-539U"
    return [p for p in pdfs if pat.search(p.stem.upper().replace("/", "-"))]


def canonical_file_name(path_or_name):
    stem = Path(str(path_or_name)).stem
    stem = re.sub(r"(?i)[\s_().-]*(?:not[\s_.-]*used|unused|used|u|nu)$", "", stem)
    return re.sub(r"[^A-Z0-9]", "", stem.upper())


def display_file_name(path_or_name):
    p = Path(str(path_or_name))
    stem = re.sub(r"(?i)[\s_().-]*(?:not[\s_.-]*used|unused|used|u|nu)$", "", p.stem)
    return stem.rstrip(" _.-") + p.suffix


def is_marked_used(path_or_name):
    """True only for a positive Used suffix, not 'Not Used' or 'Unused'."""
    stem = Path(str(path_or_name)).stem.lower()
    compact = re.sub(r"[\s_.()-]+", "", stem)
    if compact.endswith(("notused", "unused", "nu")):
        return False
    return compact.endswith(("used", "u"))


def dedupe_files(files, preferred_status=""):
    chosen = {}
    want_used = str(preferred_status).strip().lower() == "used"
    for p in files:
        key = canonical_file_name(p.name)
        if key not in chosen:
            chosen[key] = p
            continue
        old = chosen[key]
        p_used, old_used = is_marked_used(p.name), is_marked_used(old.name)
        if (want_used and p_used and not old_used) or (not want_used and old_used and not p_used):
            chosen[key] = p
    return sorted(chosen.values(), key=lambda p: p.name.casefold())


def index_summary_files(summary, folder, recursive):
    pdfs = find_pdfs(folder, recursive, None)
    by_name = {}
    for p in pdfs:
        by_name.setdefault(p.name.casefold(), p)
        by_name.setdefault(p.stem.casefold(), p)
    for rows in summary.values():
        for row in rows:
            filename = row.get("filename", "")
            if filename:
                key = Path(filename).name.casefold()
                resolved = by_name.get(key) or by_name.get(Path(filename).stem.casefold())
                if resolved:
                    row["_path"] = resolved


def read_pass_from_pdf(pdf_path: Path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ""
        m = PASS_RE.search(text)
        return norm(m.group(1)) if m else None
    except Exception:
        return None


# ---------- analysis + report ----------
def analyse(wanted, matches, summary):
    """Compare each required pass with the PDFs found and the pass summary."""
    results = []
    for w in wanted:
        rows = summary.get(norm(w), []) if summary is not None else None
        preferred_status = next((r["status"] for r in rows if r.get("status")), "") if rows else ""
        files = dedupe_files(matches[w], preferred_status)
        file_used = any(is_marked_used(p.name) for p in files)
        sum_used = any(r["status"].lower() == "used" for r in rows) if rows else False
        status_known = bool(rows) and any(r["status"] for r in rows)

        remarks, level = [], 0          # level: 0 ok, 1 check, 2 used, 3 not found
        if not files:
            remarks.append("PDF not found")
            level = 3
        if summary is not None and not rows:
            remarks.append("Not in pass summary")
            level = max(level, 1)
        if files and rows and status_known and file_used != sum_used:
            remarks.append(
                f"Status mismatch (file: {'Used' if file_used else 'Not used'}, "
                f"summary: {'Used' if sum_used else 'Not used'})"
            )
            level = max(level, 1)
        if (file_used or sum_used) and level < 3:
            remarks.append("Already used")
            level = max(level, 2)
        if len(files) > 1:
            remarks.append(f"{len(files)} files matched")
            level = max(level, 1)
        if rows and len(rows) > 1:
            remarks.append(f"Listed {len(rows)} times in pass summary")
            level = max(level, 1)

        results.append(
            {
                "pass_no": w, "files": files, "rows": rows,
                "level": level, "remarks": "; ".join(remarks) if remarks else "OK",
            }
        )
    return results


def style_sheet(ws, widths):
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def result_to_row(r):
    """One required pass (this run) -> a report row."""
    rows = r["rows"]
    if rows is None:
        in_sum, date, qty, status = "Not checked", None, None, "Not checked"
    elif not rows:
        in_sum, date, qty, status = "No", None, None, ""
    else:
        first = rows[0]
        in_sum, date, qty, status = "Yes", first["date"], first["qty"], first["status"]
    folders = []
    for p in r["files"]:
        if str(p.parent) not in folders:
            folders.append(str(p.parent))
    return {
        "Pass No": r["pass_no"], "In Pass Summary": in_sum, "Date": date, "Quantity (Cum)": qty,
        "Summary Status": status, "PDF Found": "Yes" if r["files"] else "No",
        "File Name(s)": ", ".join(dict.fromkeys(display_file_name(p.name) for p in r["files"])),
        "Found In": "; ".join(folders), "Remarks": r["remarks"],
    }


def load_existing(path: Path):
    """Read an earlier report: (rows of 'All Required', rows of 'Run History')."""
    rows, runs = [], []
    if not path.is_file():
        return rows, runs
    wb = load_workbook(path)
    if "All Required" in wb.sheetnames:
        ws = wb["All Required"]
        header = [c.value for c in ws[1]]
        for vals in ws.iter_rows(min_row=2, values_only=True):
            d = dict(zip(header, vals))
            if d.get("Pass No"):
                rows.append(d)
    if "Run History" in wb.sheetnames:
        runs = [list(v) for v in wb["Run History"].iter_rows(min_row=2, values_only=True) if any(v)]
    return rows, runs


def split_names(text):
    return [t.strip() for t in str(text or "").split(", ") if t.strip()]


def merge_rows(old_rows, new_rows):
    """Update the earlier report with this run's results (never loses a 'found')."""
    index, order = {}, []
    for r in old_rows:
        k = norm(r["Pass No"])
        if k not in index:
            order.append(k)
        index[k] = r
    for nr in new_rows:
        k = norm(nr["Pass No"])
        old = index.get(k)
        if old is None:
            index[k] = nr
            order.append(k)
            continue
        old_found, new_found = old.get("PDF Found") == "Yes", nr["PDF Found"] == "Yes"
        if old_found and not new_found:
            continue                                  # keep what an earlier run found
        if new_found and not old_found:
            index[k] = nr                             # found now - replaces the 'not found' row
        elif new_found and old_found:
            names = split_names(old.get("File Name(s)"))
            existing_keys = {canonical_file_name(n) for n in names}
            added = []
            for n in split_names(nr["File Name(s)"]):
                key_name = canonical_file_name(n)
                if key_name not in existing_keys:
                    added.append(display_file_name(n))
                    existing_keys.add(key_name)
            if added:
                old["File Name(s)"] = ", ".join(names + added)
                folders = [f.strip() for f in str(old.get("Found In") or "").split(";") if f.strip()]
                for f in nr["Found In"].split("; "):
                    if f and f not in folders:
                        folders.append(f)
                old["Found In"] = "; ".join(folders)
                note = "found again in another folder (extra file(s) added)"
                rem = str(old.get("Remarks") or "")
                if note not in rem:
                    old["Remarks"] = note if rem in ("", "OK") else f"{rem}; {note}"
        else:
            index[k] = nr                             # still not found - refresh the details
    return [index[k] for k in order]


def row_level(row):
    rem = str(row.get("Remarks") or "")
    if row.get("PDF Found") != "Yes":
        return 3
    if "Already used" in rem:
        return 2
    return 0 if rem in ("", "OK") else 1


def save_report(path: Path, merged, runs):
    wb = Workbook()
    fills = {1: YELLOW, 2: ORANGE, 3: RED}
    cols = ["Pass No", "In Pass Summary", "Date", "Quantity (Cum)", "Summary Status",
            "PDF Found", "File Name(s)", "Found In", "Remarks"]

    ws = wb.active
    ws.title = "All Required"
    ws.append(["Sl No"] + cols)
    for i, row in enumerate(merged, start=1):
        ws.append([i] + [row.get(c) for c in cols])
        ws.cell(ws.max_row, 4).number_format = "DD-MMM-YYYY"
        ws.cell(ws.max_row, 5).number_format = "0.00"
        lvl = row_level(row)
        if lvl in fills:
            for c in ws[ws.max_row]:
                c.fill = fills[lvl]
    style_sheet(ws, [7, 24, 16, 14, 16, 16, 11, 46, 46, 46])

    ws2 = wb.create_sheet("Not Found")
    ws2.append(["Sl No", "Pass No", "In Pass Summary", "Date", "Quantity (Cum)", "Summary Status", "Remarks"])
    missing = [r for r in merged if r.get("PDF Found") != "Yes"]
    for i, row in enumerate(missing, start=1):
        ws2.append([i, row["Pass No"], row.get("In Pass Summary"), row.get("Date"), row.get("Quantity (Cum)"),
                    row.get("Summary Status"), row.get("Remarks")])
        ws2.cell(ws2.max_row, 4).number_format = "DD-MMM-YYYY"
        ws2.cell(ws2.max_row, 5).number_format = "0.00"
        for c in ws2[ws2.max_row]:
            c.fill = RED
    if not missing:
        ws2.append(["", "All required passes have been found."])
    style_sheet(ws2, [7, 24, 16, 14, 16, 16, 48])

    ws3 = wb.create_sheet("Run History")
    ws3.append(["Run Date & Time", "Source Folder", "Pass List", "Required", "Found", "Not Found"])
    for run in runs:
        ws3.append(run)
        ws3.cell(ws3.max_row, 1).number_format = "DD-MMM-YYYY HH:MM"
    style_sheet(ws3, [20, 60, 30, 10, 10, 10])

    wb.save(path)


def update_report(path: Path, results, run_rows):
    """Create the report, or update the one already in the destination folder.
    Returns (path_used, total, total_found, existed, fell_back)."""
    old_rows, runs = load_existing(path)
    existed = path.is_file()
    merged = merge_rows(old_rows, [result_to_row(r) for r in results])
    runs.extend(run_rows)
    used, fell_back = path, False
    try:
        save_report(path, merged, runs)
    except PermissionError:                           # report is open in Excel
        used = path.with_name(f"retrieval_report_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        save_report(used, merged, runs)
        fell_back = True
    total_found = sum(1 for r in merged if r.get("PDF Found") == "Yes")
    return used, len(merged), total_found, existed, fell_back


# ---------- searching ----------
def find_pdfs(folder: Path, recursive: bool, skip: Path):
    """PDFs in a folder (and its sub-folders if recursive), never looking inside `skip`."""
    skip_res = skip.resolve() if skip else None
    if not recursive:
        return sorted((p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"),
                      key=lambda p: str(p).lower())
    pdfs = []
    for dirpath, dirnames, filenames in os.walk(folder):
        if skip_res:
            dirnames[:] = [d for d in dirnames if (Path(dirpath) / d).resolve() != skip_res]
        pdfs += [Path(dirpath) / f for f in filenames if f.lower().endswith(".pdf")]
    return sorted(pdfs, key=lambda p: str(p).lower())


def search_folder(folder: Path, passes, recursive: bool, skip: Path, summary=None):
    """Use summary filenames first; only scan PDF contents for unresolved passes."""
    pdfs = find_pdfs(folder, recursive, skip)
    print(f"\\nSearching {len(pdfs)} PDF(s) in {folder}")
    found = {w: [] for w in passes}

    # Summary filename lookup: mapped PDFs are not opened or text-extracted.
    if summary is not None:
        index_summary_files(summary, folder, recursive)
        for w in passes:
            rows = summary.get(norm(w), [])
            found[w] = [r["_path"] for r in rows if r.get("_path") and r["_path"].is_file()]
            preferred = next((r["status"] for r in rows if r.get("status")), "")
            found[w] = dedupe_files(found[w], preferred)

    # Fallback to filename matching, then PDF text only for unresolved passes.
    missing = [w for w in passes if not found[w]]
    for w in missing:
        rows = (summary or {}).get(norm(w), [])
        preferred = next((r["status"] for r in rows if r.get("status")), "")
        found[w] = dedupe_files(find_by_name(w, pdfs), preferred)

    missing = [w for w in passes if not found[w]]
    if missing and pdfs:
        print(f"{len(missing)} pass(es) still unresolved - checking PDF contents...")
        content_map = {}
        for n, p in enumerate(pdfs, start=1):
            if n % 100 == 0:
                print(f"  read {n} of {len(pdfs)} ...")
            pn = read_pass_from_pdf(p)
            if pn:
                content_map.setdefault(pn, []).append(p)
        for w in missing:
            rows = (summary or {}).get(norm(w), [])
            preferred = next((r["status"] for r in rows if r.get("status")), "")
            found[w] = dedupe_files(content_map.get(norm(w), []), preferred)
    return found


def ask_next_folder(searched, dest: Path, n_missing: int):
    """Offer another folder for the passes still missing. Returns a folder or None."""
    while True:
        if input(f"\n{n_missing} pass(es) still not found. Search another folder for them? (y/n): ").strip().lower() != "y":
            return None
        f = choose_folder("Select another folder to search for the missing passes")
        if not f or not f.is_dir():
            print("No folder selected.")
            return None
        if f.resolve() == dest.resolve():
            print("That is the destination folder - choose a different one.")
            continue
        if f.resolve() in searched:
            print("That folder was already searched.")
            continue
        return f


# ---------- main ----------
def main():
    print("=== Transit Pass Retriever ===\n")

    source = choose_folder("Select the folder containing the PDFs")
    if not source or not source.is_dir():
        print("No valid source folder selected.")
        return

    dest = choose_folder(f"Select the destination folder (Cancel = '{DEFAULT_DEST}' inside source)")
    if not dest or dest.resolve() == source.resolve():
        dest = source / DEFAULT_DEST

    recursive = input("\nAlso search inside sub-folders? (y/n): ").strip().lower() == "y"

    # required passes: from the destination's report (still not found) or from a list
    wanted, list_label = None, ""
    try:
        prev_rows, _ = load_existing(dest / REPORT_NAME)
    except Exception:
        prev_rows = []
    still = [r["Pass No"] for r in prev_rows if r.get("PDF Found") != "Yes"]
    if still:
        ans = input(f"\nThe report in the destination folder has {len(still)} pass(es) still not found.\n"
                    "Use those as the required list? (y/n): ")
        if ans.strip().lower() == "y":
            wanted, list_label = clean_list(still), "not-found list from report"
    if wanted is None:
        list_file = choose_file("Select the file with required pass numbers (Cancel = type them)")
        if list_file and list_file.is_file():
            wanted, list_label = read_pass_list(list_file), list_file.name
        else:
            wanted, list_label = type_pass_list(), "typed in"
    if not wanted:
        print("No pass numbers provided.")
        return
    print(f"\n{len(wanted)} pass number(s) required.")

    # pass summary: use the one in the first PDF folder, else ask
    summary_path = source / SUMMARY_NAME
    if not summary_path.is_file():
        summary_path = choose_file(
            f"Select {SUMMARY_NAME} to check against (Cancel = skip the check)", "*.xlsx *.xlsm"
        )
    summary = None
    if summary_path and summary_path.is_file():
        try:
            summary = read_summary(summary_path)
            print(f"Checking against pass summary: {summary_path.name} ({len(summary)} pass nos)")
        except Exception as e:
            print(f"Could not read the pass summary ({e}) - check skipped.")
    else:
        print("No pass summary selected - check against summary skipped.")

    # search the first folder, then (if asked) more folders for whatever is still missing
    matches = {w: [] for w in wanted}
    run_rows, searched = [], {source.resolve()}
    folder, to_search = source, list(wanted)
    while True:
        res = search_folder(folder, to_search, recursive, dest, summary)
        got = 0
        for w in to_search:
            if res[w]:
                matches[w] = res[w]
                got += 1
        print(f"{got} of {len(to_search)} found in this folder.")
        run_rows.append([datetime.now(), str(folder), list_label, len(to_search), got, len(to_search) - got])

        to_search = [w for w in wanted if not matches[w]]
        if not to_search:
            break
        folder = ask_next_folder(searched, dest, len(to_search))
        if folder is None:
            break
        searched.add(folder.resolve())
        extra_summary = folder / SUMMARY_NAME
        if extra_summary.is_file():
            try:
                extra_data = read_summary(extra_summary)
                if summary is None:
                    summary = {}
                for k, rows in extra_data.items():
                    summary.setdefault(k, rows)
                print(f"Also checking against {extra_summary}")
            except Exception as e:
                print(f"Could not read {extra_summary}: {e}")

    results = analyse(wanted, matches, summary)
    found = [r for r in results if r["files"]]
    not_found = [r for r in results if not r["files"]]
    attention = [r for r in results if r["remarks"] != "OK"]

    print("\nFOUND" if found else "\nFOUND: none")
    for r in found:
        for p in r["files"]:
            flag = "   [marked Used]" if is_marked_used(p.name) else ""
            print(f"  {r['pass_no']}  ->  {p.name}{flag}")
    print("\nNOT FOUND" if not_found else "\nNOT FOUND: none")
    for r in not_found:
        print(f"  {r['pass_no']}   ({r['remarks']})")
    check = [r for r in attention if r["files"] or r["level"] != 3]
    if check:
        print("\nCHECK" + (" (against pass summary and files)" if summary is not None else ""))
        for r in check:
            print(f"  {r['pass_no']}: {r['remarks']}")

    print(f"\n{len(wanted)} required: {len(found)} found, {len(not_found)} not found, "
          f"{len(attention)} need attention.")

    files_to_copy = []
    for r in found:
        for p in r["files"]:
            if p not in files_to_copy:
                files_to_copy.append(p)

    copied = skipped = 0
    cancelled = False
    if files_to_copy:
        print(f"\n{len(files_to_copy)} file(s) to copy to: {dest}")
        if input("Proceed? (y/n): ").strip().lower() == "y":
            dest.mkdir(parents=True, exist_ok=True)
            for p in files_to_copy:
                target = dest / p.name
                if target.exists():
                    skipped += 1
                    continue
                shutil.copy2(p, target)
                copied += 1
        else:
            cancelled = True

    # one report per destination folder: created the first time, updated afterwards
    dest.mkdir(parents=True, exist_ok=True)
    try:
        report, total, total_found, existed, fell_back = update_report(dest / REPORT_NAME, results, run_rows)
        print(f"\nReport {'updated' if existed else 'created'}: {report}")
        print(f"The report now covers {total} pass(es): {total_found} found, {total - total_found} still not found.")
        if fell_back:
            print(f"NOTE: {REPORT_NAME} is open in Excel, so this run was saved as a separate file instead.")
            print("      Close the old one and delete/merge as needed.")
    except Exception as e:
        print(f"\nCould not save the report: {e}")

    if cancelled:
        print("Cancelled. Nothing was copied.")
    elif files_to_copy:
        print(f"Copied {copied} file(s) to {dest}")
        if skipped:
            print(f"{skipped} file(s) were already in the destination and were skipped.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nUnexpected error: {e}")
    input("\nPress Enter to exit...")

