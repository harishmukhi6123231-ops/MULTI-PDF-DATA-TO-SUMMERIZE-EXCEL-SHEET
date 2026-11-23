"""
Pass Renamer + Pass Summary   (fast PyMuPDF version)
----------------------------------------------------
Put this file in the folder that holds your Transit Pass PDFs and run it.
It works on the PDFs in ITS OWN folder (sub-folders are not touched).

pass_summary.xlsx is UPDATED, never rebuilt:
  * PDFs whose file name is already in the summary are skipped completely
    (not read, not renamed, their rows are not touched).
  * Only NEW PDFs are read. Each new pass is added as a new row at the bottom.
  * Rows you typed yourself (e.g. for PDFs the program cannot read) are NEVER
    changed - existing cells are never overwritten. The only thing ever written
    into an existing row is a BLANK 'File Name' cell, to link an old row (made
    without file names) to its PDF. A copy of the workbook as it was before the
    run is kept as pass_summary_backup.xlsx.

For the new PDFs you choose:
  1 = RENAME them and record them          2 = only RECORD them (no renaming)

Renaming follows ONE rule:   <Quantity>-<Pass No>.pdf    e.g. 14-A2603070037-2-539.pdf
  - "/" in the pass no becomes "-"
  - a U / Used at the end of the ORIGINAL name is kept ("abc U.pdf", "abc_Used.pdf")
  - a file is renamed ONLY if both the Pass No and the Quantity were read properly.
    Anything else keeps its name exactly as it is.
  - a name that already exists gets _1, _2 ... (nothing is overwritten)

Sheet "Transit Passes":  Date | Quantity | Unit | Pass No | Status | File Name | Remarks
   - File Name is the CURRENT name of the PDF (the retriever uses it to find the file)
   - a transit pass whose pass no or quantity could not be read is added with those
     cells BLANK and highlighted: fill them in by hand - they will not be touched again
Sheet "Not Read":        PDFs that are not passes / could not be read (names left as they are).
                         If one of these IS a pass, add a row for it in "Transit Passes",
                         typing its file name in the File Name column.

Install:  pip install pymupdf openpyxl
Run:      python pdf_extractor.py      (or double-click app.py (choose option 2))
"""

import re
import shutil
import sys
from collections import defaultdict
from copy import copy
from datetime import datetime
from pathlib import Path

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

EXCEL_NAME = "pass_summary.xlsx"
BACKUP_NAME = "pass_summary_backup.xlsx"
SHEET_MAIN = "Transit Passes"
SHEET_NOT_READ = "Not Read"
STD_HEAD = ["Date", "Quantity", "Unit", "Pass No", "Status", "File Name", "Remarks"]

# A real pass no looks like A2606040089/1/1  (letter + digits + /n/n). Anything else is NOT accepted,
# so a file can never be renamed to a wrong word such as ORIGINAL.
PASS_VALID = re.compile(r"[A-Z]\d{8,}(?:/\d+){1,3}")
PASS_TOKEN_RE = re.compile(r"\b[A-Z]\d{8,}(?:/\d+){1,3}\b")
PASS_LABEL_RE = re.compile(r"Pass\s*No\.?\s*[:\-]?\s*([A-Z0-9/]+)", re.I)
# fallback only: the number must be followed by its unit (14.00 Cum), so the next item's number is never taken
QTY_RE = re.compile(r"Quantity\s*Permitted[^:\n]*:\s*([\d,]*\.?\d+)[ \t]+(Cum|Cu\.?\s?m|Tonnes?|Tons?|MT)\b", re.I)
DATE_RE = re.compile(r"Date\s*&\s*Time\s*:?\s*(\d{1,2}-[A-Za-z]{3}-\d{4})", re.I)
DATE_TOKEN = re.compile(r"\d{1,2}-[A-Za-z]{3}-\d{4}")
NUM_TOKEN = re.compile(r"[\d,]*\.?\d+")
USED_RE = re.compile(r"[\s_\-.(\[](?:U|USED)[)\]]?\s*$", re.I)     # U / Used at the end of the name

HEADER = PatternFill("solid", fgColor="1F4E78")
USED_FILL = PatternFill("solid", fgColor="F8CBAD")
CHECK_FILL = PatternFill("solid", fgColor="FFEB9C")


def script_folder() -> Path:
    """The folder this program lives in (works for .py and for a built .exe)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent / "input_pdfs"


# ======================= reading a pass =======================
def pick_clip(page):
    """ORIGINAL | DUPLICATE side by side: return the box of the ORIGINAL copy, else None."""
    orig, dup = page.search_for("ORIGINAL COPY"), page.search_for("DUPLICATE COPY")
    if not (orig and dup):
        return None
    r = page.rect
    if abs(orig[0].x0 - dup[0].x0) >= abs(orig[0].y0 - dup[0].y0):
        mid = (r.x0 + r.x1) / 2
        return fitz.Rect(r.x0, r.y0, mid, r.y1) if orig[0].x0 < mid else fitz.Rect(mid, r.y0, r.x1, r.y1)
    mid = (r.y0 + r.y1) / 2
    return fitz.Rect(r.x0, r.y0, r.x1, mid) if orig[0].y0 < mid else fitz.Rect(r.x0, mid, r.x1, r.y1)


def row_words(words, anchor):
    """Words on the same line as `anchor`, to its right, left to right."""
    return sorted((w for w in words if abs(w["top"] - anchor["top"]) <= 3 and w["x0"] > anchor["x0"]),
                  key=lambda w: w["x0"])


def find_pass(words, text):
    """Return (pass_no or None, problem text). The value must look like a real pass no."""
    for a in words:                                              # 1) on the same line as the word "Pass"
        if a["text"].lower() == "pass":
            for w in row_words(words, a):
                if PASS_VALID.fullmatch(w["text"].upper()):
                    return w["text"].upper(), ""
    for m in PASS_LABEL_RE.finditer(text):                       # 2) written right after "Pass No."
        if PASS_VALID.fullmatch(m.group(1).upper()):
            return m.group(1).upper(), ""
    tokens = {t.upper() for t in PASS_TOKEN_RE.findall(text)}    # 3) the only pass-looking token on the page
    if len(tokens) == 1:
        return tokens.pop(), ""
    if len(tokens) > 1:
        return None, "more than one pass-like number on the page"
    return None, "pass no not found"


def find_qty(words, text):
    for a in words:                                              # on the same line as "Quantity"
        if a["text"].lower() == "quantity":
            rw = row_words(words, a)
            for i, w in enumerate(rw):
                if NUM_TOKEN.fullmatch(w["text"]):
                    unit = rw[i + 1]["text"] if i + 1 < len(rw) and re.fullmatch(r"[A-Za-z]{2,10}", rw[i + 1]["text"]) else ""
                    return float(w["text"].replace(",", "")), unit.title()
    m = QTY_RE.search(text)
    if m:
        return float(m.group(1).replace(",", "")), m.group(2).title()
    return None, ""


def find_date(words, text):
    for a in words:                                              # on the same line as "Time" (Date & Time)
        if a["text"] == "Time":
            for w in row_words(words, a):
                if DATE_TOKEN.fullmatch(w["text"]):
                    return parse_date(w["text"])
    m = DATE_RE.search(text)
    return parse_date(m.group(1)) if m else None


def parse_date(s):
    try:
        return datetime.strptime(s, "%d-%b-%Y")
    except ValueError:
        return None


def read_pdf(path: Path):
    """Return a dict: kind = 'PASS' | 'NOT PASS' | 'NO TEXT', plus fields / reason."""
    with fitz.open(path) as doc:
        if doc.page_count == 0:
            return {"kind": "NO TEXT", "reason": "no pages"}
        page = doc[0]
        clip = pick_clip(page)
        text = page.get_text("text", clip=clip)
        words = [{"x0": w[0], "x1": w[2], "top": w[1], "text": w[4]} for w in page.get_text("words", clip=clip)]
    if not text.strip():
        return {"kind": "NO TEXT", "reason": "no text (scanned image - needs OCR)"}
    if "TRANSIT PASS" not in re.sub(r"\s+", " ", text).upper():
        return {"kind": "NOT PASS", "reason": "not a transit pass"}
    pass_no, problem = find_pass(words, text)
    qty, unit = find_qty(words, text)
    return {"kind": "PASS", "pass_no": pass_no, "problem": problem, "qty": qty, "unit": unit,
            "date": find_date(words, text)}


# ======================= small helpers =======================
def clean_name(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', "-", name)
    return re.sub(r"\s+", " ", name).strip(" .")


def used_suffix(stem: str) -> str:
    m = USED_RE.search(stem)
    return m.group(0).rstrip() if m else ""


def norm_pass(p) -> str:
    return re.sub(r"\s+", "", str(p)).upper().replace("/", "-")


def blank(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


# ======================= the summary workbook =======================
class SummaryBook:
    """pass_summary.xlsx: opened if it exists (cells are never overwritten), otherwise created."""

    def __init__(self, path: Path):
        self.path = path
        self.existed = path.is_file()
        if self.existed:
            self.wb = load_workbook(path)
            self.ws = self.wb[SHEET_MAIN] if SHEET_MAIN in self.wb.sheetnames else self.wb.worksheets[0]
        else:
            self.wb = Workbook()
            self.ws = self.wb.active
            self.ws.title = SHEET_MAIN
            self.ws.append(STD_HEAD)
            style_header(self.ws, [14, 11, 8, 24, 11, 40, 60])
            self.ws.freeze_panes = "A2"
        self._map_columns()
        self._scan_rows()

    # ---- columns (found by their heading, so older layouts still work) ----
    def _map_columns(self):
        heads = {}
        for c in range(1, self.ws.max_column + 1):
            h = str(self.ws.cell(1, c).value or "").strip().lower()
            if h:
                heads[c] = h

        def first(test):
            return next((c for c, h in heads.items() if test(h)), None)

        self.c_date = first(lambda h: "date" in h)
        self.c_qty = first(lambda h: "quant" in h)
        self.c_unit = first(lambda h: h == "unit")
        self.c_pass = first(lambda h: "pass" in h)
        self.c_status = first(lambda h: "status" in h)
        self.c_file = first(lambda h: h in ("file name", "filename"))
        self.c_new = first(lambda h: h == "new file name")
        self.c_orig = first(lambda h: h == "original file name")
        self.c_remarks = first(lambda h: h in ("remarks", "remark"))
        if self.c_pass is None:
            raise ValueError("no 'Pass No' heading found in row 1 of the summary")
        if not (self.c_file or self.c_new or self.c_orig):
            self.c_file = self._add_column("File Name")
        if self.c_remarks is None:
            self.c_remarks = self._add_column("Remarks")

    def _add_column(self, title):
        c = self.ws.max_column + 1
        cell = self.ws.cell(1, c, title)
        left = self.ws.cell(1, c - 1)
        cell._style = copy(left._style)
        self.ws.column_dimensions[get_column_letter(c)].width = 40 if "File" in title else 55
        return c

    def _scan_rows(self):
        ws = self.ws
        name_cols = [c for c in (self.c_file, self.c_new, self.c_orig) if c]
        self.row_names, self.rows_by_pass, last = {}, defaultdict(list), 1
        for r in range(2, ws.max_row + 1):
            pn = ws.cell(r, self.c_pass).value
            names = {str(ws.cell(r, c).value).strip() for c in name_cols if not blank(ws.cell(r, c).value)}
            others = any(not blank(ws.cell(r, c).value) for c in (self.c_date, self.c_qty) if c)
            if not blank(pn) or names or others:
                last = r
            if names:
                self.row_names[r] = names
            if not blank(pn):
                self.rows_by_pass[norm_pass(pn)].append(r)
        self.last_row = last
        self.recorded = {n.casefold() for names in self.row_names.values() for n in names}
        self.not_read_names = set()
        if SHEET_NOT_READ in self.wb.sheetnames:
            nr = self.wb[SHEET_NOT_READ]
            self.not_read_names = {str(nr.cell(r, 1).value or "").strip().casefold() for r in range(2, nr.max_row + 1)}
            self.not_read_names.discard("")

    # ---- writing (append only) ----
    def add_row(self, date, qty, unit, pass_no, status, orig_name, new_name, remarks, flag, used):
        ws, r = self.ws, self.last_row + 1
        self.last_row = r

        def put(col, val, fmt=None):
            if col and not blank(val):
                cell = ws.cell(r, col, val)
                if fmt:
                    cell.number_format = fmt

        put(self.c_date, date, "DD-MMM-YYYY")
        put(self.c_qty, qty, "0.00")
        put(self.c_unit, unit)
        put(self.c_pass, pass_no)
        put(self.c_status, status)
        if self.c_file:
            put(self.c_file, new_name or orig_name)
        else:
            put(self.c_orig, orig_name)
            put(self.c_new, new_name)
        put(self.c_remarks, remarks)
        fill = CHECK_FILL if flag else USED_FILL if used else None
        if fill:
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = fill
        return r

    def link_row(self, row, orig_name, new_name):
        """Write the file name into a row that has NONE (blank cells only - nothing is overwritten)."""
        ws = self.ws
        if self.c_file:
            if blank(ws.cell(row, self.c_file).value):
                ws.cell(row, self.c_file, new_name or orig_name)
        else:
            if self.c_orig and blank(ws.cell(row, self.c_orig).value):
                ws.cell(row, self.c_orig, orig_name)
            if self.c_new and new_name and blank(ws.cell(row, self.c_new).value):
                ws.cell(row, self.c_new, new_name)

    def add_not_read(self, name, reason):
        if SHEET_NOT_READ in self.wb.sheetnames:
            ws = self.wb[SHEET_NOT_READ]
        else:
            ws = self.wb.create_sheet(SHEET_NOT_READ)
            ws.append(["File Name", "Reason", "Action"])
            style_header(ws, [44, 60, 22])
        known = {str(ws.cell(r, 1).value or "").strip().casefold() for r in range(2, ws.max_row + 1)}
        if name.casefold() in known:
            return False
        ws.append([name, reason, "Name left as it is"])
        return True

    def save(self):
        if self.existed:
            shutil.copy2(self.path, self.path.with_name(BACKUP_NAME))        # the version before this run
        elif self.ws.max_row > 1:
            self.ws.auto_filter.ref = self.ws.dimensions
        self.wb.save(self.path)


def style_header(ws, widths):
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ======================= reading the new PDFs =======================
def read_new(new_pdfs):
    items = []
    for n, pdf in enumerate(new_pdfs, start=1):
        if n % 200 == 0:
            print(f"  read {n} of {len(new_pdfs)} ...")
        it = {"path": pdf, "sfx": used_suffix(pdf.stem), "remarks": [], "action": "new", "row": None,
              "pass_no": None, "qty": None, "unit": "", "date": None, "new_name": ""}
        it["used"] = bool(it["sfx"])
        try:
            r = read_pdf(pdf)
        except Exception as e:
            it.update(kind="NOT READ", reason=f"could not be opened/read: {e}")
            items.append(it)
            continue
        if r["kind"] != "PASS":
            it.update(kind="NOT READ", reason=r["reason"])
        elif not r["pass_no"]:
            it.update(kind="NO PASS", date=r["date"], qty=r["qty"] if r["qty"] and r["qty"] > 0 else None,
                      unit=r["unit"], reason=f"transit pass, but {r['problem']}")
        else:
            qty = r["qty"] if r["qty"] and r["qty"] > 0 else None
            it.update(kind="READY" if qty else "NO QTY", pass_no=r["pass_no"], qty=qty, unit=r["unit"], date=r["date"])
        if it["kind"] in ("READY", "NO QTY") and it["date"] is None:
            it["remarks"].append("date not read")
        items.append(it)
    return items


def link_to_summary(items, book: SummaryBook, present: set):
    """Decide for every new pass: add a new row, link an existing row, or leave it for the user."""
    claimed, seen_new = set(), defaultdict(list)
    for it in sorted(items, key=lambda i: i["path"].name.lower()):
        if it["kind"] not in ("READY", "NO QTY"):
            continue
        key = norm_pass(it["pass_no"])
        rows = book.rows_by_pass.get(key, [])
        unlinked = [r for r in rows if r not in book.row_names and r not in claimed]
        present_rows = [r for r in rows if r in book.row_names and any(n.casefold() in present for n in book.row_names[r])]
        missing_rows = [r for r in rows if r in book.row_names and r not in present_rows]
        if unlinked:
            it["action"], it["row"] = "link", unlinked[0]
            claimed.add(unlinked[0])
        elif present_rows:
            it["remarks"].append(f"same pass no also in row {present_rows[0]} ({sorted(book.row_names[present_rows[0]])[0]})")
        elif missing_rows:
            it["action"], it["row"] = "renamed?", missing_rows[0]
        if seen_new[key] and it["action"] == "new":
            it["remarks"].append(f"same pass no also in new file {seen_new[key][0]}")
        if it["action"] != "renamed?":
            seen_new[key].append(it["path"].name)


def plan_renames(items, pdfs):
    taken = {p.name.lower() for p in pdfs}
    plan = []
    for it in sorted(items, key=lambda i: i["path"].name.lower()):
        if it["kind"] != "READY" or it["action"] not in ("new", "link"):
            continue                                            # condition not met / left for the user
        base = clean_name(f"{it['qty']:g}-{it['pass_no']}") + it["sfx"]
        if re.fullmatch(re.escape(base) + r"(_\d+)?", it["path"].stem, re.I):
            continue                                            # already named correctly
        new_name, k = f"{base}.pdf", 1
        while new_name.lower() in taken:
            new_name, k = f"{base}_{k}.pdf", k + 1
        taken.add(new_name.lower())
        it["new_name"] = new_name
        plan.append(it)
    return plan


# ======================= main =======================
def build_individual_pass_summary():
    folder = script_folder()
    path = folder.parent / "output" / EXCEL_NAME
    print("=== Pass Renamer + Summary ===\n")
    print(f"Folder: {folder}\n")

    try:
        book = SummaryBook(path)
    except Exception as e:
        print(f"Could not use {EXCEL_NAME}: {e}\nNothing was changed.")
        return
    pdfs = sorted((p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"),
                  key=lambda p: p.name.lower())
    present = {p.name.casefold() for p in pdfs}
    skipped_nr = [p for p in pdfs if p.name.casefold() in book.not_read_names and p.name.casefold() not in book.recorded]
    new_pdfs = [p for p in pdfs if p.name.casefold() not in book.recorded and p.name.casefold() not in book.not_read_names]
    stale = sum(1 for names in book.row_names.values() if not any(n.casefold() in present for n in names))

    print(f"{len(pdfs)} PDF(s) in the folder: {len(pdfs) - len(new_pdfs) - len(skipped_nr)} already in the summary (left alone), "
          f"{len(skipped_nr)} listed in 'Not Read' (skipped), {len(new_pdfs)} new.")
    if skipped_nr:
        print("  (delete a file's row from the 'Not Read' sheet if you want it checked again)")
    if stale:
        print(f"{stale} summary row(s) refer to files that are not in this folder (left as they are).")
    if not new_pdfs:
        print("\nNothing new - the summary is up to date.")
        return

    print(f"\nReading the {len(new_pdfs)} new PDF(s)...\n")
    items = read_new(new_pdfs)
    link_to_summary(items, book, present)

    ready = [i for i in items if i["kind"] in ("READY", "NO QTY") and i["action"] in ("new", "link")]
    fill_in = [i for i in items if i["kind"] in ("NO QTY", "NO PASS")]
    not_read = [i for i in items if i["kind"] == "NOT READ"]
    renamed_q = [i for i in items if i["action"] == "renamed?"]
    print(f"New PDFs: {len(ready)} pass(es) read"
          f", {len([i for i in items if i['kind'] == 'NO PASS'])} transit pass(es) with an unreadable pass no"
          f", {len(not_read)} not read.\n")

    print("What do you want to do with the new PDFs?")
    print("  1 = RENAME them (<Quantity>-<Pass No>) and record them in the summary")
    print("  2 = only RECORD them in the summary (no file is renamed)")
    choice = input("Type 1 or 2 (anything else cancels): ").strip()
    if choice not in ("1", "2"):
        print("Cancelled. Nothing was changed.")
        return
    rename = choice == "1"
    plan = plan_renames(items, pdfs) if rename else []

    if rename:
        if plan:
            print("\nPREVIEW - will be renamed")
            for it in plan:
                print(f"  {it['path'].name}  ->  {it['new_name']}")
        else:
            print("\nNo new files need renaming.")
    else:
        print("\nRecord only: no file will be renamed.")
    links = [i for i in items if i["action"] == "link"]
    if links:
        print("\nLINKED to an existing row (only its blank File Name cell is filled)")
        for it in links:
            print(f"  {it['path'].name}  ->  row {it['row']}  [{it['pass_no']}]")
    if fill_in:
        print("\nTO FILL IN BY HAND (added with blank cells, highlighted; name left as it is)")
        for it in fill_in:
            print(f"  {it['path'].name}: {it['reason'] if it['kind'] == 'NO PASS' else 'quantity not read'}")
    if renamed_q:
        print("\nNOT ADDED - pass no already recorded under a file name that is no longer in this folder")
        for it in renamed_q:
            print(f"  {it['path'].name}  [{it['pass_no']}]  = row {it['row']} "
                  f"({', '.join(sorted(book.row_names[it['row']]))}) - update that row yourself if it is the same file")
    if not_read:
        print("\nNOT READ - names left as they are, listed in the 'Not Read' sheet")
        for it in not_read[:25]:
            print(f"  {it['path'].name}: {it['reason']}")
        if len(not_read) > 25:
            print(f"  ... and {len(not_read) - 25} more")

    if input(f"\nProceed? (y/n): ").strip().lower() != "y":
        print("Cancelled. Nothing was changed.")
        return
    if path.exists():                                           # fail BEFORE renaming if Excel has it open
        try:
            open(path, "ab").close()
        except PermissionError:
            print(f"\n{EXCEL_NAME} is open in Excel. Close it and run again. Nothing was changed.")
            return

    failed = 0
    for it in plan:
        try:
            it["path"].rename(it["path"].with_name(it["new_name"]))
        except Exception as e:
            failed += 1
            it["remarks"].append(f"RENAME FAILED: {e}")
            it["new_name"] = ""
            print(f"  FAILED: {it['path'].name}: {e}")
    if rename:
        print(f"\nRenamed {len(plan) - failed} file(s)." + (f"  {failed} failed (open in another program?)" if failed else ""))

    added = linked = nr_added = 0
    for it in sorted(items, key=lambda i: (i["date"] or datetime.max, i["path"].name.lower())):
        status = "Used" if it["used"] else "Not used"
        if it["kind"] == "NOT READ":
            nr_added += book.add_not_read(it["path"].name, it["reason"])
        elif it["action"] == "renamed?":
            continue
        elif it["action"] == "link":
            book.link_row(it["row"], it["path"].name, it["new_name"])
            linked += 1
        elif it["kind"] == "NO PASS":
            book.add_row(it["date"], it["qty"], it["unit"], None, status, it["path"].name, None,
                         "pass no not read - please fill in Date, Quantity, Pass No and Status", True, False)
            added += 1
        else:
            rem = list(it["remarks"])
            if it["kind"] == "NO QTY":
                rem.insert(0, "quantity not read - name left as it is - please fill in the quantity")
            book.add_row(it["date"], it["qty"], it["unit"], it["pass_no"], status, it["path"].name,
                         it["new_name"] or None, "; ".join(rem), it["kind"] == "NO QTY", it["used"])
            added += 1
    try:
        book.save()
        print(f"\nSummary updated: {path}")
        print(f"  {added} row(s) added, {linked} existing row(s) linked to their file, {nr_added} file(s) added to 'Not Read'.")
        if book.existed:
            print(f"  Your existing rows were not changed. Previous version kept as {BACKUP_NAME}.")
    except PermissionError:
        print(f"\nCould not save {path}. Close it in Excel and run again.")



# Project-specific extraction into the supplied TP Summary workbook.
PROJECT_ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = PROJECT_ROOT / "input_data" / "TP Summary.xlsx"
RESULT_PATH = PROJECT_ROOT / "output" / "TP Summary.xlsx"
PASS_RE = re.compile(r"\b[A-Z]\d{8,}(?:/\d+){1,3}\b", re.I)


def extract_pass_record(pdf_path: Path):
    """Read the first pass on a PDF page and return the workbook fields."""
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    if not reader.pages:
        return None
    text = reader.pages[0].extract_text() or ""
    flat = re.sub(r"\s+", " ", text)
    pass_numbers = list(dict.fromkeys(m.upper() for m in PASS_RE.findall(text)))
    if not pass_numbers:
        return None
    pass_no = pass_numbers[0]

    date_match = re.search(r"Date\s*&\s*Time\s*:?\s*(\d{1,2}-[A-Za-z]{3}-\d{4})", flat, re.I)
    pass_date = datetime.strptime(date_match.group(1), "%d-%b-%Y").date() if date_match else None

    qty_match = re.search(r"Quantity\s+Permitted\s*\(Cum/Tonnes\)\s*:?\s*([\d,]+(?:\.\d+)?)\s*(Cum|Tonnes?)?", flat, re.I)
    quantity = float(qty_match.group(1).replace(",", "")) if qty_match else None

    vehicle_match = re.search(r"\b[A-Z]{2}\d{1,2}[A-Z]{1,3}\d{3,4}\b", flat, re.I)
    vehicle = vehicle_match.group(0).upper() if vehicle_match else None

    dest_match = re.search(r"5\.\s*Destination\s*:?\s*(.*?)\s*6\.\s*Minor Mineral", flat, re.I)
    destination = dest_match.group(1).strip(" ,:") if dest_match else None
    destination = re.sub(r"\s+Route\s*:\s*\w+\s*$", "", destination or "", flags=re.I).strip(" ,:") or None

    party_match = re.search(r"4\.\s*Name of the Licensee/Lessee/Permit\s+Holder/Auction Holder/Auction Purchaser\s*:?\s*(.*?)\s*5\.", flat, re.I)
    first_party = party_match.group(1).strip(" ,:") if party_match else None
    return {
        "Pass No": pass_no,
        "Date": pass_date,
        "Quantity": quantity,
        "Carrier Vehicle no": vehicle,
        "Destination": destination,
        "1st Party": first_party,
        "2nd Party": None,
        "Return": None,
    }


def extract_into_template(pdf_paths=None, output_path=None):
    """Extract supplied PDFs (or project inputs) into the TP Summary workbook."""
    import shutil
    from openpyxl import load_workbook

    result_path = Path(output_path) if output_path else RESULT_PATH
    if not result_path.exists():
        if not TEMPLATE_PATH.is_file():
            raise FileNotFoundError(f"Workbook template not found: {TEMPLATE_PATH}")
        shutil.copy2(TEMPLATE_PATH, result_path)

    if pdf_paths is None:
        pdf_paths = []
        for folder_name in ("input_pdfs", "sample_pdfs"):
            folder = PROJECT_ROOT / folder_name
            if folder.exists():
                pdf_paths.extend(sorted(folder.glob("*.pdf"), key=lambda p: p.name.casefold()))
    else:
        pdf_paths = [Path(p) for p in pdf_paths]

    if not pdf_paths:
        print("No PDFs supplied or found in input_pdfs/sample_pdfs.")
        return {"items": [], "added": 0, "updated": 0, "ignored": 0, "failures": [], "result_path": result_path}

    wb = load_workbook(result_path)
    ws = wb.active
    headers = {str(cell.value or "").strip().casefold(): cell.column for cell in ws[1]}
    needed = ["pass no", "date", "quantity", "carrier vehicle no", "destination", "1st party", "2nd party", "return"]
    missing = [h for h in needed if h not in headers]
    if missing:
        wb.close()
        raise ValueError(f"Template is missing columns: {', '.join(missing)}")

    existing = {}
    for row in range(2, ws.max_row + 1):
        key = str(ws.cell(row, headers["pass no"]).value or "").strip().upper()
        if key:
            existing[key] = row

    added = updated = ignored = 0
    items, failures, batch_passes = [], [], set()
    for pdf_path in pdf_paths:
        try:
            record = extract_pass_record(pdf_path)
        except Exception as exc:
            failures.append({"source": pdf_path.name, "reason": str(exc)})
            continue
        if not record:
            ignored += 1
            failures.append({"source": pdf_path.name, "reason": "No transit-pass number or readable pass text found."})
            continue
        key = record["Pass No"]
        if key in batch_passes:
            ignored += 1
            continue
        batch_passes.add(key)
        row = existing.get(key)
        if row is None:
            row = ws.max_row + 1
            existing[key] = row
            added += 1
        else:
            updated += 1
        for name, value in record.items():
            cell = ws.cell(row, headers[name.casefold()], value)
            if name == "Date" and value is not None:
                cell.number_format = "dd-mmm-yyyy"
        items.append({"source": pdf_path.name, **record})

    if items:
        wb.save(result_path)
    wb.close()
    print(f"Saved: {result_path} ({added} added, {updated} refreshed, {ignored} skipped).")
    return {"items": items, "added": added, "updated": updated, "ignored": ignored,
            "failures": failures, "result_path": result_path}

def main():
    extract_into_template()


if __name__ == "__main__":
    main()




