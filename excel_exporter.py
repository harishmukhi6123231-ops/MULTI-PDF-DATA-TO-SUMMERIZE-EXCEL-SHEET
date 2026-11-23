"""
Invoice + Transit Pass Merger
-----------------------------
Put this file in the folder that holds your TAX INVOICE PDFs and run it.

For every tax invoice it reads  Invoice No, Invoice Date  and the  "Other References"
box. That box holds one or more TP (transit pass) numbers - separated by a space, slash,
comma, dot or nothing at all - so they are picked out by their PATTERN
(a letter + 10 digits + /n/n, e.g. A2603070037/2/120), never by the separator.

Each TP number is then matched to its TP PDF:
  1. the pass_summary.xlsx in the TP folder(s): the File Name column of the row with that pass no
  2. otherwise the PDF whose file name contains the pass no (e.g. 20-A2603070037-2-120.pdf)
  3. otherwise the PDFs are opened and the pass no read from inside (only for what is still missing)

When ALL the TPs of an invoice are found, one PDF is made:  invoice + TP1 + TP2 ...
saved as  <this folder>\\Merged\\RSCL-26-27-1031.pdf   (the originals are never changed or moved).
An invoice with a missing TP is NOT merged; it is listed in the summary and is merged
automatically on a later run, once the TP is available.

invoice_summary.xlsx (this folder):
  Invoice Date | Invoice No | Other Ref 1 | Other Ref 2 | Matched TP1 | Matched TP2 | Merged PDF | Invoice File | Remarks
  Invoices already merged are left exactly as they were (not read again).

Install:  pip install pymupdf openpyxl
Run:      python excel_exporter.py     (or double-click app.py (choose option 4))
"""

import os
import re
import sys
from collections import defaultdict
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

OUT_DIR = "Merged"
SUMMARY_NAME = "invoice_summary.xlsx"
TP_SUMMARY_NAME = "TP Summary.xlsx"
MAX_PAGES = 50

# A TP number: letter + 10 digits + /n/n   (A2603070037/2/120). Change here if your numbers differ.
TP_RE = re.compile(r"[A-Z]\d{10}/\d{1,3}/\d{1,4}")
INVOICE_RE = re.compile(r"\b[A-Z0-9]{2,10}/\d{2}-\d{2}/\d+\b")
DATE_TOKEN = re.compile(r"\d{1,2}-[A-Za-z]{3}-\d{2,4}")

HEADER = PatternFill("solid", fgColor="1F4E78")
RED = PatternFill("solid", fgColor="FFC7CE")
YELLOW = PatternFill("solid", fgColor="FFEB9C")
COLUMNS = ["Invoice Date", "Invoice No", "Other Ref 1", "Other Ref 2", "Matched TP1", "Matched TP2",
           "Merged PDF", "Invoice File", "Remarks"]


def script_folder() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent / "input_pdfs"


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
        typed = input(f"{title}\nEnter path (blank = this folder): ").strip().strip('"')
        return Path(typed) if typed else None


def norm_pass(p) -> str:
    return re.sub(r"\s+", "", str(p)).upper().replace("/", "-")


def blank(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


# ======================= reading an invoice =======================
def words_to_lines(ws):
    """Words -> lines (top to bottom, left to right)."""
    ws = sorted(ws, key=lambda w: (w["top"], w["x0"]))
    lines, cur, cur_top = [], [], None
    for w in ws:
        if cur_top is None or abs(w["top"] - cur_top) <= 2.5:
            if cur_top is None:
                cur_top = w["top"]
            cur.append(w)
        else:
            lines.append(cur)
            cur, cur_top = [w], w["top"]
    if cur:
        lines.append(cur)
    return [sorted(l, key=lambda w: w["x0"]) for l in lines]


def parse_inv_date(s):
    for fmt in ("%d-%b-%y", "%d-%b-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def tp_numbers(squashed: str):
    """TP numbers in a string, in order, without repeats."""
    out = []
    for m in TP_RE.findall(squashed.upper()):
        if m not in out:
            out.append(m)
    return out


def read_invoice(path: Path):
    """Return a dict with kind = 'INVOICE' | 'NOT INVOICE' | 'NO TEXT' (+ fields)."""
    with fitz.open(path) as doc:
        if doc.page_count == 0:
            return {"kind": "NO TEXT", "reason": "no pages"}
        page = doc[0]
        text = page.get_text("text")
        words = [{"x0": w[0], "x1": w[2], "top": w[1], "text": w[4]} for w in page.get_text("words")]
    if not text.strip():
        return {"kind": "NO TEXT", "reason": "no text (scanned image)"}
    if not re.search(r"TAX\s+INVOICE", text, re.I):
        return {"kind": "NOT INVOICE", "reason": "not a tax invoice"}

    ordered = [w for line in words_to_lines(words) for w in line]
    inv_w = next((w for w in ordered if INVOICE_RE.fullmatch(w["text"])), None)
    if inv_w is not None:
        invoice_no = inv_w["text"]
    else:
        m = INVOICE_RE.search(text)
        invoice_no = m.group(0) if m else None
    if not invoice_no:
        return {"kind": "NO TEXT", "reason": "tax invoice, but the invoice number was not found"}

    # invoice date: the value under the first "Dated" label (right-hand column), else beside the invoice no
    date = None
    dated = sorted((w for w in words if w["text"] == "Dated"), key=lambda w: w["top"])
    if dated:
        lab = dated[0]
        for w in ordered:
            if DATE_TOKEN.fullmatch(w["text"]) and abs(w["x0"] - lab["x0"]) <= 8 and lab["top"] < w["top"] <= lab["top"] + 25:
                date = parse_inv_date(w["text"])
                break
    if date is None and inv_w is not None:
        for w in ordered:
            if DATE_TOKEN.fullmatch(w["text"]) and abs(w["top"] - inv_w["top"]) <= 3 and w["x0"] > inv_w["x1"]:
                date = parse_inv_date(w["text"])
                break

    # "Other References" box: everything under its label, down to the next label ("Dated")
    remarks, shown = [], ""
    label = next((a for a, b in zip(ordered, ordered[1:])
                  if a["text"] == "Other" and b["text"] == "References" and abs(a["top"] - b["top"]) <= 3), None)
    if label is not None:
        below = [w for w in words if abs(w["x0"] - label["x0"]) <= 3 and w["text"] == "Dated" and w["top"] > label["top"] + 6]
        y_hi = min((w["top"] for w in below), default=label["top"] + 45)
        box = [w for w in words if w["x0"] >= label["x0"] - 3 and label["top"] + 4 < w["top"] < y_hi - 1]
        lines = words_to_lines(box)
        shown = " ".join(" ".join(w["text"] for w in l) for l in lines)
        refs = tp_numbers("".join(w["text"] for l in lines for w in l))     # squashed: separators don't matter
        left = re.sub(r"\s+", "", TP_RE.sub("", "".join(w["text"] for l in lines for w in l).upper()))
        left = re.sub(r"[/,.\-]+", "", left)
        if left:
            remarks.append(f"other text in the box: {shown}")
    else:
        refs = []
        for line in text.split("\n"):
            for r in tp_numbers(line):
                if r not in refs:
                    refs.append(r)
        remarks.append("'Other References' box not found - TP numbers taken from the whole page")
    return {"kind": "INVOICE", "invoice_no": invoice_no, "date": date, "refs": refs, "shown": shown,
            "remarks": remarks}


# ======================= finding the TP PDFs =======================
def pick_clip(page):
    orig, dup = page.search_for("ORIGINAL COPY"), page.search_for("DUPLICATE COPY")
    if not (orig and dup):
        return None
    r = page.rect
    if abs(orig[0].x0 - dup[0].x0) >= abs(orig[0].y0 - dup[0].y0):
        mid = (r.x0 + r.x1) / 2
        return fitz.Rect(r.x0, r.y0, mid, r.y1) if orig[0].x0 < mid else fitz.Rect(mid, r.y0, r.x1, r.y1)
    mid = (r.y0 + r.y1) / 2
    return fitz.Rect(r.x0, r.y0, r.x1, mid) if orig[0].y0 < mid else fitz.Rect(r.x0, mid, r.x1, r.y1)


def read_tp_pass(path: Path):
    """The pass no inside a TP PDF (only if it is unambiguous), else None."""
    try:
        with fitz.open(path) as doc:
            if doc.page_count == 0:
                return None
            page = doc[0]
            if "TRANSIT PASS" not in re.sub(r"\s+", " ", page.get_text("text")).upper():
                return None
            text = page.get_text("text", clip=pick_clip(page))
    except Exception:
        return None
    tokens = set(TP_RE.findall(text))
    return norm_pass(tokens.pop()) if len(tokens) == 1 else None


def find_pdfs(folder: Path, recursive: bool, skip: Path):
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


def read_tp_summaries(folder: Path, recursive: bool, skip: Path):
    """{normalised pass no: [existing PDF paths]} from every pass_summary.xlsx in the TP folder(s)."""
    index, n_files = defaultdict(list), 0
    skip_res = skip.resolve() if skip else None
    folders = [folder]
    if recursive:
        folders = []
        for dirpath, dirnames, _ in os.walk(folder):
            if skip_res:
                dirnames[:] = [d for d in dirnames if (Path(dirpath) / d).resolve() != skip_res]
            folders.append(Path(dirpath))
    for d in folders:
        path = d / TP_SUMMARY_NAME
        if not path.is_file():
            continue
        try:
            ws = load_workbook(path, data_only=True).worksheets[0]
            heads = {c: str(ws.cell(1, c).value or "").strip().lower() for c in range(1, ws.max_column + 1)}
            c_pass = next((c for c, h in heads.items() if "pass" in h), None)
            name_cols = [c for want in ("file name", "filename", "new file name", "original file name")
                         for c, h in heads.items() if h == want]
            if c_pass is None or not name_cols:
                continue
            n_files += 1
            for r in range(2, ws.max_row + 1):
                pn = ws.cell(r, c_pass).value
                if blank(pn):
                    continue
                for c in name_cols:
                    nm = ws.cell(r, c).value
                    if not blank(nm) and (d / str(nm).strip()).is_file():
                        p = d / str(nm).strip()
                        if p not in index[norm_pass(pn)]:
                            index[norm_pass(pn)].append(p)
        except Exception:
            continue
    return index, n_files


class TPFinder:
    def __init__(self, folder: Path, recursive: bool, skip: Path):
        self.pdfs = find_pdfs(folder, recursive, skip)
        self.summary_index, self.n_summaries = read_tp_summaries(folder, recursive, skip)
        self.content_map = None

    def by_name(self, key):
        pat = re.compile(r"(?<![A-Z0-9])" + re.escape(key) + r"(?![A-Z0-9])")
        return [p for p in self.pdfs if pat.search(p.stem.upper().replace("/", "-"))]

    def find(self, tp):
        key = norm_pass(tp)
        files = [p for p in self.summary_index.get(key, []) if p.is_file()]
        if files:
            return files, "pass summary"
        files = self.by_name(key)
        if files:
            return files, "file name"
        if self.content_map is not None:
            files = self.content_map.get(key, [])
            if files:
                return files, "read inside PDF"
        return [], ""

    def build_content_map(self):
        self.content_map = defaultdict(list)
        for n, p in enumerate(self.pdfs, start=1):
            if n % 200 == 0:
                print(f"  read {n} of {len(self.pdfs)} ...")
            pn = read_tp_pass(p)
            if pn:
                self.content_map[pn].append(p)


def best_file(files):
    """Prefer a normal file over a DUPLICATE_ copy."""
    return sorted(files, key=lambda p: (p.name.upper().startswith("DUPLICATE"), p.name.lower()))[0]


# ======================= merging =======================
def merge_pdfs(parts, out_path: Path):
    out = fitz.open()
    try:
        for p in parts:
            with fitz.open(p) as src:
                out.insert_pdf(src)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out.save(out_path, garbage=3, deflate=True)
    finally:
        out.close()


def clean_file_name(invoice_no: str) -> str:
    return re.sub(r'[\\:*?"<>|]', "-", invoice_no.replace("/", "-")).strip(" .")


# ======================= summary workbook =======================
def read_old_rows(path: Path, out_dir: Path):
    """Rows of an earlier run whose merged PDF still exists - they are kept exactly as they are."""
    rows = []
    if not path.is_file():
        return rows
    try:
        ws = load_workbook(path).worksheets[0]
        head = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]
        for r in range(2, ws.max_row + 1):
            d = {h: ws.cell(r, c).value for c, h in enumerate(head, start=1) if h}
            merged = d.get("Merged PDF")
            if not blank(merged) and (out_dir / str(merged).strip()).is_file():
                rows.append(d)
    except Exception:
        pass
    return rows


def write_summary(path: Path, kept, rows, skipped):
    all_rows = [{**{k: r.get(k) for k in COLUMNS}, "_fill": None} for r in kept] + rows
    all_rows.sort(key=lambda r: (r["Invoice Date"] if isinstance(r["Invoice Date"], datetime) else datetime.max,
                                 str(r["Invoice No"])))
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    ws.append(COLUMNS)
    for r in all_rows:
        ws.append([r.get(c) for c in COLUMNS])
        ws.cell(ws.max_row, 1).number_format = "DD-MMM-YYYY"
        if r.get("_fill"):
            for c in ws[ws.max_row]:
                c.fill = r["_fill"]
        for col in (5, 6):                                       # a missing TP stands out
            if ws.cell(ws.max_row, col).value == "NOT FOUND":
                ws.cell(ws.max_row, col).fill = RED
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, w in enumerate([14, 20, 20, 20, 30, 30, 26, 30, 60], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    ws2 = wb.create_sheet("Skipped")
    ws2.append(["File Name", "Reason"])
    for name, reason in skipped:
        ws2.append([name, reason])
    if not skipped:
        ws2.append(["(none)", ""])
    for c in ws2[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER
    ws2.column_dimensions["A"].width = 44
    ws2.column_dimensions["B"].width = 60
    wb.save(path)


# ======================= main =======================
def main():
    folder = script_folder()
    out_dir = folder.parent / "output" / OUT_DIR
    summary_path = folder.parent / "output" / SUMMARY_NAME
    print("=== Invoice + Transit Pass Merger ===\n")
    print(f"Invoice folder: {folder}")

    tp_folder = choose_folder("Select the folder with the TRANSIT PASS PDFs (Cancel = this same folder)") or folder
    recursive = input("Also search inside the sub-folders of the TP folder? (y/n): ").strip().lower() == "y"
    print(f"TP folder: {tp_folder}\n")

    kept = read_old_rows(summary_path, out_dir)
    done_files = {str(r.get("Invoice File") or "").strip().casefold() for r in kept}

    pdfs = sorted((p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"),
                  key=lambda p: p.name.lower())
    todo = [p for p in pdfs if p.name.casefold() not in done_files]
    print(f"{len(pdfs)} PDF(s) here: {len(pdfs) - len(todo)} invoice(s) already merged earlier (left as they are), "
          f"{len(todo)} to check.")

    invoices, skipped, not_invoice = [], [], 0
    for n, p in enumerate(todo, start=1):
        if n % 200 == 0:
            print(f"  read {n} of {len(todo)} ...")
        try:
            r = read_invoice(p)
        except Exception as e:
            skipped.append((p.name, f"could not be read: {e}"))
            continue
        if r["kind"] == "INVOICE":
            r["path"] = p
            invoices.append(r)
        elif r["kind"] == "NOT INVOICE":
            not_invoice += 1
        else:
            skipped.append((p.name, r["reason"]))
    print(f"{len(invoices)} tax invoice(s) read" + (f", {not_invoice} other PDF(s) ignored (not tax invoices)" if not_invoice else "")
          + (f", {len(skipped)} could not be read" if skipped else "") + ".")
    if not invoices:
        if kept or skipped:
            write_summary(summary_path, kept, [], skipped)
        print("\nNothing to do.")
        return

    print("\nLooking for the TP PDFs...")
    finder = TPFinder(tp_folder, recursive, out_dir)
    print(f"  {len(finder.pdfs)} PDF(s) in the TP folder" +
          (f", {finder.n_summaries} pass_summary.xlsx file(s) found" if finder.n_summaries else ", no pass_summary.xlsx found"))
    needed = {t for inv in invoices for t in inv["refs"]}
    missing = [t for t in needed if not finder.find(t)[0]]
    if missing and finder.pdfs:
        print(f"  {len(missing)} TP number(s) not found by summary or file name - reading inside the PDFs...")
        finder.build_content_map()

    # ---- decide what to do with every invoice ----
    rows, used_targets = [], set()
    for inv in invoices:
        p = inv["path"]
        refs, notes = inv["refs"], list(inv["remarks"])
        chosen, shown = [], []
        for t in refs:
            files, how = finder.find(t)
            if files:
                f = best_file(files)
                chosen.append(f)
                shown.append(f.name)
                if len(files) > 1:
                    notes.append(f"{t}: {len(files)} PDFs found, used {f.name}")
                if how == "read inside PDF":
                    notes.append(f"{t}: found by reading inside the PDF")
            else:
                chosen.append(None)
                shown.append("NOT FOUND")
        row = {"Invoice Date": inv["date"], "Invoice No": inv["invoice_no"],
               "Other Ref 1": refs[0] if len(refs) > 0 else None, "Other Ref 2": refs[1] if len(refs) > 1 else None,
               "Matched TP1": shown[0] if len(shown) > 0 else None, "Matched TP2": shown[1] if len(shown) > 1 else None,
               "Merged PDF": None, "Invoice File": p.name, "_fill": None, "path": p, "parts": None}
        if len(refs) > 2:
            notes.append("more TP numbers: " + ", ".join(f"{t} -> {s}" for t, s in zip(refs[2:], shown[2:])))
        if inv["date"] is None:
            notes.append("invoice date not read")
        target = clean_file_name(inv["invoice_no"]) + ".pdf"
        if not refs:
            notes.append("no TP numbers found in 'Other References' - not merged")
            row["_fill"] = YELLOW
        elif any(c is None for c in chosen):
            notes.append("TP PDF missing - not merged yet (it is merged automatically on a later run once the PDF is available)")
            row["_fill"] = RED
        elif target.casefold() in used_targets:
            notes.append(f"invoice no {inv['invoice_no']} appears in more than one PDF - this one was not merged")
            row["_fill"] = YELLOW
        elif (out_dir / target).is_file():
            notes.append("a merged PDF with this name already exists - left as it is")
            row["Merged PDF"] = target
            used_targets.add(target.casefold())
        else:
            row["parts"] = [p] + chosen
            row["Merged PDF"] = target
            used_targets.add(target.casefold())
        row["Remarks"] = "; ".join(notes)
        rows.append(row)

    to_merge = [r for r in rows if r["parts"]]
    print("\n----- SUMMARY -----")
    print(f"Invoices checked        : {len(rows)}")
    print(f"Ready to merge          : {len(to_merge)}")
    print(f"TP missing (not merged) : {sum(1 for r in rows if r['_fill'] is RED)}")
    print(f"No TP numbers           : {sum(1 for r in rows if not r['Other Ref 1'])}")
    for r in rows[:12]:
        refs = [x for x in (r['Other Ref 1'], r['Other Ref 2']) if x]
        state = "merge" if r["parts"] else ("exists" if r["Merged PDF"] else "NOT merged")
        print(f"  {r['Invoice No']:<18} {', '.join(refs) or '(no TP)':<42} {state}")
    if len(rows) > 12:
        print(f"  ... and {len(rows) - 12} more")

    if input(f"\nMerge {len(to_merge)} invoice(s) into {out_dir.name}\\ and save the summary now? (y/n): ").strip().lower() != "y":
        print("Cancelled. Nothing was changed.")
        return
    if summary_path.exists():
        try:
            open(summary_path, "ab").close()
        except PermissionError:
            print(f"\n{SUMMARY_NAME} is open in Excel. Close it and run again. Nothing was changed.")
            return

    merged = failed = 0
    for r in to_merge:
        try:
            merge_pdfs(r["parts"], out_dir / r["Merged PDF"])
            merged += 1
        except Exception as e:
            failed += 1
            r["Remarks"] = (r["Remarks"] + "; " if r["Remarks"] else "") + f"MERGE FAILED: {e}"
            r["Merged PDF"] = None
            r["_fill"] = RED
    write_summary(summary_path, kept, rows, skipped)
    print(f"\nMerged {merged} invoice(s) into {out_dir}" + (f"  ({failed} FAILED - see Remarks)" if failed else ""))
    print(f"Summary saved: {summary_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nUnexpected error: {e}")
    input("\nPress Enter to exit...")


