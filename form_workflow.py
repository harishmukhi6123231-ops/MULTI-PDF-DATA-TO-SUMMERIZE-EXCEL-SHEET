"""Group PDFs by form layout, detect their labeled fields, and export checked fields."""
from collections import Counter
from pathlib import Path
import hashlib
import re
import tempfile
import uuid

ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "output"
BATCHES = {}
RESULTS = {}

STOP_WORDS = {
    "the", "and", "for", "from", "with", "this", "that", "are", "was", "were", "has", "have", "had",
    "not", "all", "see", "page", "copy", "original", "duplicate", "form", "government", "department",
    "name", "date", "number", "no", "of", "to", "in", "on", "by", "as", "at", "is", "be", "or",
}
SORT_PRIORITY = ("pass no", "pass number", "invoice no", "invoice number", "application no", "reference no", "permit no")


def read_pdf(pdf_path: Path):
    """Return text from every page and the first page separately."""
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    pages = [(page.extract_text() or "") for page in reader.pages]
    return "\n\n".join(pages), (pages[0] if pages else "")


def signature_for(text: str):
    """Make a value-insensitive signature from headings and labels."""
    labels = []
    for raw in text.splitlines()[:100]:
        line = re.sub(r"\s+", " ", raw).strip(" \t:|-")
        if not line:
            continue
        label = line.split(":", 1)[0].strip() if ":" in line else line
        label = re.sub(r"\b\d+(?:[./-]\d+)*\b", " ", label)
        label = re.sub(r"[^\w ]+", " ", label, flags=re.UNICODE)
        label = re.sub(r"\s+", " ", label).strip().lower()
        tokens = re.findall(r"[a-z]{2,}", label)
        if tokens and len(tokens) <= 9 and any(word not in STOP_WORDS for word in tokens):
            labels.append(" ".join(tokens))
    tokens = {word for label in labels for word in re.findall(r"[a-z]{2,}", label) if word not in STOP_WORDS}
    return tokens, labels


def _similarity(left, right):
    if not left or not right:
        return 0.0
    return len(left & right) / len(left | right)


def _candidate_id(key):
    return hashlib.sha1(key.casefold().encode("utf-8")).hexdigest()[:12]


def detect_field_candidates(text):
    """Find visible field labels and common pass/invoice identifiers."""
    found = {}

    def add(label, key=None):
        label = re.sub(r"\s+", " ", label).strip(" \t:|-–—.")
        label = re.sub(r"^\d+\s*[.)]\s*", "", label).strip()
        label = re.sub(r"^(?:N/?A|NOT APPLICABLE|[-–—]{2,})\s+", "", label, flags=re.I).strip()
        label = re.sub(r"\s+", " ", label)
        if re.search(r"\bRoute\s*$", label, re.I) and len(label.split()) > 1:
            label = "Route"
        if not label or len(label) > 78 or not re.search(r"[A-Za-z]", label):
            return
        low = label.casefold()
        if low.startswith(("signature", "office seal", "see rule", "government of", "department of")):
            return
        tokens = re.findall(r"[A-Za-z]{2,}", label)
        if not tokens or len(tokens) > 13:
            return
        if low in {"original copy", "duplicate copy", "length breadth height"}:
            return
        key = key or label
        identity = key.casefold()
        found.setdefault(identity, {"id": _candidate_id(identity), "column": label, "key": key, "found_in": 0})

    raw_lines = [line.strip() for line in text.splitlines() if line.strip()]
    expanded = []
    index = 0
    while index < len(raw_lines):
        line = re.sub(r"\s+", " ", raw_lines[index]).strip()
        if re.match(r"^\d+\s*[.)]", line) and ":" not in line and not re.search(r"\bPass\s*No\b", line, re.I):
            cursor = index + 1
            combined = line
            while cursor < len(raw_lines) and ":" not in combined and cursor <= index + 3 and len(combined) < 150:
                combined += " " + re.sub(r"\s+", " ", raw_lines[cursor]).strip()
                cursor += 1
            if ":" in combined:
                line = combined
                index = cursor - 1
        expanded.append(line)
        index += 1

    for line in expanded:
        if ":" not in line:
            continue
        parts = line.split(":")
        add(parts[0])
        # Handle adjacent fields such as "Permit No.: NA Date: NA".
        for middle in parts[1:-1]:
            marker = re.match(r"^\s*(?:N/?A|NOT APPLICABLE|[-–—]{2,})\s+(.+?)\s*$", middle, re.I)
            if marker:
                add(marker.group(1))

    if re.search(r"\bPass\s*No\.?\b", text, re.I):
        add("Pass No.", "__pass_no__")
    if re.search(r"\bInvoice\s*No\.?\b", text, re.I):
        add("Invoice No.", "__invoice_no__")
    if re.search(r"\bApplication\s*No\.?\b", text, re.I):
        add("Application No.", "__application_no__")

    vehicle = re.search(r"\b[A-Z]{2}\d{1,2}[A-Z]{1,3}\d{3,4}\b", text, re.I)
    if vehicle:
        carrier_labels = [
            (identity, field) for identity, field in found.items()
            if "measurement" in field["column"].casefold() and "carrier" in field["column"].casefold()
        ]
        for identity, _field in carrier_labels:
            found.pop(identity, None)
        add("Carrier Vehicle No." if carrier_labels else "Vehicle Registration No.", "__vehicle_registration__")

    return list(found.values())

def _choose_sort_field(fields):
    for phrase in SORT_PRIORITY:
        match = next((field for field in fields if phrase in field["column"].casefold()), None)
        if match:
            return match["id"]
    return None


def classify_uploads(uploads):
    """Save uploaded PDFs temporarily and group similar document forms."""
    from werkzeug.utils import secure_filename

    temp_obj = tempfile.TemporaryDirectory(prefix="pdf_form_batch_")
    temp_dir = Path(temp_obj.name)
    documents, groups, skipped = {}, [], []

    try:
        for index, upload in enumerate(uploads, start=1):
            original_name = (upload.filename or f"document_{index}.pdf").replace("\\", "/")
            display_name = "/".join(part for part in original_name.split("/") if part not in ("", ".", ".."))
            if Path(display_name).suffix.lower() != ".pdf":
                skipped.append({"name": display_name or f"document_{index}", "reason": "Not a PDF file."})
                continue
            safe_name = secure_filename(Path(display_name).name) or f"document_{index}.pdf"
            saved_path = temp_dir / f"{index:04d}_{safe_name}"
            upload.save(str(saved_path))
            try:
                full_text, first_page = read_pdf(saved_path)
                signature, labels = signature_for(first_page)
                read_error = "" if full_text.strip() else "No selectable text was found; this file may need OCR."
            except Exception as exc:
                full_text, first_page, signature, labels = "", "", set(), []
                read_error = f"Could not read PDF text: {exc}"

            best_group, best_score = None, 0.0
            for group in groups:
                score = _similarity(signature, group["signature"])
                if score > best_score:
                    best_group, best_score = group, score
            if best_group is None or best_score < 0.30:
                best_group = {"id": f"type-{len(groups) + 1}", "signature": set(), "label_counts": Counter(), "documents": []}
                groups.append(best_group)
            best_group["signature"].update(signature)
            best_group["label_counts"].update(set(labels))

            doc_id = f"doc-{index}"
            document = {
                "id": doc_id, "name": display_name or safe_name, "path": saved_path,
                "text": full_text, "first_page": first_page, "category_id": best_group["id"],
                "read_error": read_error,
            }
            documents[doc_id] = document
            best_group["documents"].append(doc_id)

        if not documents:
            temp_obj.cleanup()
            return None, [], skipped

        categories = []
        for number, group in enumerate(groups, start=1):
            group_docs = [documents[doc_id] for doc_id in group["documents"]]
            candidate_map = {}
            for document in group_docs:
                seen = set()
                for field in detect_field_candidates(document["first_page"]):
                    if field["id"] in seen:
                        continue
                    seen.add(field["id"])
                    existing = candidate_map.setdefault(field["id"], {**field, "found_in": 0})
                    existing["found_in"] += 1
            fields = sorted(candidate_map.values(), key=lambda field: (-field["found_in"], field["column"].casefold()))
            sort_key = _choose_sort_field(fields)
            for field in fields:
                field["sort_key"] = field["id"] == sort_key
            is_unreadable = all(not document["text"].strip() for document in group_docs)
            categories.append({
                "id": group["id"],
                "name": "Scanned / unreadable forms" if is_unreadable else f"Form Type {number}",
                "count": len(group_docs),
                "documents": group_docs,
                "labels": [field["column"] for field in fields[:6]],
                "fields": fields,
                "sort_key": sort_key,
                "unreadable": is_unreadable,
            })

        batch_id = uuid.uuid4().hex
        BATCHES[batch_id] = {"temp": temp_obj, "documents": documents, "categories": categories, "skipped": skipped}
        return batch_id, categories, skipped
    except Exception:
        temp_obj.cleanup()
        raise


def get_category(batch_id, category_id):
    batch = BATCHES.get(batch_id)
    if not batch:
        return None, None
    category = next((item for item in batch["categories"] if item["id"] == category_id), None)
    return batch, category


def _natural_sort_key(value):
    parts = re.split(r"(\d+)", str(value or ""))
    return [(0, int(part)) if part.isdigit() else (1, part.casefold()) for part in parts]


def extract_value(text: str, key, stop_labels=()):
    """Extract the value beside a selected detected label."""
    if key == "__pass_no__":
        match = re.search(r"\bPass\s*No\.?\s*:?\s*([A-Z][A-Z0-9/-]{5,})", text, re.I)
        return match.group(1).rstrip(".,:") if match else ""
    if key == "__invoice_no__":
        match = re.search(r"\bInvoice\s*No\.?\s*:?\s*([A-Z0-9][A-Z0-9/-]{3,})", text, re.I)
        return match.group(1).rstrip(".,:") if match else ""
    if key == "__application_no__":
        match = re.search(r"\bApplication\s*No\.?\s*:?\s*([A-Z0-9][A-Z0-9/-]{3,})", text, re.I)
        return match.group(1).rstrip(".,:") if match else ""
    if key == "__vehicle_registration__":
        match = re.search(r"\b[A-Z]{2}\d{1,2}[A-Z]{1,3}\d{3,4}\b", text, re.I)
        return match.group(0).upper() if match else ""

    flat = re.sub(r"\s+", " ", text)
    aliases = [part.strip() for part in re.split(r"[|,]", key) if part.strip()]
    aliases.sort(key=len, reverse=True)
    other_labels = [label for label in stop_labels if label and label != key and not label.startswith("__")]
    other_labels.extend(("Signature of", "Office Seal", "ORIGINAL COPY", "DUPLICATE COPY"))

    for alias in aliases:
        pattern = r"\s+".join(re.escape(part) for part in alias.split())
        for match in re.finditer(pattern, flat, re.I):
            tail = flat[match.end():]
            if alias.casefold() == "date" and re.match(r"\s*&\s*time\b", tail, re.I):
                continue
            remainder = tail.lstrip(" 	:. -#")
            if not remainder:
                continue
            stop_at = len(remainder)
            for label in other_labels:
                stop_pattern = r"\s+".join(re.escape(part) for part in label.split())
                boundary = re.search(r"(?:^|\s+)(?:\d+\s*[.)]\s*)?" + stop_pattern + r"\s*:?", remainder, re.I)
                if boundary:
                    stop_at = min(stop_at, boundary.start())
            numbered_boundary = re.search(r"\s+\d+\s*[.)]\s+[A-Z][^:]{1,90}\s*:", remainder)
            if numbered_boundary:
                stop_at = min(stop_at, numbered_boundary.start())
            value = remainder[:stop_at].strip(" 	:. -#")
            if value:
                return value[:400]
    return ""

def create_workbook(batch_id, category_id, fields):
    """Export checked fields, automatically adding and sorting by the detected key."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    batch, category = get_category(batch_id, category_id)
    if not batch or not category:
        raise KeyError("This upload session expired. Upload the PDFs again.")
    fields = list(fields)
    sort_field = next((field for field in fields if field.get("sort_key")), None)
    if category["sort_key"] and not sort_field:
        sort_field = next((field for field in category["fields"] if field["id"] == category["sort_key"]), None)
        if sort_field:
            fields.insert(0, sort_field)

    rows = []
    stop_labels = [field["key"] for field in category["fields"]]
    for document in category["documents"]:
        values = [extract_value(document["text"], field["key"], stop_labels) for field in fields]
        row = {"Source PDF": document["name"], "Form Type": category["name"]}
        row.update({field["column"]: value for field, value in zip(fields, values)})
        rows.append(row)

    if sort_field:
        rows.sort(key=lambda row: _natural_sort_key(row.get(sort_field["column"])))

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Extracted Data"
    headers = ["Source PDF", "Form Type"] + [field["column"] for field in fields]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="334FD1")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, len(rows) + 1)}"
    for row in rows:
        sheet.append([row.get(header, "") for header in headers])
    for column, header in enumerate(headers, start=1):
        values = [len(str(header))] + [len(str(sheet.cell(row, column).value or "")) for row in range(2, sheet.max_row + 1)]
        sheet.column_dimensions[get_column_letter(column)].width = min(max(max(values) + 2, 14), 42)
    sheet.row_dimensions[1].height = 32

    result_id = uuid.uuid4().hex
    safe_type = re.sub(r"[^A-Za-z0-9_-]+", "_", category["name"]).strip("_") or "Form"
    filename = f"Extracted_{safe_type}_{result_id[:8]}.xlsx"
    result_path = OUTPUT_DIR / filename
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        workbook.save(result_path)
    except PermissionError:
        result_path = Path(tempfile.gettempdir()) / f"{result_id}_{filename}"
        workbook.save(result_path)
    finally:
        workbook.close()

    result = {
        "id": result_id, "category": category["name"], "count": len(rows),
        "fields": fields, "rows": rows, "filename": filename, "path": result_path,
        "sort_by": sort_field["column"] if sort_field else "",
    }
    RESULTS[result_id] = result
    return result


def cleanup_batch(batch_id):
    batch = BATCHES.pop(batch_id, None)
    if batch:
        batch["temp"].cleanup()



