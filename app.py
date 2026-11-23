"""Flask web application for user-configurable PDF form extraction."""
import os
from pathlib import Path

from flask import Flask, abort, redirect, render_template, request, send_file, url_for

import form_workflow

ROOT = Path(__file__).resolve().parent
app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["SECRET_KEY"] = os.environ.get("PDF_APP_SECRET", "local-development-key")
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/classify")
def classify():
    uploads = [item for item in request.files.getlist("pdf_files") if item and item.filename]
    if not uploads:
        return render_template("index.html", error="Choose PDFs or a folder containing PDFs first."), 400

    pdf_uploads = [
        item for item in uploads
        if Path((item.filename or "").replace("\\", "/")).suffix.lower() == ".pdf"
    ]
    if not pdf_uploads:
        return render_template("index.html", error="The selected files did not include any PDFs."), 400

    try:
        batch_id, categories, skipped = form_workflow.classify_uploads(pdf_uploads)
    except Exception as exc:
        app.logger.exception("Could not classify uploaded PDFs")
        return render_template("index.html", error=f"Could not read the uploaded files: {exc}"), 500

    if not batch_id:
        return render_template("index.html", error="No readable PDF files were found."), 400
    return render_template("types.html", batch_id=batch_id, categories=categories, skipped=skipped)


@app.get("/configure/<batch_id>/<category_id>")
def configure(batch_id, category_id):
    batch, category = form_workflow.get_category(batch_id, category_id)
    if not batch or not category:
        return render_template("index.html", error="This upload session expired. Please upload the PDFs again."), 410

    selected_id = request.args.get("document")
    selected = next((doc for doc in category["documents"] if doc["id"] == selected_id), category["documents"][0])
    return render_template(
        "configure.html",
        batch_id=batch_id,
        category=category,
        selected=selected,

    )


@app.get("/preview/<batch_id>/<document_id>")
def preview_pdf(batch_id, document_id):
    batch = form_workflow.BATCHES.get(batch_id)
    if not batch:
        abort(404)
    document = batch["documents"].get(document_id)
    if not document or not document["path"].is_file():
        abort(404)
    return send_file(document["path"], mimetype="application/pdf", as_attachment=False, conditional=True)


@app.post("/extract-fields/<batch_id>/<category_id>")
def extract_fields(batch_id, category_id):
    batch, category = form_workflow.get_category(batch_id, category_id)
    if not batch or not category:
        return render_template("index.html", error="This upload session expired. Please upload the PDFs again."), 410

    selected_ids = set(request.form.getlist("fields"))
    fields = [
        field for field in category["fields"]
        if field["id"] in selected_ids or field.get("sort_key")
    ]
    if not fields:
        return render_template("configure.html", batch_id=batch_id, category=category,
                               selected=category["documents"][0],
                               error="Tick at least one detected field to extract."), 400

    try:
        result = form_workflow.create_workbook(batch_id, category_id, fields)
        form_workflow.cleanup_batch(batch_id)
    except Exception as exc:
        app.logger.exception("Could not extract selected form type")
        return render_template("configure.html", batch_id=batch_id, category=category,
                               selected=category["documents"][0],
                               error=f"Could not create the workbook: {exc}"), 500
    return redirect(url_for("show_results", result_id=result["id"]))

@app.get("/results/<result_id>")
def show_results(result_id):
    result = form_workflow.RESULTS.get(result_id)
    if not result:
        abort(404)
    return render_template("result.html", result=result)


@app.get("/download/<result_id>")
def download(result_id):
    result = form_workflow.RESULTS.get(result_id)
    if not result or not result["path"].is_file():
        abort(404)
    return send_file(result["path"], as_attachment=True, download_name=result["filename"])


@app.errorhandler(413)
def upload_too_large(_error):
    return render_template("index.html", error="The upload exceeds the 200 MB limit. Try fewer or smaller PDFs."), 413


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)





