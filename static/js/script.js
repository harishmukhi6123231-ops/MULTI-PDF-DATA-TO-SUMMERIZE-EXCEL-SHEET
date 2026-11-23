const pdfPicker = document.getElementById("pdfFiles");
const folderPicker = document.getElementById("folderFiles");
const dropzone = document.getElementById("dropzone");
const fileList = document.getElementById("fileList");
const uploadButton = document.getElementById("submitButton");

function currentFiles() {
  return [...(pdfPicker?.files || []), ...(folderPicker?.files || [])]
    .filter((file) => file.name.toLowerCase().endsWith(".pdf"));
}

function showFiles() {
  if (!fileList || !uploadButton) return;
  const files = currentFiles();
  fileList.replaceChildren();
  files.slice(0, 12).forEach((file) => {
    const row = document.createElement("div");
    row.className = "file-item";
    const name = document.createElement("span");
    name.textContent = file.webkitRelativePath || file.name;
    const size = document.createElement("span");
    size.className = "file-size";
    size.textContent = file.size < 1024 * 1024
      ? Math.max(1, Math.round(file.size / 1024)) + " KB"
      : (file.size / (1024 * 1024)).toFixed(1) + " MB";
    row.append(name, size);
    fileList.append(row);
  });
  if (files.length > 12) {
    const more = document.createElement("div");
    more.className = "muted small";
    more.textContent = "and " + (files.length - 12) + " more PDF files";
    fileList.append(more);
  }
  uploadButton.disabled = files.length === 0;
}

pdfPicker?.addEventListener("change", () => {
  if (pdfPicker.files.length && folderPicker) folderPicker.value = "";
  showFiles();
});
folderPicker?.addEventListener("change", () => {
  if (folderPicker.files.length && pdfPicker) pdfPicker.value = "";
  showFiles();
});

if (dropzone && pdfPicker) {
  ["dragenter", "dragover"].forEach((eventName) => dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("is-dragging");
  }));
  ["dragleave", "drop"].forEach((eventName) => dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("is-dragging");
  }));
  dropzone.addEventListener("drop", (event) => {
    const transfer = new DataTransfer();
    [...event.dataTransfer.files]
      .filter((file) => file.name.toLowerCase().endsWith(".pdf"))
      .forEach((file) => transfer.items.add(file));
    pdfPicker.files = transfer.files;
    if (folderPicker) folderPicker.value = "";
    showFiles();
  });
}

const fieldChecks = [...document.querySelectorAll("[data-field-check]")];
const selectedCount = document.getElementById("selectedCount");

function updateSelectedCount() {
  if (selectedCount) {
    selectedCount.textContent = String(fieldChecks.filter((input) => input.checked).length);
  }
}

document.querySelector("[data-select-all]")?.addEventListener("click", () => {
  fieldChecks.filter((input) => !input.disabled).forEach((input) => { input.checked = true; });
  updateSelectedCount();
});
document.querySelector("[data-clear-all]")?.addEventListener("click", () => {
  fieldChecks.filter((input) => !input.disabled).forEach((input) => { input.checked = false; });
  updateSelectedCount();
});
fieldChecks.forEach((input) => input.addEventListener("change", updateSelectedCount));
updateSelectedCount();

const previewSelect = document.getElementById("previewSelect");
const pdfFrame = document.getElementById("pdfFrame");
previewSelect?.addEventListener("change", () => {
  if (pdfFrame) pdfFrame.src = previewSelect.value;
});

document.getElementById("fieldForm")?.addEventListener("submit", () => {
  const button = document.querySelector("#fieldForm button[type='submit']");
  if (button) {
    button.disabled = true;
    button.textContent = "Extracting checked fields…";
  }
});
document.getElementById("uploadForm")?.addEventListener("submit", () => {
  if (uploadButton) {
    uploadButton.disabled = true;
    uploadButton.textContent = "Grouping PDFs…";
  }
});
