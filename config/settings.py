from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
INPUT_PDFS = PROJECT_ROOT / "input_pdfs"
SAMPLE_PDFS = PROJECT_ROOT / "sample_pdfs"
INPUT_DATA = PROJECT_ROOT / "input_data"
OUTPUT = PROJECT_ROOT / "output"
TP_TEMPLATE = INPUT_DATA / "TP Summary.xlsx"
TP_SUMMARY = OUTPUT / "TP Summary.xlsx"
