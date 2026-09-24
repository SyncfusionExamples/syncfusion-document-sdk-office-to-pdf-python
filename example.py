"""Run all seven operations on files in the documents folder.

Use: python3 example.py --trial
Existing output files are replaced by the current .NET worker.
"""
import argparse
from pathlib import Path
from document_sdk import load_service

# Read --trial from the command line; it permits evaluation without a license key.
parser = argparse.ArgumentParser(description="Run all seven document operations.")
parser.add_argument("--trial", action="store_true", help="Allow evaluation without a license key")
args = parser.parse_args()

# Create the Python client. Each method call starts a local .NET worker.
sdk = load_service(allow_trial=args.trial)

# Use the documents folder under the terminal's current directory.
# It must already contain input.docx, input.xlsx and input.pptx.
folder = Path("documents").resolve()

# Add classification headers and footers to the Word document.
sdk.MarkWord(str(folder / "input.docx"), str(folder / "marked.docx"), "CONFIDENTIAL")

# Mark Excel print headers/footers; worksheet cells remain unchanged.
sdk.MarkExcel(str(folder / "input.xlsx"), str(folder / "marked.xlsx"), "CONFIDENTIAL")

# Add visible classification text boxes to every PowerPoint slide.
sdk.MarkPowerPoint(str(folder / "input.pptx"), str(folder / "marked.pptx"), "CONFIDENTIAL")

# Convert the marked Word document to PDF.
sdk.WordToPdf(str(folder / "marked.docx"), str(folder / "word.pdf"))

# Convert the marked workbook to PDF using its print settings.
sdk.ExcelToPdf(str(folder / "marked.xlsx"), str(folder / "excel.pdf"))

# Convert the marked presentation to PDF.
sdk.PowerPointToPdf(str(folder / "marked.pptx"), str(folder / "powerpoint.pdf"))

# Add RESTRICTED diagonally to every page of the Word-generated PDF.
sdk.WatermarkPdf(str(folder / "word.pdf"), str(folder / "watermarked.pdf"), "RESTRICTED")
