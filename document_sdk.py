"""Run Syncfusion document operations from Python 3.9+ through a local .NET worker."""
import argparse
import json
import shutil
import subprocess
from pathlib import Path


class DocumentService:
    """Python 3.9-compatible client for a local .NET process (no Python.NET)."""

    def __init__(self, dll, dotnet, allow_trial):
        """Store the worker path, .NET executable and evaluation setting for later calls."""
        self._dll = dll
        self._dotnet = dotnet
        self._allow_trial = allow_trial

    def _call(self, operation, source, destination, label=None):
        """Run one document operation and raise RuntimeError if the .NET worker fails."""
        # Build the request as JSON. Resolve paths relative to the terminal directory.
        request = {"Operation": operation, "Input": str(Path(source).resolve()),
                   "Output": str(Path(destination).resolve()), "Label": label,
                   "AllowTrial": self._allow_trial}
        # No shell: paths and labels are JSON data, never executable commands.
        result = subprocess.run(
            [self._dotnet, str(self._dll)], input=json.dumps(request),
            text=True, encoding="utf-8", capture_output=True, check=False,
        )
        # A nonzero exit code means the worker failed; show its error to the caller.
        if result.returncode != 0:
            detail = result.stderr.strip() or "Worker exited with code {}".format(result.returncode)
            raise RuntimeError(detail)

    def MarkWord(self, source, destination, label):
        """Add classification text to Word headers and footers; save the DOCX output."""
        self._call("MarkWord", source, destination, label)

    def MarkExcel(self, source, destination, label):
        """Replace Excel center print headers and footers with the classification label."""
        self._call("MarkExcel", source, destination, label)

    def MarkPowerPoint(self, source, destination, label):
        """Add or update classification text boxes at the top and bottom of each slide."""
        self._call("MarkPowerPoint", source, destination, label)

    def WordToPdf(self, source, destination):
        """Convert a DOCX file to PDF using the Syncfusion Word renderer."""
        self._call("WordToPdf", source, destination)

    def ExcelToPdf(self, source, destination):
        """Convert an XLSX workbook to PDF using its print settings."""
        self._call("ExcelToPdf", source, destination)

    def PowerPointToPdf(self, source, destination):
        """Convert a PPTX presentation to PDF using the Syncfusion presentation renderer."""
        self._call("PowerPointToPdf", source, destination)

    def WatermarkPdf(self, source, destination, label):
        """Draw a translucent diagonal label on every page of an existing PDF."""
        self._call("WatermarkPdf", source, destination, label)


def load_service(bundle=None, *, allow_trial=True):
    """Return a local worker client; no pip packages or CLR embedding required.

    The worker inherits SYNCFUSION_LICENSE_KEY from the environment.
    Keys are never sent as command-line arguments.
    """
    # Find the published worker beside this Python file, unless a folder was supplied.
    bundle = Path(bundle or Path(__file__).parent / "artifacts").resolve()
    dll = bundle / "DocumentBridge.dll"
    config = bundle / "DocumentBridge.runtimeconfig.json"
    # Both files are required to start the framework-dependent .NET worker.
    if not dll.is_file() or not config.is_file():
        raise FileNotFoundError("Publish DocumentBridge into {} first.".format(bundle))
    # Locate the installed .NET runtime through the operating system PATH.
    dotnet = shutil.which("dotnet")
    if dotnet is None:
        raise RuntimeError("Install the .NET 8 runtime/SDK and put dotnet on PATH.")
    # This only creates a client; the worker starts when an operation is called.
    # allow_trial=True permits evaluation without a key. The worker handles licensing.
    return DocumentService(dll, dotnet, allow_trial)


def main():
    """Read CLI arguments, validate file paths and run the selected operation."""
    # Example: python3 document_sdk.py mark-word input.docx marked.docx --trial
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("operation", choices=["mark-word", "mark-excel", "mark-ppt",
                        "word-to-pdf", "excel-to-pdf", "ppt-to-pdf", "watermark-pdf"])
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--label", default="CONFIDENTIAL")
    parser.add_argument("--bundle", type=Path)
    parser.add_argument("--trial", action="store_true")
    args = parser.parse_args()
    # Map each CLI command to its method, expected extensions and label requirement.
    methods = {
        "mark-word": ("MarkWord", ".docx", ".docx", True),
        "mark-excel": ("MarkExcel", ".xlsx", ".xlsx", True),
        "mark-ppt": ("MarkPowerPoint", ".pptx", ".pptx", True),
        "word-to-pdf": ("WordToPdf", ".docx", ".pdf", False),
        "excel-to-pdf": ("ExcelToPdf", ".xlsx", ".pdf", False),
        "ppt-to-pdf": ("PowerPointToPdf", ".pptx", ".pdf", False),
        "watermark-pdf": ("WatermarkPdf", ".pdf", ".pdf", True),
    }
    method, input_ext, output_ext, needs_label = methods[args.operation]
    # Check input/output paths before starting the .NET worker.
    # Existing outputs are allowed: the current C# code overwrites them.
    source, destination = args.input.resolve(), args.output.resolve()
    if source.suffix.lower() != input_ext or destination.suffix.lower() != output_ext:
        parser.error(f"This operation requires {input_ext} input and {output_ext} output.")
    if not source.is_file():
        parser.error("Input file does not exist.")
    if not destination.parent.is_dir():
        parser.error("Output directory does not exist.")
    # Pass the selected method its paths and, for marking operations, the label.
    sdk = load_service(args.bundle, allow_trial=args.trial)
    parameters = [str(source), str(destination)]
    if needs_label:
        parameters.append(args.label)
    getattr(sdk, method)(*parameters)
    # Print the output location only after successful completion.
    print(destination)



# Run the CLI only when this file is executed, not when imported by example.py.
if __name__ == "__main__":
    main()
