# Python + Syncfusion Document SDK

This sample uses Python to call Syncfusion .NET libraries to:

- Add classification labels to Word, Excel and PowerPoint documents.
- Convert these documents to PDF.
- Add a text watermark to every page of a PDF.

Processing runs locally. Microsoft Office is not required.

## How to run

### 1. Prerequisites

- Python 3.9 or newer (tested with Python 3.9.6).
- .NET 8 SDK.

No Python packages or pip installation steps are required.

### 2. Build the .NET worker

Open a terminal in the project folder and run:

```bash
dotnet publish DocumentBridge\DocumentBridge.csproj -c Release -r win-x64 --self-contained false -o artifacts
```

Use the runtime identifier for your machine:

| Machine | Runtime identifier |
| --- | --- |
| Apple Silicon Mac | `osx-arm64` |
| Intel Mac | `osx-x64` |
| Linux x64 | `linux-x64` |
| Windows x64 | `win-x64` |

On Windows, enter the publish command on one line. Keep the entire generated `artifacts` folder. Publish again after changing any C# code.

For Debian/Ubuntu Linux, also install the rendering dependencies:

```bash
sudo apt-get install -y libfontconfig1 libfreetype6 fontconfig fonts-liberation
```

Install the fonts used by your documents for accurate PDF rendering. Runtime testing was performed on macOS; verify the sample on your target platform.

### 3. Add input files

Create a `documents` folder inside the project folder and add:

```text
documents/
  input.docx
  input.xlsx
  input.pptx
```

### 4. Run the example

From the project folder:

```bash
python example.py --trial
```

On Windows, use `python` if `python3` is unavailable.

`--trial` allows evaluation without a license key. Generated documents may contain Syncfusion evaluation watermarks.

The results are saved in `documents`:

```text
marked.docx
marked.xlsx
marked.pptx
word.pdf
excel.pdf
powerpoint.pdf
watermarked.pdf
```

Running the example again overwrites these output files. Input files remain unchanged. Excel labels appear in print headers and footers, not in worksheet cells.

For licensed use, set the `SYNCFUSION_LICENSE_KEY` environment variable and run without `--trial`.

## How Python and .NET connect

```text
example.py → document_sdk.py → local .NET worker → Syncfusion libraries
```

1. `example.py` selects the document operations and file paths.
2. `document_sdk.py` starts `dotnet artifacts/DocumentBridge.dll` as a separate local process for each operation.
3. Python sends the operation name, paths and optional label as JSON through the process's standard input.
4. The C# worker calls the Syncfusion libraries and saves the output file.
5. Python waits for completion and raises an error if the worker fails.

The bridge uses Python's built-in `subprocess` module. It does not require Python.NET, a web server or HTTP requests.

For example:

```python
from document_sdk import load_service

sdk = load_service(allow_trial=True)
sdk.MarkWord("documents/input.docx", "documents/marked.docx", "CONFIDENTIAL")
sdk.WordToPdf("documents/marked.docx", "documents/word.pdf")
```

`DocumentBridge` contains the C# processing code. `artifacts` contains the published worker that Python runs.
