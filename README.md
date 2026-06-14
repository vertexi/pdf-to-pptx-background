# PDF to PPTX Background

Convert each page of a PDF into a PowerPoint slide whose rendered page is the
slide's true picture background.

Unlike tools that insert a full-slide picture shape, the generated background
cannot be accidentally selected or moved while editing the presentation.

## Requirements

- Python 3.10 or newer
- PyMuPDF
- python-pptx

PowerPoint, VBA, and Windows COM automation are not required.

## Installation

```powershell
python -m pip install -r requirements.txt
```

Tkinter is included with standard Windows Python installations.

## Graphical Interface

Start the Tkinter application:

```powershell
python .\pdf_to_pptx_background_gui.py
```

The interface provides:

- PDF and output file selection
- Rendering resolution
- Start page and page count
- Output overwrite control
- Conversion progress and completion status

The conversion runs on a worker thread, so the interface remains responsive.

## Windows EXE

A packaged executable can run on Windows without a separate Python
installation. Build it with:

```powershell
python -m pip install -r requirements-build.txt
.\build_exe.ps1
```

If the desired Python executable is not named `python`, pass it explicitly:

```powershell
.\build_exe.ps1 -Python "C:\Path\To\python.exe"
```

The executable is created at:

```text
dist\PDF-to-PPTX-Background.exe
```

The EXE is unsigned, so Windows SmartScreen may show a warning when it is
downloaded on another computer.

## Command-Line Usage

```powershell
python .\pdf_to_pptx_background.py .\slides.pdf
```

The default output is:

```text
slides-background.pptx
```

Choose an output path:

```powershell
python .\pdf_to_pptx_background.py .\slides.pdf `
  --output .\presentation.pptx
```

Use a lower rendering resolution to reduce the PPTX file size:

```powershell
python .\pdf_to_pptx_background.py .\slides.pdf --resolution 150
```

Convert only part of a PDF:

```powershell
python .\pdf_to_pptx_background.py .\slides.pdf `
  --from-page 3 --count 5
```

Replace an existing output file:

```powershell
python .\pdf_to_pptx_background.py .\slides.pdf --overwrite
```

Run `python .\pdf_to_pptx_background.py --help` for all options.

## Notes

- Each PDF page is rasterized to PNG before being embedded as a background.
- Text and vector graphics in the PDF are therefore not editable in PowerPoint.
- The slide aspect ratio follows the first selected PDF page.
