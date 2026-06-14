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

## Usage

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

