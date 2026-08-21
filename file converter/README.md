# FileForge — PDF, PowerPoint and Image Converter

A self-hosted iLovePDF-style starter website built with HTML, CSS, JavaScript and Flask.

## Features

- PDF → PPTX
- PPT/PPTX → PDF (requires LibreOffice)
- JPG/JPEG → PNG
- PNG → JPG
- PDF → JPG (all pages are returned in a ZIP when there are multiple pages)
- JPG → PDF
- PNG → PDF
- Merge multiple PDFs
- Drag & drop UI
- 100 MB upload limit
- No third-party conversion API required

## Windows setup

1. Install Python 3.11+.
2. Open CMD/PowerShell in this folder.
3. Create a virtual environment:

   `python -m venv .venv`

4. Activate it:

   PowerShell:
   `.venv\Scripts\Activate.ps1`

   CMD:
   `.venv\Scripts\activate`

5. Install packages:

   `pip install -r requirements.txt`

6.  For PPT → PDF, install LibreOffice.

   The application automatically searches common
   LibreOffice installation locations on Windows,
   so manually adding soffice.exe to PATH is usually
   not required.

   Official download:
   https://www.libreoffice.org/download/

7. Start:

   `python app.py`

8. Open:

   `http://127.0.0.1:5000`

## Important conversion note

PDF → PPTX in this starter converts each PDF page into a full-slide image. This preserves the visual appearance very well, but the text is not independently editable in PowerPoint.

For a production-grade editable PDF → PPT converter, use a specialized document conversion engine/API or build a much more complex PDF layout reconstruction pipeline.

## Production hardening

Before public deployment, add:
- authentication/rate limiting
- virus/malware scanning
- automatic cleanup of old files
- HTTPS
- per-user storage isolation
- background job queue (Celery/RQ)
- upload size and page-count limits
- logging/monitoring
- stronger filename/path validation
- privacy policy and terms
