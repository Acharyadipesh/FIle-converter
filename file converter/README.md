# FileForge — PDF, PowerPoint and Image Converter

A self-hosted iLovePDF-style file conversion and compression website built with HTML, CSS, JavaScript, Python and Flask.

## Features

- PDF → PPTX
- PPT/PPTX → PDF (requires LibreOffice)
- JPG/JPEG → PNG
- PNG → JPG
- PDF → JPG
- JPG → PDF
- PNG → PDF
- Merge multiple PDFs
- Image compression
- Supports JPG, JPEG, PNG and WEBP image compression
- Drag & drop UI
- 100 MB upload limit
- No third-party conversion API required

## Windows setup

1. Install Python 3.11+.
2. Open CMD/PowerShell in this folder.
3. Create a virtual environment:

   `python -m venv .venv`

4. Activate it:

   PowerShell: `.venv\Scripts\Activate.ps1`

   CMD: `.venv\Scripts\activate`

5. Install packages:

   `pip install -r requirements.txt`

6. For PPT → PDF, install LibreOffice.

   The application automatically searches common LibreOffice installation locations on Windows, so manually adding `soffice.exe` to PATH is usually not required.

   Official download: https://www.libreoffice.org/download/

7. Start the application:

   `python app.py`

8. Open:

   `http://127.0.0.1:5000`

## Image compression

FileForge supports image compression for:

- JPG
- JPEG
- PNG
- WEBP

The compression tool allows users to reduce image file size while maintaining a reasonable level of image quality.

## Important conversion note

PDF → PPTX in this starter converts each PDF page into a full-slide image. This preserves the visual appearance very well, but the text is not independently editable in PowerPoint.

For a production-grade editable PDF → PPT converter, use a specialized document conversion engine/API or build a much more complex PDF layout reconstruction pipeline.

## Privacy

FileForge does not use a third-party online conversion API.

When running locally, files are processed by the Flask application on the user's computer.

However, this is a server-side Flask application. If deployed publicly, uploaded files will be sent to the server for processing.

Do not describe the current version as a fully client-side or "files never leave your device" application.

## Production hardening

Before public deployment, add:

- authentication/rate limiting
- virus/malware scanning
- automatic cleanup of uploaded and generated files
- HTTPS
- per-user storage isolation
- background job queue (Celery/RQ)
- upload size and page-count limits
- logging/monitoring
- stronger filename/path validation
- secure temporary file handling
- protection against malicious files
- production WSGI server instead of Flask development server


Technologies
Python
Flask
HTML5
CSS3
JavaScript
Pillow
PyMuPDF
python-pptx
LibreOffice
Future improvements
PDF compression
PDF splitting
PDF page extraction
PDF page reordering
PDF rotation
PDF metadata removal
PDF password protection
WebP → JPG
WebP → PNG
Image resizing
Image cropping
Batch conversion
Batch image compression
OCR support
Better editable PDF → PPT conversion
Dark mode
Progress indicators
Automatic file cleanup
Client-side processing
Offline/PWA support

Author
Dipesh Acharya
GitHub: https://github.com/Acharyadipesh


## Project structure

```text
file converter/
│
├── app.py
├── requirements.txt
├── README.md
│
├── templates/
│   └── index.html
│
├── static/
│   ├── style.css
│   └── app.js
│
├── uploads/
│
└── outputs/
