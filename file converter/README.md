# FileForge — PDF, PowerPoint and Image Converter

A self-hosted file conversion and compression website built with HTML, CSS, JavaScript, Python and Flask.

## Live Demo

The application is live at: [https://f-ile-converter.vercel.app/](https://f-ile-converter.vercel.app/)

## Features

- **PDF → PPTX**: Convert PDF pages to a PowerPoint presentation.
- **JPG/JPEG ↔ PNG**: Convert between JPG and PNG formats.
- **PDF → JPG**: Extract all pages from a PDF as individual JPG images.
- **JPG/PNG → PDF**: Create a single PDF from multiple images.
- **Merge PDFs**: Combine multiple PDF files into one.
- **Image Compression**: Reduce the file size of JPG, JPEG, PNG, and WEBP images.
- **Drag & Drop UI**: An intuitive and user-friendly interface.
- **100 MB upload limit**: Handle reasonably sized files.
- **No third-party conversion API required**: All processing is done by the application itself.

## ⚠️ Important Feature Note

- **PPT/PPTX → PDF**: This feature is **not available** on the live Vercel deployment. It requires LibreOffice, which cannot be installed in Vercel's serverless environment. If you need this feature, please run the application locally or on a server where you can install LibreOffice.

## Running Locally (Windows Setup)
1. **Install Python 3.11+**.
2. **Open CMD/PowerShell** in this folder.
3. **Create a virtual environment**:
python -m venv .venv

4. **Activate it**:
- **PowerShell**: `.venv\Scripts\Activate.ps1`
- **CMD**: `.venv\Scripts\activate`
5. **Install packages**:
  pip install -r requirements.txt
6. **For PPT → PDF (Local Only)**, install LibreOffice.
- The application automatically searches common LibreOffice installation locations on Windows.
- Official download: [https://www.libreoffice.org/download/](https://www.libreoffice.org/download/)
7. **Start the application**:
  python api/index.py
8. **Open in your browser**:
  http://127.0.0.1:5000

## Image Compression

FileForge supports image compression for:
- JPG
- JPEG
- PNG
- WEBP

The compression tool allows users to reduce image file size while maintaining a reasonable level of image quality.

## Important Conversion Note

**PDF → PPTX** in this starter converts each PDF page into a full-slide image. This preserves the visual appearance very well, but the text is **not independently editable** in PowerPoint.

For a production-grade editable PDF → PPT converter, use a specialized document conversion engine/API or build a much more complex PDF layout reconstruction pipeline.

## Privacy & Security

- FileForge does **not** use a third-party online conversion API.
- When running locally, files are processed by the Flask application on the user's computer.
- **However**, this is a server-side Flask application. If deployed publicly (like on Vercel), uploaded files will be sent to the server for processing. Do not describe the current version as a fully client-side or "files never leave your device" application.

## Production Hardening

Before public deployment, add:
- Authentication / Rate limiting
- Virus / Malware scanning
- Automatic cleanup of uploaded and generated files
- HTTPS
- Per-user storage isolation
- Background job queue (Celery/RQ)
- Upload size and page-count limits
- Logging / Monitoring
- Stronger filename/path validation
- Secure temporary file handling
- Protection against malicious files
- ~~Production WSGI server instead of Flask development server~~ **(Completed for Vercel deployment)**

## Technologies

- Python
- Flask
- HTML5
- CSS3
- JavaScript
- Pillow
- PyMuPDF
- python-pptx
- LibreOffice (for local PPT → PDF only)

## Future Improvements

- PDF compression
- PDF splitting
- PDF page extraction
- PDF page reordering
- PDF rotation
- PDF metadata removal
- PDF password protection
- WebP → JPG
- WebP → PNG
- Image resizing
- Image cropping
- Batch conversion
- Batch image compression
- OCR support
- Better editable PDF → PPT conversion
- Dark mode
- Progress indicators
- Automatic file cleanup
- Client-side processing
- Offline/PWA support

## Author

**Dipesh Acharya**  
GitHub: [https://github.com/Acharyadipesh](https://github.com/Acharyadipesh)

## Project Structure (Updated)
file converter/
│
├── api/
│ └── index.py # Main Flask application
├── static/
│ ├── style.css
│ └── app.js
├── templates/
│ └── index.html
├── requirements.txt
├── vercel.json # Vercel deployment configuration
└── README.md
