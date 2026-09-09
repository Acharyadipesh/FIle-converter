import os
import uuid
import shutil
import zipfile
from pathlib import Path
import tempfile

from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    send_file
)

from werkzeug.utils import secure_filename

from PIL import Image

import fitz

from pptx import Presentation
from pptx.util import Inches


# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

BASE_DIR = Path(tempfile.gettempdir())
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


app = Flask(
    __name__,
    template_folder='../templates',
    static_folder='../static'
)

app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB total upload


# ============================================================
# ALLOWED FILE TYPES
# ============================================================

ALLOWED = {
    "pdf-to-ppt": {".pdf"},
    "jpg-to-png": {".jpg", ".jpeg"},
    "png-to-jpg": {".png"},
    "pdf-to-jpg": {".pdf"},
    "jpg-to-pdf": {".jpg", ".jpeg"},
    "png-to-pdf": {".png"},
    "merge-pdf": {".pdf"},
    "compress-image": {".jpg", ".jpeg", ".png", ".webp"}
}


# ============================================================
# HELPER
# ============================================================

def safe_filename(filename):
    filename = secure_filename(filename)
    if not filename:
        filename = "file"
    return filename


# ============================================================
# PDF → PPT
# ============================================================

def convert_pdf_to_ppt(pdf_file, output_file):
    pdf = fitz.open(str(pdf_file))
    if len(pdf) == 0:
        pdf.close()
        raise ValueError("The PDF contains no pages.")

    first_page = pdf[0].rect
    presentation = Presentation()
    presentation.slide_width = Inches(10)
    presentation.slide_height = Inches(
        10 * first_page.height / first_page.width
    )
    blank_layout = presentation.slide_layouts[6]

    for page in pdf:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
        temp_img = OUTPUT_DIR / f"temp_{uuid.uuid4().hex}.jpg"
        pixmap.save(str(temp_img))
        slide = presentation.slides.add_slide(blank_layout)
        slide.shapes.add_picture(
            str(temp_img),
            0, 0,
            width=presentation.slide_width,
            height=presentation.slide_height
        )
        temp_img.unlink(missing_ok=True)

    presentation.save(str(output_file))
    pdf.close()


# ============================================================
# IMAGE CONVERSION (single)
# ============================================================

def convert_image(input_file, output_file, output_format):
    with Image.open(input_file) as image:
        if output_format == "JPEG":
            if image.mode in ("RGBA", "LA", "P"):
                if image.mode == "P":
                    image = image.convert("RGBA")
                background = Image.new("RGB", image.size, "white")
                if "A" in image.getbands():
                    background.paste(image, mask=image.getchannel("A"))
                else:
                    background.paste(image)
                image = background
            else:
                image = image.convert("RGB")
        image.save(output_file, format=output_format, quality=95)


# ============================================================
# IMAGE COMPRESSION
# ============================================================

def compress_image(input_file, output_file, quality):
    with Image.open(input_file) as image:
        fmt = image.format
        if fmt in ("JPEG", "JPG"):
            if image.mode != "RGB":
                image = image.convert("RGB")
            image.save(output_file, format="JPEG", quality=quality, optimize=True, progressive=True)
        elif fmt == "PNG":
            if "A" in image.getbands():
                image.save(output_file, format="PNG", optimize=True, compress_level=9)
            else:
                image = image.convert("RGB")
                image.save(output_file, format="JPEG", quality=quality, optimize=True, progressive=True)
        elif fmt == "WEBP":
            image.save(output_file, format="WEBP", quality=quality, method=6)
        else:
            raise ValueError("Unsupported image format.")


# ============================================================
# PDF → JPG
# ============================================================

def convert_pdf_to_jpg(pdf_file, output_directory):
    pdf = fitz.open(str(pdf_file))
    pages = []
    for i, page in enumerate(pdf, start=1):
        pixmap = page.get_pixmap(matrix=fitz.Matrix(1.8, 1.8), alpha=False)
        out = output_directory / f"page-{i}.jpg"
        pixmap.save(str(out))
        pages.append(out)
    pdf.close()
    return pages


# ============================================================
# IMAGES → PDF 
# ============================================================

def convert_images_to_pdf(image_files, output_file, max_dimension=1500):
    """
    Convert a list of image files to a single PDF.
    - Rejects any image larger than 5 MB before processing.
    - Resizes images to max_dimension (default 1500px) to keep PDF size small.
    - Uses JPEG compression with quality 85 for the PDF.
    """
    images = []
    try:
        for f in image_files:
            # Check file size 
            file_size_mb = f.stat().st_size / (1024 * 1024)
            if file_size_mb > 5:
                raise ValueError(f"Image '{f.name}' is {file_size_mb:.1f} MB, which exceeds the 5 MB limit.")

            img = Image.open(f)

            # Resize if too large
            if max(img.size) > max_dimension:
                ratio = max_dimension / max(img.size)
                new_size = (int(img.size[0] * ratio), int(img.size[1] * ratio))
                img = img.resize(new_size, Image.Resampling.LANCZOS)

            # Convert to RGB (PDF doesn't support transparency)
            if img.mode != "RGB":
                img = img.convert("RGB")

            images.append(img)

        if not images:
            raise ValueError("No valid images provided.")

        # Save as PDF with compression
        images[0].save(
            output_file,
            save_all=True,
            append_images=images[1:],
            quality=85,
            optimize=True
        )
    finally:
        for img in images:
            img.close()


# ============================================================
# MERGE PDF
# ============================================================

def merge_pdf_files(pdf_files, output_file):
    merged = fitz.open()
    for f in pdf_files:
        pdf = fitz.open(str(f))
        merged.insert_pdf(pdf)
        pdf.close()
    merged.save(str(output_file))
    merged.close()


# ============================================================
# ROUTES
# ============================================================

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/api/convert", methods=["POST"])
def convert():
    task = request.form.get("task")
    uploaded_files = request.files.getlist("files")

    if task not in ALLOWED:
        return jsonify({"error": "Invalid conversion tool."}), 400

    if not uploaded_files or not uploaded_files[0].filename:
        return jsonify({"error": "Please select a file."}), 400

    job_id = uuid.uuid4().hex
    job_upload = UPLOAD_DIR / job_id
    job_output = OUTPUT_DIR / job_id
    job_upload.mkdir(parents=True, exist_ok=True)
    job_output.mkdir(parents=True, exist_ok=True)

    saved_files = []
    output_file = None

    try:
        # Save files (and check individual size for image-to-PDF tasks)
        for uploaded in uploaded_files:
            filename = safe_filename(uploaded.filename)
            ext = Path(filename).suffix.lower()
            if ext not in ALLOWED[task]:
                raise ValueError(f"File type {ext} not supported for this tool.")
            path = job_upload / filename
            uploaded.save(str(path))
            saved_files.append(path)

        # Perform conversion
        if task == "pdf-to-ppt":
            output_file = job_output / f"{saved_files[0].stem}.pptx"
            convert_pdf_to_ppt(saved_files[0], output_file)

        elif task == "jpg-to-png":
            output_file = job_output / f"{saved_files[0].stem}.png"
            convert_image(saved_files[0], output_file, "PNG")

        elif task == "png-to-jpg":
            output_file = job_output / f"{saved_files[0].stem}.jpg"
            convert_image(saved_files[0], output_file, "JPEG")

        elif task == "pdf-to-jpg":
            pages = convert_pdf_to_jpg(saved_files[0], job_output)
            if len(pages) == 1:
                output_file = pages[0]
            else:
                output_file = job_output / "converted-pages.zip"
                with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as archive:
                    for page in pages:
                        archive.write(page, page.name)

        elif task == "jpg-to-pdf":
            output_file = job_output / "converted-images.pdf"
            convert_images_to_pdf(saved_files, output_file)

        elif task == "png-to-pdf":
            output_file = job_output / "converted-images.pdf"
            convert_images_to_pdf(saved_files, output_file)

        elif task == "merge-pdf":
            if len(saved_files) < 2:
                raise ValueError("Select at least two PDF files to merge.")
            output_file = job_output / "merged.pdf"
            merge_pdf_files(saved_files, output_file)

        elif task == "compress-image":
            if len(saved_files) != 1:
                raise ValueError("Select exactly one image.")
            quality = int(request.form.get("quality", 70))
            quality = max(10, min(95, quality))
            source = saved_files[0]
            with Image.open(source) as img:
                has_alpha = "A" in img.getbands()
            ext = source.suffix.lower()
            if ext == ".png" and has_alpha:
                output_file = job_output / f"{source.stem}-compressed.png"
            elif ext == ".webp":
                output_file = job_output / f"{source.stem}-compressed.webp"
            else:
                output_file = job_output / f"{source.stem}-compressed.jpg"
            compress_image(source, output_file, quality)

        else:
            raise ValueError("Unknown conversion task.")

        # Verify output
        if not output_file or not output_file.exists():
            raise RuntimeError("Output file was not created.")

        original_size = saved_files[0].stat().st_size if saved_files else 0
        output_size = output_file.stat().st_size
        saved_percent = 0
        if original_size > 0:
            saved_percent = round((1 - (output_size / original_size)) * 100)

        return jsonify({
            "success": True,
            "filename": output_file.name,
            "url": f"/download/{job_id}/{output_file.name}",
            "original_size": original_size,
            "output_size": output_size,
            "saved_percent": saved_percent
        })

    except Exception as e:
        
        print("=" * 50)
        print("CONVERSION ERROR")
        print("Task:", task)
        print("Files:", [f.name for f in saved_files] if saved_files else "None")
        print("Error:", str(e))
        import traceback
        traceback.print_exc()
        print("=" * 50)

        return jsonify({"error": str(e)}), 500

    finally:
        shutil.rmtree(job_upload, ignore_errors=True)


@app.route("/download/<job_id>/<filename>")
def download(job_id, filename):
    filename = secure_filename(filename)
    file_path = OUTPUT_DIR / job_id / filename
    if not file_path.exists():
        return "File not found.", 404
    return send_file(
        str(file_path),
        as_attachment=True,
        download_name=file_path.name
    )


@app.errorhandler(413)
def too_large(error):
    return jsonify({"error": "File too large. Max total upload size is 100 MB."}), 413


if __name__ == "__main__":
    print("FileForge started at http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=True)
