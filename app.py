from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from werkzeug.utils import secure_filename
from pdf2docx import Converter
import os
import logging
from pathlib import Path
import shutil
from PIL import Image
import img2pdf
import io
from docx2pdf import convert as docx2pdf_convert
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_bytes
import fitz  # PyMuPDF

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

CONVERSION_RULES = {
    'word2pdf': {'allowed': ['docx'], 'output': 'pdf'},
    'pdf2word': {'allowed': ['pdf'], 'output': 'docx'},
    'img2pdf': {'allowed': ['jpg', 'jpeg', 'png'], 'output': 'pdf'},
    'pdf2img': {'allowed': ['pdf'], 'output': 'zip'}
}


def allowed_file(filename, conversion_type):
    if '.' not in filename:
        return False
    ext = filename.rsplit('.', 1)[1].lower()
    return conversion_type in CONVERSION_RULES and ext in CONVERSION_RULES[conversion_type]['allowed']


def convert_pdf_to_images(pdf_path, output_dir):
    """Convert PDF to images using PyMuPDF (fitz)"""
    try:
        # Open the PDF
        pdf_document = fitz.open(pdf_path)

        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)

        # Iterate through pages
        for page_number in range(pdf_document.page_count):
            # Get the page
            page = pdf_document[page_number]

            # Convert page to image
            pix = page.get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72))  # 300 DPI

            # Save the image
            image_path = os.path.join(output_dir, f'page_{page_number + 1}.png')
            pix.save(image_path)

        pdf_document.close()
        return True
    except Exception as e:
        logger.error(f"Error in PDF to Image conversion: {str(e)}")
        return False


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['GET', 'POST'])
def convert_file():
    if request.method == 'GET':
        return redirect(url_for('index'))

    if 'file' not in request.files:
        flash('No file selected')
        return redirect(url_for('index'))

    file = request.files['file']
    conversion_type = request.form.get('conversion_type')

    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))

    if not allowed_file(file.filename, conversion_type):
        flash(f'Invalid file type for {conversion_type} conversion')
        return redirect(url_for('index'))

    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        output_filename = os.path.splitext(filename)[0]

        try:
            if conversion_type == 'word2pdf':
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{output_filename}.pdf')
                try:
                    # First attempt with docx2pdf
                    docx2pdf_convert(filepath, output_path)
                except Exception as e:
                    logger.error(f"docx2pdf conversion failed: {str(e)}")
                    # Fallback to alternative method if needed
                    flash('Error in conversion, please try again')
                    return redirect(url_for('index'))

            elif conversion_type == 'pdf2word':
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{output_filename}.docx')
                cv = Converter(filepath)
                cv.convert(output_path)
                cv.close()

            elif conversion_type == 'img2pdf':
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{output_filename}.pdf')
                image = Image.open(filepath)
                if image.mode == 'RGBA':
                    image = image.convert('RGB')
                with open(output_path, "wb") as f:
                    f.write(img2pdf.convert(filepath))

            elif conversion_type == 'pdf2img':
                temp_dir = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                if convert_pdf_to_images(filepath, temp_dir):
                    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{output_filename}.zip')
                    shutil.make_archive(output_path[:-4], 'zip', temp_dir)
                    shutil.rmtree(temp_dir)
                else:
                    flash('Error converting PDF to images')
                    return redirect(url_for('index'))

            return send_file(
                output_path,
                as_attachment=True,
                download_name=os.path.basename(output_path)
            )

        except Exception as e:
            logger.error(f"Conversion error: {str(e)}")
            flash(f'Error during conversion: {str(e)}')
            return redirect(url_for('index'))

        finally:
            # Clean up
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
                if 'output_path' in locals() and os.path.exists(output_path):
                    os.remove(output_path)
                if 'temp_dir' in locals() and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
            except Exception as e:
                logger.error(f"Cleanup error: {str(e)}")

    except Exception as e:
        logger.error(f"General error: {str(e)}")
        flash('An unexpected error occurred')
        return redirect(url_for('index'))


if __name__ == '__main__':
    logger.info("Starting File Converter application...")
    app.run(debug=True)
