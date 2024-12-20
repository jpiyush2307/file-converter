# app.py
from flask import Flask, render_template, request, send_file, redirect, url_for
from werkzeug.utils import secure_filename
from docx2pdf import convert
from pdf2docx import Converter
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'docx', 'pdf'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return redirect(request.url)

    file = request.files['file']
    conversion_type = request.form.get('conversion_type')

    if file.filename == '':
        return redirect(request.url)

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        output_filename = os.path.splitext(filename)[0]

        if conversion_type == 'word2pdf':
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{output_filename}.pdf')
            convert(filepath, output_path)
        elif conversion_type == 'pdf2word':
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{output_filename}.docx')
            cv = Converter(filepath)
            cv.convert(output_path)
            cv.close()

        return send_file(output_path, as_attachment=True)

    return redirect(request.url)


if __name__ == '__main__':
    app.run(debug=True)
