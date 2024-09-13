import os
import shutil

from flask import Flask, render_template, request, redirect, send_from_directory, url_for

from report import report_creator

app = Flask(__name__)

app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024 #CONFIGURING THE FILE SIZE

UPLOAD_FOLDER = 'uploads'
GENERATED_FOLDER = 'generated'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

# Create the uploads and generated directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['GENERATED_FOLDER'] = GENERATED_FOLDER


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file and allowed_file(file.filename):
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))
            return redirect(url_for('index'))
    return redirect(url_for('index'))


@app.route('/generate', methods=['POST'])
def generate_report():
    if request.method == 'POST':
        # Process the uploaded files and generate the RTGS report
        report_filename = report_creator()
        if report_filename:
            # Move the generated report to the 'generated' folder
            os.replace(report_filename, os.path.join(app.config['GENERATED_FOLDER'], report_filename))
            return send_from_directory(app.config['GENERATED_FOLDER'], report_filename, as_attachment=True)
    return redirect(url_for('index'))


@app.route('/clear', methods=['POST'])
def clear_folders():
    try:
        shutil.rmtree(app.config['UPLOAD_FOLDER'])
        shutil.rmtree(app.config['GENERATED_FOLDER'])
    except OSError as e:
        print(f"Error: {e.filename} - {e.strerror}")
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)
    return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)
