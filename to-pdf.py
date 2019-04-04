import os
import shutil
import requests
import tempfile

from gevent.pywsgi import WSGIServer
from flask import Flask, after_this_request, render_template, request, send_file
from subprocess import call
from PyPDF2 import PdfFileMerger, PdfFileReader

UPLOAD_FOLDER = '/tmp'
ALLOWED_EXTENSIONS = set(['doc', 'docx', 'xls', 'xlsx'])

app = Flask(__name__)


# Convert using Libre Office
def convert_file(output_dir, input_file):
    call('libreoffice --headless --convert-to pdf --outdir %s %s ' %
         (output_dir, input_file), shell=True)
    return add_cover(input_file)
    # return '.pdf'


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def api():
    work_dir = tempfile.TemporaryDirectory()
    file_name = 'document'
    input_file_path = os.path.join(work_dir.name, file_name)

    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            return 'No file provided'
        file = request.files['file']
        if file.filename == '':
            return 'No file provided'
        if file and allowed_file(file.filename):
            file.save(input_file_path)

    if request.method == 'GET':
        url = request.args.get('url', type=str)
        if not url:
            return render_template('index.html')
        # Download from URL
        response = requests.get(url, stream=True)
        with open(input_file_path, 'wb') as file:
            shutil.copyfileobj(response.raw, file)
        del response

    extension = convert_file(work_dir.name, input_file_path)
    output_file_path = os.path.join(work_dir.name, file_name + extension)

    @after_this_request
    def cleanup(response):
        work_dir.cleanup()
        return response

    return send_file(output_file_path, mimetype='application/pdf')


def add_cover(input_path):
    extension = '.cover.pdf'
    output_file_path = input_path + extension

    # add cover sheet pdf page
    cover_file = './cover.pdf'
    merger = PdfFileMerger()
    merger.append(PdfFileReader(open(cover_file, 'rb')))
    merger.append(PdfFileReader(open(input_path + '.pdf', 'rb')))
    merger.write(output_file_path)

    return extension


if __name__ == "__main__":
    http_server = WSGIServer(('', int(os.environ.get('PORT', 8080))), app)
    http_server.serve_forever()
