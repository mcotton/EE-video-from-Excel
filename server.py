from flask import Flask
from flask import render_template
from flask import request, redirect, abort, url_for

from werkzeug.utils import secure_filename

from openpyxl import Workbook
from openpyxl import load_workbook

app = Flask(__name__)

@app.route("/", methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)

        f = request.files['file']
        location = secure_filename(f.filename)
        f.save('./uploads/' + location)
        return redirect('/view/%s' % location, code=302)

@app.route('/view/<path:filename>', methods=['GET'])
def view(filename):
    wb = load_workbook('./uploads/' + filename)
    ret_obj = {}

    if wb:
        ws = wb["Events"]

        for row in ws.iter_rows(row_offset=2):
            esn = row[0].value
            start = row[1].value
            end = row[2].value

            if esn:
                if esn in ret_obj:
                    ret_obj[esn].append(
                        {
                            'start_timestamp': start,
                            'end_timestamp': end
                        }   
                    )
                else:
                    ret_obj[esn] = [{
                        'start_timestamp': start,
                        'end_timestamp': end
                    }]

    return render_template('upload.html', template_values=ret_obj)
