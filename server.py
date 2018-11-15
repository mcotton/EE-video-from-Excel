from flask import Flask
from flask import render_template
from openpyxl import Workbook
from openpyxl import load_workbook

app = Flask(__name__)

@app.route("/")
def index():
    
    wb = load_workbook("Example.xlsx")
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

    return render_template('index.html', template_values=ret_obj)