# from flask import Flask, render_template, request, send_file, after_this_request
# import tempfile
# import os
# from datetime import datetime
# import calendar
# from common import generate_schedule, coworkers

# app = Flask(__name__)

# @app.route('/')
# def index():
#     current_year = datetime.now().year
#     current_month = datetime.now().month
#     return render_template('index.html', current_year=current_year, current_month=current_month, month_name=calendar.month_name)

# @app.route('/generate', methods=['POST'])
# def generate():
#     year = int(request.form['year'])
#     month = int(request.form['month'])
#     wb = generate_schedule(year, month, coworkers)
#     temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
#     wb.save(temp_file.name)
    
#     @after_this_request
#     def remove_file(response):
#         os.remove(temp_file.name)
#         return response
    
#     return send_file(
#         temp_file.name,
#         as_attachment=True,
#         download_name=f"schedule_{year}_{month}.xlsx"
#     )

# if __name__ == '__main__':
#     app.run(debug=True)


from flask import Flask, render_template, request, send_file
import io
from datetime import datetime
import calendar
from common import generate_schedule, coworkers

app = Flask(__name__)

@app.route('/')
def index():
    current_year = datetime.now().year
    current_month = datetime.now().month
    return render_template('index.html', current_year=current_year, current_month=current_month, month_name=calendar.month_name)

@app.route('/generate', methods=['POST'])
def generate():
    year = int(request.form['year'])
    month = int(request.form['month'])
    wb = generate_schedule(year, month, coworkers)
    
    # Use BytesIO to store the workbook in memory
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)  # Move to the beginning of the BytesIO buffer
    
    # Send the file from memory
    return send_file(
        output,
        as_attachment=True,
        download_name=f"schedule_{year}_{month}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=True)