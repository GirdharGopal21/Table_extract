from apply_password import *
from flask import Flask, request, render_template, flash, redirect, send_file
from werkzeug.utils import secure_filename
import shutil
import pandas as pd
from table_extract import auto_table_extract
import webbrowser
from pathlib import Path
import os

# Set up paths
cwd = Path.cwd()
ate = str(cwd)
upload_folder = os.path.join(ate, "upload")
excel_folder = os.path.join(ate, "excel")
csv_folder = os.path.join(ate, "csv")

# Ensure directories exist
for folder in [upload_folder, excel_folder, csv_folder]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# Flask app configuration
app = Flask(__name__)
app.secret_key = "secret key"
app.config['UPLOAD_FOLDER'] = upload_folder
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
ALLOWED_EXTENSIONS = {'pdf'}

# Global Variables
flag = session = export = alpha = beta = count = 0
a = ""
processed_text = "1234"

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    flash("Upload a File!")
    return render_template('start.html')

@app.route('/uploads')
def upload_form():
    return render_template('Index.html')

@app.route('/how_to_use')
def how_to_use():
    return render_template("use.html")

@app.route('/uploads', methods=['POST'])
def upload_file():
    global a, flag, session, export, count
    flag = session = 1
    export = count = 0
    file = request.files.get('file')
    
    if not file or file.filename == '':
        flash('Please Upload a File!')
        return redirect("/uploads")
    
    if allowed_file(file.filename):
        a = secure_filename(file.filename)
        file.save(os.path.join(upload_folder, a))
        flash('File Successfully Uploaded!')
    else:
        flash('ONLY PDF FILES ARE ALLOWED!')
    return redirect('/uploads')

@app.route('/pdf_view/')
def pdf_view():
    global session
    if session == 1:
        try:
            webbrowser.open(os.path.join(upload_folder, a))
        except:
            flash("Please Upload a File to view!")
    else:
        flash("New Session! Please Upload a File to view!")
    return redirect("/uploads")

@app.route('/download/')
def download():
    global flag, session, export, alpha, beta
    if session == 1 or count == 1:
        if flag == 1 or alpha == 1:
            session = flag = 0
            export = 1 - export
            alpha, beta = 0, 1
            try:
                filename = a.replace(".pdf", ".xlsx")
                path = os.path.join(excel_folder, "output.xlsx")
                return send_file(path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                                 download_name=filename, as_attachment=True)
            except:
                flash("Please Extract the tables")
        else:
            flash("Please Upload a file or Extract the tables")
    else:
        flash("New Session! Please Upload a File")
    return redirect("/uploads")

@app.route('/extract_table/')
def table_extraction():
    global flag, session
    if session == 1:
        if not os.listdir(upload_folder):
            flash("Please Upload a File!")
        else:
            try:
                auto_table_extract(os.path.join(upload_folder, a))
                flash("Your Excel File is Ready!")
                flag = 1
            except:
                flash("Extraction Failed! Ensure the file is valid.")
    else:
        flash("New Session! Please Upload a File!")
    return redirect("/uploads")

@app.route('/preview/<filename>')
def data_frame_show(filename):
    path = os.path.join(csv_folder, filename)
    try:
        df = pd.read_csv(path, encoding='latin1')
        return df.to_html(header=True, table_id="table")
    except:
        flash("Error loading the file")
        return redirect("/uploads")

@app.route('/password_apply', methods=['POST'])
def password_apply():
    global processed_text
    if not os.listdir(csv_folder):
        flash("Please Extract Tables!")
    else:
        processed_text = request.form.get('CONFIRM_PASSWORD', processed_text)
    return render_template("password.html")

@app.after_request
def add_header(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

if __name__ == "__main__":
    webbrowser.open("http://127.0.0.1:5000/")
    app.run(threaded=True, debug=True)
