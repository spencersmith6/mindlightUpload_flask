import os
import sqlite3
from flask import Flask, request, session, g, redirect, url_for, abort, render_template, flash
from wtforms import Form, StringField, TextAreaField, validators, StringField, SubmitField
import time
app = Flask(__name__)

def is_zip(filename):
    if filename.rsplit('.', 1)[1] =="zip":
        return True
    else:
        return False

def is_excel(filename):
    if filename.rsplit('.', 1)[1] =="xlsx":
        return True
    else:
        return False

def archive(excel, zip, name):
    ts = time.time()
    excel_file_name = name + '_' +str(int(round(ts))) + '.xlsx'
    zip_files_name = name + '_' +str(int(round(ts))) + '.zip'
    excel.save(os.path.join('EXCELfiles', excel_file_name))
    zip.save(os.path.join('ZIPfiles', zip_files_name))


@app.route("/upload", methods =["POST", "GET"])
def upload():

    if request.method == "POST":
        if 'excel' not in request.files:
            print request.files
            return app.send_static_file('errorUploadPage.html')
        if 'zip' not in request.files:
            print "fail#1"
            return app.send_static_file('errorUploadPage.html')


        clinic_name = request.form['clinic']
        excel_file = request.files['excel']
        zip_file = request.files['zip']

        if is_excel(excel_file.filename) and is_zip(zip_file.filename):

            archive(excel_file, zip_file, clinic_name)
            return "Data has been recieved. Thank you for participating."

    return app.send_static_file('uploadPage.html')

@app.route("/thanks", methods = ["POST"])
def thanks():
    return "Data has been recieved. Thank you for participating."


app.run()