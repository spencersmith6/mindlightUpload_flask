import os
from flask import Flask, request, session, g, redirect, url_for, abort, render_template, flash
import time
from collections import OrderedDict
from xlrd import open_workbook
from pymongo import MongoClient
import zipfile
from os import listdir
from os.path import isfile, join

app = Flask(__name__)


def doc_2_DB(dict, file_stamp):
    client = MongoClient()
    db = client['newDB']  # Database Name
    collection = db['newCol']  # Collection Name
    collection.insert_many(dict)


def compile_doc(file_stamp, clinic):
    f = open_workbook('EXCELfiles/%s.xlsx' % file_stamp)  # Recieves data from excel file
    sheet = f.sheets()[0]

###########
    feature_dict = edf_2_feature_dict(file_stamp) #dictionary containing features from all .edf file in zip file
###########

    myDict = []
    headers = sheet.row_values(0)
    for row in range(1, sheet.nrows):
        subjectData = OrderedDict()
        values = sheet.row_values(row)
        subjectData['Clinic Name'] = clinic
        for col in range(sheet.ncols):
            if headers[col] == 'Medical Condition':
                conditions = values[col].split(',')
                subjectData[headers[col]] = conditions
            else:
                subjectData[headers[col]] = values[col]

        subjectData['Features'] = feature_dict[values[0]] #Insert feature into Dict

        myDict.append(subjectData)
    return myDict

def is_zip(filename):
    if filename.rsplit('.', 1)[1] == "zip":
        return True
    else:
        return False

def is_excel(filename):
    if filename.rsplit('.', 1)[1] == "xlsx":
        return True
    else:
        return False


def archive(excel, zip, name):
    # Creates unique file name using the name of the clinic and an epoch time stamp
    # Saves excelfiles and zip files in local directories EXCELfiles and ZIPfiles respectively
    ts = time.time()
    stamp = '%s_%s'  % (name, str(int(round(ts))))
    excel_file_name = '%s.xlsx' % stamp
    zip_file_name = '%s.zip' % stamp
    excel.save(os.path.join('EXCELfiles', excel_file_name))
    zip.save(os.path.join('ZIPfiles', zip_file_name))
    return stamp

def dump_zip(file_stamp):
    zip_ref = zipfile.ZipFile('ZIPfiles/%s.zip' % file_stamp, 'r')
    os.makedirs("EEGfiles/%s/" % file_stamp)
    zip_ref.extractall("EEGfiles/%s/" % file_stamp)
    zip_ref.close()


def edf_2_feature_dict(file_stamp):
    edf_files = [f for f in listdir("EEGfiles/%s/" % file_stamp) if isfile(join("EEGfiles/%s/" % file_stamp, f))]
    feature_dict = OrderedDict()
    for file in edf_files:
        #  features[file] = extract_features(file) #### REPLACE WITH FEATURE EXTRACTION ALGO #######
        feature_dict[file] = [[1,2,3,4], [2,3,4,5], [3,4,5,6]]        #features
    return feature_dict


@app.route("/upload", methods=["POST", "GET"])
def upload():
    if request.method == "POST":
        if request.files['excel'].filename == '':
            return app.send_static_file('Missing_File_Error_Page.html')
        if request.files['zip'].filename == '':
            return app.send_static_file('Missing_File_Error_Page.html')
        if request.form['clinic'] == '':
            return app.send_static_file('Missing_File_Error_Page.html')
        clinic_name = request.form['clinic']
        excel_file = request.files['excel']
        zip_file = request.files['zip']
        if is_excel(excel_file.filename) and is_zip(zip_file.filename):
            file_stamp = archive(excel_file, zip_file, clinic_name)
            dump_zip(file_stamp)
            doc = compile_doc(file_stamp, clinic_name)
            doc_2_DB(doc, file_stamp)
            return "Data has been received. Thank you for participating."
        else:
            return app.send_static_file('Wrong_File_Error.html')
    return app.send_static_file('uploadPage.html')


@app.route("/thanks", methods=["POST"])
def thanks():
    return "Data has been recieved. Thank you for participating."


app.run()
