from pymongo import MongoClient
import simplejson as json
from xlrd import open_workbook
from collections import OrderedDict


def excel_2_dict(file_name):
    f = open_workbook('EXCELfiles/%s' % file_name)
    sheet = f.sheets()[0]
    myDict = []
    headers = sheet.row_values(0)
    for row in range(1, sheet.nrows):
        subjectData = OrderedDict()
        values = sheet.row_values(row)
        for col in range(sheet.ncols):
            if headers[col] == 'Medical Condition':
                conditions = values[col].split(',')
                subjectData[headers[col]] = conditions
            else:
                subjectData[headers[col]] = values[col]
        myDict.append(subjectData)
    return myDict

print excel_2_dict('Harvard_1478294513.xlsx')

doc = excel_2_dict('Harvard_1478294513.xlsx')

client = MongoClient()
db = client['newDB']
collection =db['newCol']
collection.insert_many(doc)