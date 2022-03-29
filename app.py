import base64
from email.mime import image
from unicodedata import name
import requests
import json
from flask import Flask, jsonify, request
import os
import pythoncom
import glob
import re
from docx2pdf import convert
import win32com.client


app = Flask(__name__)


# Create a directory in a known location to save files to.
uploads_dir = os.path.join(app.instance_path, 'uploads')
os.makedirs(name=uploads_dir, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
# def dummy_api(name: str):
#     # print(name
#     data = open("base64.txt", "r").read()
#     decoded = base64.b64decode(data)
#     print decoded
#     return base64.b64encode(open("scan0081.pdf", "rb").read())




def dummy_api():
    from docx2pdf import convert

    body=request.data
    # EncodedFile = base64.urlsafe_b64encode(open("Document3.docx", "rb").read())
    EncodedFile=request.get_json()

    decoded = base64.urlsafe_b64decode(EncodedFile["name"])
    image_result = open(os.path.join(uploads_dir, "uploadedDocx.docx"), 'wb')
    image_result.write(decoded)
    image_result.close()
    # wdFormatPDF = 17

    # inputFile = os.path.join(uploads_dir, "uploadedDocx.docx")
    inputFile= os.path.join(uploads_dir, "uploadedDocx.docx")
    outputFile = os.path.join(uploads_dir, "helo899.pdf")

    print("hello")
    wdFormatPDF = 17
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch('Word.Application')

    word.Visible = True
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


    with open('json_data'+'.json', 'w', encoding="utf-8") as outfile:
        json.dump(str(base64.b64encode(open("C:\\Users\\user\\Downloads\\pythonAPIv2\\instance\\uploads\\helo899.pdf", "rb").read())), outfile,ensure_ascii=False)
    # pythoncom.CoInitialize()
    # try:
    #     if convert(inputFile,outputFile):
    #         print(f'File is success!')
    # except Exception as e:
    #     print(f'Error with \n{e}')
    
    # return os.path.join(app.instance_path)
    return "SUCCESS"

if __name__ == "__main__":
    app.run()
