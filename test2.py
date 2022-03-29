# import glob
# import os
# import re
# import comtypes.client
# fileslist=glob.glob(r"C:\Users\user\Downloads\pythonAPI\Document3.docx")
# regex_filename=r".*\\"
# regex_withoutext=r".docx"
# for i in range(0, len(fileslist)):
#     filename=re.sub(regex_filename, "", fileslist[i])
#     filename=re.sub(regex_withoutext, "", filename)
#     in_file = os.path.abspath(fileslist[i])

#     inputFile = os.path.abspath( "uploadedDocx.docx")
#     out_file = os.path.abspath("uyhF.pdf")
#     word = comtypes.client.CreateObject('Word.Application')
#     doc = word.Documents.Open(in_file)
#     doc.SaveAs(out_file, FileFormat=17)

#     doc.Close()

from docx2pdf import convert

convert("Document3.docx")