import os
import win32com.client


def doConvert():

    wdFormatPDF = 17

    inputFile = os.path.abspath("Document3.docx")
    outputFile = os.path.abspath("GA_CIRCULAR_FINAL_DOCUMENT_DOCX.pdf")
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

doConvert()