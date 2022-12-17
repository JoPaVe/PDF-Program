import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename, askdirectory
import os
import fitz
import tempfile
import io

from PyPDF2 import PdfFileReader

def merge_files(path_docs, tempor = False):

    merger = fitz.open()
    namelist = []
    for pdf in path_docs:
        namelist.append(pdf[0])
        sitetup = (int(pdf[1])-1,int(pdf[2]-1))
        pdffile = fitz.open(pdf[3])
        merger.insert_pdf(pdffile,from_page=sitetup[0],to_page=sitetup[1])
        pdffile.close()
    
    tk.Tk().withdraw()

    if tempor == False:
        path = askdirectory() #Replace by asksaveasfilename

        documentname = "_".join(namelist) + "_Zusammengefügt.pdf"
        try:
            output = os.path.join(path,documentname)
        except:
            output = os.path.join(path,"ZusammengefuegteDatei.pdf")
        merger.save(output)
    elif tempor == True:
        pdf_temp_file_path = os.path.join(tempfile.gettempdir(),"tempfile.pdf")
        merger.save(pdf_temp_file_path)
        return pdf_temp_file_path


    merger.close()

def pdf_convert(path, filename):
    doc = fitz.open()

    imgdoc = fitz.open(path)
    pdfbytes = imgdoc.convert_to_pdf()
    imgdoc.close()
    
    imgpdf = fitz.open("pdf",pdfbytes)
    doc.insert_pdf(imgpdf)

    path = asksaveasfilename(initialfile = filename, title="Bild als PDf-speichern",defaultextension=".pdf")
    doc.save(path)

    return (PdfFileReader(path, strict = False), path)