#from PyPDF2 import PdfFileMerger
import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename, askdirectory
import os
import fitz

def merge_files(path_docs):

    merger = fitz.open()
    namelist = []
    for pdf in path_docs:
        namelist.append(pdf[0])
        print(pdf)
        sitetup = (int(pdf[1])-1,int(pdf[2]-1))
        pdffile = fitz.open(pdf[3])
        merger.insert_pdf(pdffile,from_page=sitetup[0],to_page=sitetup[1])
        pdffile.close()
    
    tk.Tk().withdraw()
    path = askdirectory() #Replace by asksaveasfilename

    documentname = "_".join(namelist) + "_Zusammengefügt.pdf"
    try:
        output = os.path.join(path,documentname)
    except:
        output = os.path.join(path,"ZusammengefuegteDatei.pdf")
    merger.save(output)
    merger.close()
