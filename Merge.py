from PyPDF2 import PdfFileMerger
import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename, askdirectory
from tkinter import messagebox
import os

def merge_files(path_docs):

    # Question boxes
    merger = PdfFileMerger()
    namelist = []
    for pdf in path_docs:
        namelist.append(pdf[0])
        sitetup = (int(pdf[1])-1,int(pdf[2]))
        dokinput = open(pdf[3], "rb")
        merger.append(fileobj=dokinput, pages=sitetup)
    
    tk.Tk().withdraw()
    path = askdirectory() #Replace by asksaveasfilename

    documentname = "_".join(namelist) + "_Zusammengefügt.pdf"
    output = open(os.path.join(path,documentname), "wb")
    
    merger.write(output)
    merger.close()
    output.close()


