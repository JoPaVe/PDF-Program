import tkinter as tk
from unicodedata import decimal
import PyPDF2
from PyPDF2 import PdfFileReader, PdfReader, PdfWriter
from PIL import Image, ImageTk
from pygments import highlight
from Merge import merge_files, pdf_convert
from Extract import extract_files
import os
import re
import tempfile

from tkinter.filedialog import askopenfile, askopenfilename, askopenfilenames, askdirectory
from tkinter import RAISED, Frame, scrolledtext, Label
from tkinter import LEFT, RIDGE, ttk
from tkinter.messagebox import showinfo
from tkinterdnd2 import * 
from PIL import ImageTk, Image  

from pdf2image import convert_from_path, convert_from_bytes


class tkinterApp(TkinterDnD.Tk):
    def __init__(self, *args,**kwargs):
        TkinterDnD.Tk.__init__(self,*args,**kwargs)
        
        container = tk.Frame(self) 
        container.pack(side = "top", fill = "both", expand = True)
  
        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)
  
        self.frames = {} 
        self.title('PDF-Tool')
        self.geometry('650x400')
        self.columnconfigure(0, weight=3)
  
        for F in (StartPage, MergePage, ExtractPage):
  
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row = 0, column = 0, sticky ="nsew")
            frame.config(bg="#FFD580")
            frame.columnconfigure((0,2), weight=1)
            frame.columnconfigure(1,weight=3)
            frame.rowconfigure(0,weight=1)
            frame.rowconfigure((1,2,3),weight=2)
  
        self.show_frame(StartPage)

    def show_frame(self, cont):
        if cont == ExtractPage:
            self.geometry('650x500')
        else:
            self.geometry('650x400')
        frame = self.frames[cont]
        frame.tkraise()

class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)   

### Titel / Überschrift
        # Titleframe
        instruct_frame = tk.Frame(self, bg="#FFA500")
        instruct_frame.columnconfigure(0,weight=1)
        instruct_frame.grid(row = 0, column = 0)

        # Label als Überschriften
        #> PDF-Tool
        label_pdf_title = tk.Label(instruct_frame, text = "PDF-Programm von Jonas Verch", font = "Helvetica 10 bold", bg="#FFA500", anchor=tk.CENTER, width=100, height=1)
        label_pdf_title.grid(column=0,row=0, padx=(10), pady=10)

### Instructions
        #> Instruction Frame
        instruct_frame = tk.Frame(self,bg="#FFD580")
        instruct_frame.columnconfigure(0,weight=1)
        instruct_frame.grid(row = 1, column = 0)

        #> Controller Button
        upper_btn = tk.Button(instruct_frame, text ="PDF zusammenfassen", command = lambda : controller.show_frame(MergePage), height=1, width = 30, pady=20, anchor=tk.CENTER, relief="groove", bg="#dadada")
        upper_btn.grid(row=0,column=0)

        lower_btn = tk.Button(instruct_frame, text ="PDF-Kommentare extrahieren", command = lambda : controller.show_frame(ExtractPage), height=1, width = 30, pady=20, anchor=tk.CENTER, relief="groove", bg="#dadada")
        lower_btn.grid(row=0,column=1)

        #> Vergrößern der Bilder beim Klick
        def enlarge_descr(image,parent):
            working_path = os.getcwd()
            image = Image.open(os.path.join(working_path,image))
            
            # Geometry von Neuem Windows
            descr_top = tk.Toplevel(parent)
            descr_top.geometry(str(int(image.size[0]*(2/3)))+'x'+str(int(image.size[1]*(2/3))))

            # Frame von neuem Window
            descr_top_frame = tk.Frame(descr_top, width=(int(image.size[0]*(2/3))), height=(int(image.size[1]*(2/3))))
            descr_top_frame.columnconfigure(0,weight=1)
            descr_top_frame.rowconfigure(0,weight=1)
            descr_top_frame.grid(column=0,row=0,sticky="nsew")
            descr_top_frame.place(anchor='center', relx=0.5, rely=0.5)

            # Picture in neuem Frame
            img_resized = image.resize((int(image.size[0]*(2/3)),int(image.size[1]*(2/3))), Image.Resampling.LANCZOS)
            img_instr = ImageTk.PhotoImage(img_resized)

            label = Label(descr_top_frame,image=img_instr)
            label.image= img_instr
            label.grid(column=0,row=0)


        #> Bilder für beide Vorgänge
        working_path = os.getcwd()
        #> Merge_Instructions
        img_instr_merge = Image.open(os.path.join(working_path,"Merge_documents_description.png"))
        #> Resizes Picture
        img_instr_merge = img_instr_merge.resize((250,200), Image.Resampling.LANCZOS)

        img_instr_merge = ImageTk.PhotoImage(img_instr_merge)
        btn_instr_merge = tk.Button(instruct_frame,image=img_instr_merge,command=lambda: enlarge_descr("Merge_documents_description.png",controller))
        btn_instr_merge.image = img_instr_merge
        btn_instr_merge.grid(row=1,column=0, rowspan=2, padx=10)
        
        # Extract_Instructions_Picture
        img_instr_extract = Image.open(os.path.join(working_path,"Comment_extraction_description.png"))
        #> Resize Picture
        img_instr_extract = img_instr_extract.resize((250,200), Image.Resampling.LANCZOS)
        
        img_instr_extract = ImageTk.PhotoImage(img_instr_extract)
        btn_instr_extract = tk.Button(instruct_frame,image=img_instr_extract,command= lambda: enlarge_descr("Comment_extraction_description.png",controller))
        btn_instr_extract.image = img_instr_extract
        btn_instr_extract.grid(row=1,column=1, rowspan=2, padx=10)


class MergePage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        ## Initiierungsvariablen
### Layout
        # Controller
        #> Controller-Frame
        control_frame = tk.Frame(self, bg="#FFA500", width=600)
        control_frame.columnconfigure(0,weight=1)
        control_frame.grid(row = 0, column = 0, columnspan=3)

        #> Controller-Überschrift
        mergeinstructions = tk.Label(control_frame, text = "PDF zusammenfassen", font = "Raleway 10 bold", bg="#FFA500")
        mergeinstructions.grid(column=1,row=0, padx=30, pady=10)
        
        #> Controller Button
        upper_btn = tk.Button(control_frame, text ="Startseite", command = lambda : controller.show_frame(StartPage), height=1, width = 20,relief="groove", bg="#dadada")
        upper_btn.grid(row=0,column=0, padx=30)

        lower_btn = tk.Button(control_frame, text ="PDF Kommentare extrahieren", command = lambda : controller.show_frame(ExtractPage), height=1, width = 20, relief="groove", bg="#dadada")
        lower_btn.grid(row=0,column=2, padx=30, ipadx=30)

        # Input Frame
        input_frame = tk.Frame(self, bg="#FFD580")
        input_frame.grid(row = 2, column = 0, columnspan=3)

        # Button Frame
        button_frame = tk.Frame(self, bg="#FFD580")
        button_frame.grid(row = 1, column = 0, columnspan=3)
        
        # Tree Frame
        tree_frame = tk.Frame(self, bg="#FFD580")
        tree_frame.grid(row = 3, column = 0, columnspan=3)        
        
        #> Tabelleninputvariable
        table_fill = [] 

        #> Picture_Viewer_Startvalue
        self.cur_page = 0

        
### Funktionen Button_Frame
        # Ausführen
        def call_merge():
            list_entries = tree.get_children()
            paths = []

            if len(list_entries) == 0:
                showinfo("Uffbasse", "Es muss mindestens ein Dokument hinzugefügt werden! ")
            else:
                # Fill paths mit Name, Von, Bis, Path
                for entry in list_entries:
                    paths.append([tree.item(entry)['values'][0],tree.item(entry)['values'][1],tree.item(entry)['values'][2],tree.item(entry)['values'][3]])
                
                merge_files(paths)

        # Vorschau
        def call_preview():
            list_entries = tree.get_children()
            paths = []
            
            if len(list_entries) == 0:
                showinfo("Uffbasse", "Es muss mindestens ein Dokument hinzugefügt werden! ")
            else:
                for entry in list_entries:
                    paths.append([tree.item(entry)['values'][0],tree.item(entry)['values'][1],tree.item(entry)['values'][2],tree.item(entry)['values'][3]])
                
                self.pdf_temp = merge_files(paths, tempor = True)

                images_pdf = convert_from_path(self.pdf_temp, poppler_path = r"C:\Anaconda\pkgs\poppler-22.11.0-ha6c1112_0\Library\bin") 
                photo_viewer(images_pdf)
                
        
        def button_press(index):
            cur_col = self.buttons[index][0].cget('bg')
            if cur_col == "green":
                self.buttons[index][0].configure(bg = "black")
                self.buttons[index][2] = False

            else:
                self.buttons[index][0].configure(bg = "green")
                self.buttons[index][2] = True

        def execute_preview():
            reader = PdfReader(self.pdf_temp)
            writer = PdfWriter()
            
            sorted_buttons = sorted(self.buttons.items(), key = lambda x: x[1][2])

            for page, button_item in enumerate(sorted_buttons):
                if button_item[1][2] == True:
                    writer.add_page(reader.pages[button_item[0]])
            
            path = askdirectory()
            output_path = os.path.join(path, "PDF_zusammengefügt.pdf")
            try: 
                with open(output_path, "wb") as fp:
                    writer.write(fp)
            except:
                showinfo("Uffbasse", "Das Dokument ist beschädigt und kann nicht geladen werden! ")
            
            os.remove(self.pdf_temp)
 
        def photo_viewer(image_list):
            image_top = tk.Toplevel()
            image_top.title('Vorschau')
            image_top.resizable()
            image_top.grid_columnconfigure(0, weight=1)
            image_top.grid_columnconfigure(1, weight=1)
            
            self.top_main_frame = tk.Frame(image_top, bg="#FFA500", width=600)
            self.top_main_frame.grid(row = 0, column = 0,sticky=(tk.N, tk.S, tk.E, tk.W))
            
            # All buttons with pictures
            self.buttons = {}          

            for button_count, page in enumerate(image_list):
                page_picture = ImageTk.PhotoImage(page.resize((158,224), Image.Resampling.LANCZOS))
                self.buttons[button_count] = [tk.Button(self.top_main_frame, width = 158, height=224,  command = lambda but = button_count: button_press(but), image = page_picture, bg="green"), button_count, True] # Dict with keys: Initial count in temp_pdf, values: (buttonlink, currentpage, status)
                self.buttons[button_count][0].image = page_picture
                        
            
            show_picture() #initialize first page buttons
            print(self.buttons)
            # Button Back
            self.back_button = tk.Button(self.top_main_frame, text = 'Back', command = lambda: Back(), height=1, width = 20, relief="groove", bg="#dadada")
            self.back_button.grid(row = 3, column = 0, columnspan = 3)

            # Button Next
            self.next_button = tk.Button(self.top_main_frame, text = 'Next', command = lambda: Next(), height=1, width = 20, relief="groove", bg="#dadada")  
            self.next_button.grid(row = 3, column = 2, columnspan = 3)
            
            self.execute_preview_button = tk.Button(self.top_main_frame, text = 'Ausführen', command = lambda: execute_preview(), height=1, width = 20, relief="groove", bg="#dadada")  
            self.execute_preview_button.grid(row = 3, column = 1, columnspan = 3)
            
        
        def show_picture():
            row_count = 0
            column_count = 0
        
            
            for button in list(self.buttons.values())[self.cur_page:]:
                if row_count >= 2:
                    break
                else:
                    button[0].grid(row = row_count, column = column_count)
                    column_count += 1
                    if column_count == 5:
                        row_count += 1
                        column_count = 0
                        
        
        def Back():  
            for widget in self.top_main_frame.winfo_children():
                widget.grid_remove()
                self.back_button.grid(row = 3, column = 0, columnspan = 3)
                self.next_button.grid(row = 3, column = 2, columnspan = 3)
                self.execute_preview_button.grid(row = 3, column = 1, columnspan = 3)
            
            self.cur_page = self.cur_page - 10  
            if self.cur_page >= 0:
                show_picture()
            else:
                self.cur_page = 0
                show_picture()


        def Next():
            for widget in self.top_main_frame.winfo_children():
                widget.grid_remove()
                self.back_button.grid(row = 3, column = 0, columnspan = 3)
                self.next_button.grid(row = 3, column = 2, columnspan = 3)
                self.execute_preview_button.grid(row = 3, column = 1, columnspan = 3)
                
            self.cur_page = self.cur_page + 10 
            if self.cur_page >= len(self.buttons):
                self.cur_page = len(self.buttons) - 10
                show_picture()
            else:
                show_picture()


        # Hinzufügen
        def enter_doc(eventdrop = None):
            docpath_r = []
            if eventdrop != None: #Eventdrop, wenn per Drag and Drop ausgewählt wird.
                find_colon = [m.start() - 1 for m in re.finditer(":", eventdrop.data)] # Finde alle Doppelpunkte, schneide String bis ".pdf" or ".jpeg, .jpg, .png" aus.
                find_ext = [m.start() + 4 if m.group(0) != ".jpeg" else m.start() + 5 for m in re.finditer(".pdf|.jpeg|.jpg|.png", eventdrop.data)] #control for .jpeg with m.group(0) which gives the found extension
 
                paths_ranges = list(zip(find_colon,find_ext))
                for pathrange in paths_ranges:
                    docpath_r.append(eventdrop.data[pathrange[0]:pathrange[1]])

            else:
                docpath = [askopenfilenames(parent=self, title="Bitte gebe den Pfad des Dokuments ein: ", filetypes=[("Pdf or picture file", ("*.pdf", "*.png","*.jpeg","*.jpg"))])]
                docpath_r = [path for path in docpath[0]]
                if docpath[0] == '':
                    return #Wenn kein Dokument ausgewählt wird

            for path in docpath_r:    
                filename = path.split("/")[-1].split(".")[0]

                try: reader = PyPDF2.PdfFileReader(path, strict = False)
                except: 
                    reader, path = pdf_convert(path, filename) # Use pdf_convert from merge.py if document is jpg,jpeg,png
                count_pages = reader.getNumPages()

                if count_pages == 0:
                    showinfo("Uffbasse", "Das Dokument ist beschädigt und kann nicht geladen werden! ")
                else:
                    fill_table(table_fill,(filename,'1',count_pages,path))
                    table_fill.pop()

        # Löschen
        def delete_doc():
            for selected_item in tree.selection():
                child_id = tree.prev(selected_item)
                tree.delete(selected_item)
                if child_id == "": # Falls kein vorheriges Item mehr in Treeview ist
                    try:  # Falls kein erstes Item mehr vorhanden ist
                        child_id = tree.get_children()[0]        
                    except:
                        child_id = ""
                try: # Falls keine Items mehr in Treeview sind
                    tree.focus(child_id)
                    tree.selection_set(child_id)
                except: return
        # Verändere Dokumentenseiten
        #> Ausgewählte Input
        def select_records(event):
            # Clear entry boxes
            from_box.delete(0, 'end')
            to_box.delete(0, 'end')

            # Grab record number
            selected = tree.focus()
            # Grab record values
            values = tree.item(selected, 'values')

            try:
                from_box.insert(0,values[1])
                to_box.insert(0,values[2])
            except:
                pass

        #> Aktualisiere Dokumentenseiten
        def update_docnum(event):
            # Get selected
            selected = tree.focus()
            if selected == '':
                showinfo("Uffbasse", "Es muss zuerst die Seiten des Dokuments geladen werden!")
                return
 
            values = tree.item(selected, 'values')
            # Save new data
            min_page = from_box.get()
            max_page = to_box.get()

            if int(max_page) > int(values[2]):
                showinfo("Uffbasse", "Die höchste Seitenanzahl darf nicht höher als die Seitenanzahl der PDF sein! ")
            elif int(min_page) < 0:
                showinfo("Uffbasse", "Die niedrigste Seitenzahl darf nicht kleiner als 0 sein! ")
            elif int(min_page) == 0 & int(max_page) == 0:
                showinfo("Uffbasse", "Es muss mindest eine Seitenzahl ausgewählt werden, oder der Eintrag gelöscht werden! ")
            elif int(min_page) > int(max_page):
                showinfo("Uffbasse", "Die höchste Seite darf nicht kleiner sein als die kleinste Seite! ")
            else:
                tree.item(selected, values=(values[0],from_box.get(),to_box.get(), values[3]))
            
            #Clear Input boxes
            from_box.delete(0, 'end')
            to_box.delete(0, 'end')

        def copy_doc():
            # Get selected
            selected = tree.focus()
            if selected == '':
                showinfo("Uffbasse", "Es muss zuerst die Seiten des Dokuments geladen werden!")
                return            

            values = tree.item(selected, 'values')
            path = values[3] #Get docpath for copy

            reader = PyPDF2.PdfFileReader(path, strict = False)
            count_pages = reader.getNumPages() #load original docpages
            
            if count_pages == 0:
                showinfo("Uffbasse", "Das Dokument ist beschädigt und kann nicht geladen werden! ")
            else:
                values = list(values)
                values[2] = count_pages #update count_pages back to original document, change to list to immute
                
                tree.insert(parent='',index = 0, values = tuple(values))

### Input_Frame
        # Entries
        #> Von Box für Seitenangabe
        text_from_box = tk.StringVar()
        from_box = ttk.Entry(input_frame, textvariable=text_from_box)
        from_box.grid(row=0,column=2)

        #> Bis Box für Seitenangabe
        text_to_box = tk.StringVar()
        to_box = ttk.Entry(input_frame, textvariable=text_to_box)
        to_box.grid(row=1,column=2)

        # Update für Seitenangabe 
        #update_site = tk.StringVar()
        #update_site.set("Aktualisieren")
        #update_btn = tk.Button(input_frame, textvariable=update_site, command=lambda:update_docnum(), height=1, width=10, relief="groove", bg="#dadada")
        #update_btn.grid(column=2,row=3, padx=65,ipadx=20)

        # Copy von Dokumenten 
        copy_site = tk.StringVar()
        copy_site.set("Kopieren")
        copy_btn = tk.Button(input_frame, textvariable=copy_site, command=lambda:copy_doc(), height=1, width=10, relief="groove", bg="#dadada")
        copy_btn.grid(column=2,row=4, padx=65,ipadx=20)

        # Laden für Seitenangabe
        #load_site = tk.StringVar()
        #load_site.set("Seiten laden")
        #load_site_btn = tk.Button(input_frame, textvariable=load_site, command=lambda:select_records(), height=1, width=10, relief="groove", bg="#dadada")
        #load_site_btn.grid(column=0,row=0, padx=65,ipadx=20)
        
        ## Label
        #> Von für Seitenangabe
        von_label = tk.Label(input_frame, text="Von: ", bg="#FFD580")
        von_label.grid(column=1,row = 0)
        

        #> Bis für Seitenangabe
        bis_label = tk.Label(input_frame,text="Bis: ",bg="#FFD580")
        bis_label.grid(column=1,row = 1)

### Treeview_Frame
        # Tabelle füllen
        def fill_table(table_list, insert_tuple):
            table_list.append(insert_tuple)
            for item in table_list:
                tree.insert(parent='',index = 0, values = item)

        # Treeview
        columns = ('Dokumentenname', 'Von', 'Bis', 'Pfad')
        tree = ttk.Treeview(tree_frame,columns = columns, show='headings', height=5)
        tree.grid(row=0, column=0, sticky="W")

        #> Headings:
        tree.heading(columns[0], text = columns[0])
        tree.heading(columns[1], text = columns[1])
        tree.heading(columns[2], text = columns[2])
        tree.heading(columns[3], text = columns[3])

        #> Hide column 3
        tree["displaycolumns"] = ("0", "1", "2")

        #> Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="nse")

        #> Drag and Drop into Treeview
        tree.drop_target_register(DND_FILES)
        tree.dnd_bind('<<Drop>>', enter_doc)


        #> Automatically load pages when double clicked
        tree.bind("<Double-1>", select_records)

        #> Automatically apply changes when pressing enter in "Bis:" textbox
        to_box.bind('<Return>', update_docnum)

        #Drag and Drop in tree control - https://stackoverflow.com/a/13354454
        def bDown_Shift(event):
            tv = event.widget
            select = [tv.index(s) for s in tv.selection()]
            select.append(tv.index(tv.identify_row(event.y)))
            select.sort()
            for i in range(select[0],select[-1]+1,1):
                tv.selection_add(tv.get_children()[i])

        def bDown(event):
            tv = event.widget
            if tv.identify_row(event.y) not in tv.selection():
                tv.selection_set(tv.identify_row(event.y))    

        def bUp(event):
            tv = event.widget
            if tv.identify_row(event.y) in tv.selection():
                tv.selection_set(tv.identify_row(event.y))    

        def bMove(event):
            tv = event.widget
            moveto = tv.index(tv.identify_row(event.y))    
            for s in tv.selection():
                tv.move(s, '', moveto)

        tree.bind("<ButtonPress-1>",bDown)
        tree.bind("<ButtonRelease-1>",bUp, add='+')
        tree.bind("<B1-Motion>",bMove, add='+')
        tree.bind("<Shift-ButtonPress-1>",bDown_Shift, add='+')        
        
### Button_Frame Button
        # Hinzufügen
        add_doc = tk.StringVar()
        add_doc.set("Hinzufügen")
        add_btn = tk.Button(button_frame, textvariable=add_doc, command=lambda:enter_doc(), height=1, width=10, relief="groove", bg="#dadada")
        add_btn.grid(column=0,row=0, padx=50,ipadx=20)

        # Löschen
        del_doc = tk.StringVar()
        del_doc.set("Löschen")
        del_btn = tk.Button(button_frame, textvariable=del_doc, command=lambda:delete_doc(), height=1, width=10,  relief="groove", bg="#dadada")
        del_btn.grid(column=2,row=0, padx=50,ipadx=20)

        # Ausführen 
        run_doc = tk.StringVar()
        run_doc.set("Ausführen")
        merge_btn = tk.Button(button_frame, textvariable=run_doc, command=lambda:call_merge(), height=1, width=10, anchor=tk.CENTER,  relief="groove", bg="#dadada")
        merge_btn.grid(column=1,row=0, padx=50, ipadx=20, ipady=15)

        # Vorschau 
        preview_doc = tk.StringVar()
        preview_doc.set("Vorschau")
        preview_btn = tk.Button(button_frame, textvariable=preview_doc, command=lambda:call_preview(), height=1, width=10, anchor=tk.CENTER,  relief="groove", bg="#dadada")
        preview_btn.grid(column=1,row=1, padx=10, ipadx=20, ipady=5)


class ExtractPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

### Layout
        # Controller Frame
        control_frame = tk.Frame(self,bg="#FFA500", width=600)
        control_frame.columnconfigure(0,weight=1)
        control_frame.grid(row=0,column=0, columnspan=3)

        #> Überschrift
        mergeinstructions = tk.Label(control_frame, text = "PDF-Kommentare \n extrahieren", font = "Raleway 10 bold",bg="#FFA500")
        mergeinstructions.grid(column=1,row=0, padx=15, pady=10)

        #> Controller Button
        upper_btn = tk.Button(control_frame, text ="Startseite", command = lambda : controller.show_frame(StartPage), height=1, width = 20, relief="groove", bg="#dadada")
        upper_btn.grid(row=0,column=0, padx=35)

        lower_btn = tk.Button(control_frame, text ="PDF zusammenfassen", command = lambda : controller.show_frame(MergePage), height=1, width = 20, relief="groove", bg="#dadada")
        lower_btn.grid(row=0,column=2, padx=35, ipadx=30) 

        # Input Frame
        input_frame = tk.Frame(self, bg="#FFD580")    
        input_frame.grid(row = 2, column = 0, columnspan=3)

        # Button Frame
        button_frame = tk.Frame(self, bg="#FFD580")
        button_frame.grid(row = 1, column = 0, columnspan=3)
        
        # Tree Frame
        tree_frame = tk.Frame(self, bg="#FFD580")
        tree_frame.grid(row = 3, column = 0, columnspan=3)

        # Anpassung des Fensters
        input_frame.rowconfigure(0,weight=1)
        #> Tabelleninputvariable
        table_fill = []


### Funktionen Button_Frame
        # Ausführen
        def call_extract():
            list_entries = tree.get_children()

            if len(list_entries) == 0:
                showinfo("Uffbasse", "Es muss mindestens ein Dokument hinzugefügt werden! ")
            else:                
                paths = []

                # Fill paths mit Name, Von, Bis, Path
                for entry in list_entries:
                    paths.append(tree.item(entry)['values'][3].replace("{","").replace("}",""))
                
                extract_files(paths,open_doc_directly = open_doc_check_var.get())
            
        # Hinzufügen
        def enter_doc(eventdrop = None):
            if eventdrop != None:
                docpath_r = eventdrop.data.split()
            else:
                docpath = [askopenfilenames(parent=self, title=f"Bitte gebe den Pfad des Dokuments ein: ", filetypes=[("Pdf file", "*.pdf")])]
                docpath_r = [path for path in docpath[0]]
                if docpath[0] == '':
                    return

            for path in docpath_r:    
                filename = path.split("/")[-1].split(".")[0]
                
                reader = PyPDF2.PdfFileReader(path, strict = False)
                count_pages = reader.getNumPages()

                if count_pages == 0:
                    showinfo("Uffbasse", "Das Dokument ist beschädigt und kann nicht geladen werden! ")
                else:
                    fill_table(table_fill,(filename,'1',count_pages,path))
                    table_fill.pop()

        # Löschen
        def delete_doc():
            for selected_item in tree.selection():
                child_id = tree.prev(selected_item)
                tree.delete(selected_item)
                if child_id == "": # Falls kein vorheriges Item mehr in Treeview ist
                    try:  # Falls kein erstes Item mehr vorhanden ist
                        child_id = tree.get_children()[0]        
                    except:
                        child_id = ""
                try: # Falls keine Items mehr in Treeview sind
                    tree.focus(child_id)
                    tree.selection_set(child_id)
                except: return

### Treeview Funktionen
        # Tabelle füllen
        def fill_table(table_list, insert_tuple):
            table_list.append(insert_tuple)
            for item in table_list:
                tree.insert(parent='',index = 0, values = item)

        # Treeview
        columns = ('Dokumentenname', 'Von', 'Bis', 'Pfad')
        tree = ttk.Treeview(tree_frame,columns = columns, show='headings', height=5)
        tree.grid(row=0, column=0, sticky="W",columnspan=4)

        #> Headings:
        tree.heading(columns[0], text = columns[0])
        tree.heading(columns[1], text = columns[1])
        tree.heading(columns[2], text = columns[2])
        tree.heading(columns[3], text = columns[3])

        # Hide column 3
        tree["displaycolumns"] = ("0", "1", "2")

        #> Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.grid(row=0, column=4, sticky="nse")

        #> Drag and Drop
        tree.drop_target_register(DND_FILES)
        tree.dnd_bind('<<Drop>>', enter_doc)

        #Drag and Drop in tree control - https://stackoverflow.com/a/13354454
        def bDown_Shift(event):
            tv = event.widget
            select = [tv.index(s) for s in tv.selection()]
            select.append(tv.index(tv.identify_row(event.y)))
            select.sort()
            for i in range(select[0],select[-1]+1,1):
                tv.selection_add(tv.get_children()[i])

        def bDown(event):
            tv = event.widget
            if tv.identify_row(event.y) not in tv.selection():
                tv.selection_set(tv.identify_row(event.y))    

        def bUp(event):
            tv = event.widget
            if tv.identify_row(event.y) in tv.selection():
                tv.selection_set(tv.identify_row(event.y))    

        def bMove(event):
            tv = event.widget
            moveto = tv.index(tv.identify_row(event.y))    
            for s in tv.selection():
                tv.move(s, '', moveto)

        tree.bind("<ButtonPress-1>",bDown)
        tree.bind("<ButtonRelease-1>",bUp, add='+')
        tree.bind("<B1-Motion>",bMove, add='+')
        tree.bind("<Shift-ButtonPress-1>",bDown_Shift, add='+')

### Button_Frame Button
        # Hinzufügen
        add_doc = tk.StringVar()
        add_doc.set("Hinzufügen")
        add_btn = tk.Button(button_frame, textvariable=add_doc, command=lambda:enter_doc(), height=1, width=10, relief="groove", bg="#dadada")
        add_btn.grid(column=0,row=0, rowspan=2, padx=50,ipadx=20)

        # Löschen
        del_doc = tk.StringVar()
        del_doc.set("Löschen")
        del_btn = tk.Button(button_frame, textvariable=del_doc, command=lambda:delete_doc(), height=1, width=10, relief="groove", bg="#dadada")
        del_btn.grid(column=2,row=0, rowspan=2, padx=50,ipadx=20)

        # Ausführen 
        run_doc = tk.StringVar()
        run_doc.set("Ausführen")
        merge_btn = tk.Button(button_frame, textvariable=run_doc, command=lambda:call_extract(), height=1, width=10, relief="groove", bg="#dadada")
        merge_btn.grid(column=1,row=0, padx=50,ipadx=20, ipady=15)

        #> Check Box, ob Datei 
        open_doc_check_var = tk.IntVar()
        open_doc_check = tk.Checkbutton(button_frame, text='DOCX öffnen', variable = open_doc_check_var, onvalue=True, offvalue=False, bg="#FFD580")
        open_doc_check.grid(column=1, row=1)


### Input_Frame
        # Vorschau anzeigen
        def preview_show():
            preview_box.configure(state='normal')
            preview_box.delete(1.0, 'end')
            # Ausgwähltes Dokument
            selected = tree.focus()
            try: 
                ##### REPLACE WITH REGEX SUB
                doc_path = [tree.item(selected, 'values')[3].replace("{","").replace("}","")] #values[3] is path
                extract_files(doc_path,preview_box)
                preview_box.configure(state='disabled')


            except IndexError:
                showinfo("Uffbasse", "Es muss mindestens eine Datei ausgewählt sein! ")

        # Vorschaufeld
        preview_box = scrolledtext.ScrolledText(input_frame, height=7, width=75)
        preview_box.grid(row=0,column=0, sticky='ewns')
        preview_box.configure(state='disabled')

        # Vorschaubutton
        preview_str = tk.StringVar()
        preview_str.set("Vorschau")
        preview_btn = tk.Button(input_frame, textvariable=preview_str, command=lambda:preview_show(), height=1, width=10, relief="groove", bg="#dadada")
        preview_btn.grid(row=1,column=0,ipadx=20)  

app = tkinterApp()
app.mainloop()