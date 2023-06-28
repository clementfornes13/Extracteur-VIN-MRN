from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
from openpyxl import Workbook
from PyPDF2 import PdfReader
import os
import re
import threading

class VINMRNExtractor(Tk):
    def __init__(self):
        super().__init__()
        self.ListeVIN = []
        self.ListeMRN = []
        self.unique_VIN = []
        self.unique_MRN = []
        self.IMG_PATH = os.path.join(os.path.dirname(__file__), 'images')
        self.VIN_PATTERN = r"(?!FR)[A-HJ-NPR-Z][A-HJ-NPR-Z0-9]{15}[A-HJ-NPR-Z\d]"
        self.MRN_PATTERN = r"(?<!\S)(?!MRN)(?=[A-Z0-9]{18}\b)[A-Z0-9]*[A-Z][A-Z0-9]*\b(?!\S)"
        self.icon = os.path.join(self.IMG_PATH, 'Icone.ico').replace('\\', '/')
        self.logo = os.path.join(self.IMG_PATH, 'Logo TEA FOS.png').replace('\\', '/')
        self.initUI()

    def initUI(self):
        self.title("Extracteur VIN MRN PDF EAD")
        self.iconbitmap(self.icon)
        self.geometry("450x200")

        logoLabel = Label(self)
        logoImage = PhotoImage(file=self.logo)
        logoLabel.config(image=logoImage)
        logoLabel.image = logoImage
        logoLabel.grid(row=0, column=0, pady=10)

        frame = Frame(self)
        frame.grid(row=1, column=0)

        labelPDF = Label(frame, text="Dossier PDF:")
        labelPDF.grid(row=0, column=0, padx=5, pady=5)
        self.inputPDF = Entry(frame, width=30)
        self.inputPDF.grid(row=0, column=1, padx=5, pady=5)
        browsePDF = Button(frame, text="Parcourir", command=self.browsePDFClicked)
        browsePDF.grid(row=0, column=2, padx=5, pady=5)

        labelDest = Label(frame, text="Destination:")
        labelDest.grid(row=1, column=0, padx=5, pady=5)
        self.inputDest = Entry(frame, width=30)
        self.inputDest.grid(row=1, column=1, padx=5, pady=5)
        browseDest = Button(frame, text="Parcourir", command=self.browseDestClicked)
        browseDest.grid(row=1, column=2, padx=5, pady=5)

        extractButton = Button(self, text="Extraction", command=self.launch_extraction)
        extractButton.grid(row=2, column=0, pady=10)

        self.progress_bar = Progressbar(self, mode='determinate')
        self.progress_bar.grid(row=3, column=0, pady=5)
        self.progress_bar.grid_remove()

        self.extractOption = IntVar()

        vinRadio = Radiobutton(self, text="Extraction VIN", variable=self.extractOption, value=1)
        vinRadio.grid(row=4, column=0, padx=5, pady=5)

        mrnRadio = Radiobutton(self, text="Extraction MRN", variable=self.extractOption, value=2)
        mrnRadio.grid(row=5, column=0, padx=5, pady=5)

        vinMrnRadio = Radiobutton(self, text="Extraction VIN et MRN", variable=self.extractOption, value=3)
        vinMrnRadio.grid(row=6, column=0, padx=5, pady=5)
        
    def browsePDFClicked(self):
        if directory := filedialog.askdirectory(title="Choisir un dossier"):
            self.inputPDF.delete(0, END)
            self.inputPDF.insert(0, directory)

    def browseDestClicked(self):
        if directory := filedialog.askdirectory(title="Choisir un dossier"):
            self.inputDest.delete(0, END)
            self.inputDest.insert(0, directory)
    def extract_vins(self, emplacement_pdf, destination=None):
        if not emplacement_pdf:
            return None
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = 'Liste des VIN'
        row_num = 2
        file_list = [pdf for pdf in os.listdir(emplacement_pdf) if pdf.lower().endswith('.pdf')]
        total_files = len(file_list)
        self.progress_bar['maximum'] = total_files
        self.progress_bar.grid()
        for file_count, pdf in enumerate(file_list, start=1):
            pdf = os.path.join(emplacement_pdf, pdf).replace("\\", "/")
            with open(pdf, 'rb') as file:
                lire_pdf = PdfReader(file)
                for page_num, page in enumerate(lire_pdf.pages):
                    texte = page.extract_text()
                    vins = re.findall(self.VIN_PATTERN, texte)
                    self.ListeVIN.extend(vins)
                    for vin_unique in set(self.ListeVIN):
                        self.unique_VIN.append(vin_unique)
                    for vin in self.unique_VIN:
                        sheet.cell(row=row_num, column=1, value=vin)
                        row_num += 1
                    self.ListeVIN.clear()
                    self.unique_VIN.clear()
            file_count += 1
            self.progress_bar['value'] = file_count
            self.update()
        if not destination:
            return None
        file_name = "Extraction VIN EAD.xlsx".format(1)
        file_path = os.path.join(destination, file_name)
        i = 1
        while os.path.exists(file_path):
            i += 1
            file_name = f"Extraction VIN EAD {i}.xlsx"
            file_path = os.path.join(destination, file_name)
        workbook.save(file_path)
        return file_path
    def extract_mrns(self, emplacement_pdf, destination=None):
        if not emplacement_pdf:
            return None
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = 'Liste des MRN'

        row_num = 2
        file_list = [pdf for pdf in os.listdir(emplacement_pdf) if pdf.lower().endswith('.pdf')]
        total_files = len(file_list)
        self.progress_bar['maximum'] = total_files
        self.progress_bar.grid()
        for file_count, pdf in enumerate(file_list, start=1):
            pdf = os.path.join(emplacement_pdf, pdf).replace("\\", "/")
            with open(pdf, 'rb') as file:
                lire_pdf = PdfReader(file)
                for page_num, page in enumerate(lire_pdf.pages):
                    texte = page.extract_text()
                    mrns = re.findall(self.MRN_PATTERN, texte)
                    self.ListeMRN.extend(mrns)
                    for mrn_unique in set(self.ListeMRN):
                        self.unique_MRN.append(mrn_unique)
                    for mrn in self.unique_MRN:
                        sheet.cell(row=row_num, column=1, value=mrn)
                        row_num += 1
                    self.ListeMRN.clear()
                    self.unique_MRN.clear()
            file_count += 1
            self.progress_bar['value'] = file_count
            self.update()
        if not destination:
            return None
        file_name = "Extraction MRN EAD.xlsx".format(1)
        file_path = os.path.join(destination, file_name)
        i = 1
        while os.path.exists(file_path):
            i += 1
            file_name = f"Extraction MRN EAD {i}.xlsx"
            file_path = os.path.join(destination, file_name)
        workbook.save(file_path)
        return file_path
    def extract_vins_mrns(self, emplacement_pdf, destination=None):
        if not emplacement_pdf:
            return None
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = 'Liste des VIN'
        sheet['B1'] = 'Liste des MRN'
        row_num = 2
        file_list = [pdf for pdf in os.listdir(emplacement_pdf) if pdf.lower().endswith('.pdf')]
        total_files = len(file_list)
        self.progress_bar['maximum'] = total_files
        self.progress_bar.grid()
        for file_count, pdf in enumerate(file_list, start=1):
            pdf = os.path.join(emplacement_pdf, pdf).replace("\\", "/")
            with open(pdf, 'rb') as file:
                lire_pdf = PdfReader(file)
                for page_num, page in enumerate(lire_pdf.pages):
                    texte = page.extract_text()
                    vins = re.findall(self.VIN_PATTERN, texte)
                    mrns = re.findall(self.MRN_PATTERN, texte)
                    self.ListeVIN.extend(vins)
                    self.ListeMRN.extend(mrns)
                    for vin_unique in set(self.ListeVIN):
                        self.unique_VIN.append(vin_unique)
                    for vin in self.unique_VIN:
                        sheet.cell(row=row_num, column=1, value=vin)
                        sheet.cell(row=row_num, column=2, value=self.ListeMRN[0])
                        row_num += 1
                    self.ListeVIN.clear()
                    self.ListeMRN.clear()
                    self.unique_VIN.clear()
            file_count += 1
            self.progress_bar['value'] = file_count
            self.update()
        if not destination:
            return None
        file_name = "Extraction VIN MRN EAD.xlsx".format(1)
        file_path = os.path.join(destination, file_name)
        i = 1
        while os.path.exists(file_path):
            i += 1
            file_name = f"Extraction VIN MRN EAD {i}.xlsx"
            file_path = os.path.join(destination, file_name)
        workbook.save(file_path)
        return file_path
    
    def update_progress(self, value, maximum):
        self.progress_bar['value'] = value
        self.progress_bar['maximum'] = maximum
        self.update()
        
    def launch_extraction(self):
        emplacement_pdf = self.inputPDF.get()
        destination = self.inputDest.get()
        file_list = [pdf for pdf in os.listdir(emplacement_pdf) if pdf.lower().endswith('.pdf')]
        total_files = len(file_list)
        if total_files == 0:
            messagebox.showerror('Erreur', 'Le dossier PDF sélectionné est vide')
            return None
        extract_option = self.extractOption.get()
        if extract_option == 1:
            thread = threading.Thread(target=self.extract_vins, args=(emplacement_pdf, destination))
        elif extract_option == 2:
            thread = threading.Thread(target=self.extract_mrns, args=(emplacement_pdf, destination))
        else:
            thread = threading.Thread(target=self.extract_vins_mrns, args=(emplacement_pdf, destination))
        thread.start()
        self.progress_bar.grid()
        stop_flag = False
        while thread.is_alive() or threading.active_count() > 1:
            if not thread.is_alive() and threading.active_count() == 1:
                stop_flag = True
            self.update_progress(threading.active_count() - 1, total_files)
            if stop_flag:
                break
        self.progress_bar.grid_remove()
        for t in threading.enumerate():
            if t != threading.current_thread():
                t.join()
        file_path = self.extract_vins(emplacement_pdf, destination) if extract_option == 1 else \
                self.extract_mrns(emplacement_pdf, destination) if extract_option == 2 else \
                self.extract_vins_mrns(emplacement_pdf, destination)
        if not file_path:
            messagebox.showerror('Erreur', 'Pas de Dossier PDF / Destination choisi')
            return None
        messagebox.showinfo('Extraction réussie', f"Fichier enregistré ici : {file_path}")
        os.startfile(file_path)

app = VINMRNExtractor()
app.mainloop()