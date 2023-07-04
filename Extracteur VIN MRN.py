##############################################################################
# Nom du fichier: Extracteur VIN MRN.py
# Description: Ce script permet d'extraire les VIN et MRN de fichiers PDF
#
# Auteur: FORNES Clément
# Date: 4 Juillet 2023
#
# Copyright (c) 2023 FORNES Clément
#
# Licence: MIT License
##############################################################################
# MIT License
#
# Copyright (c) 2023 FORNES Clément
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#
# ------------------------------------------------------------------------------
#
# Ce script permet d'extraire les VIN et MRN de fichiers PDF
# Il est possible d'extraire les VIN, les MRN ou les deux
# Les VIN et MRN sont enregistrés dans un fichier Excel
# Le fichier Excel est enregistré dans le dossier de destination
# Mise à jour: 4 Juillet 2023
#################### IMPORTS ####################
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox, PhotoImage, IntVar, END
from tkinter.ttk import Progressbar, Frame, Radiobutton
from openpyxl import Workbook
from PyPDF2 import PdfReader
from os import path, startfile, listdir
from re import findall
import threading
#################################################

#################### CLASSES ####################
class VINMRNExtractor(Tk):

    def __init__(self):
        #################### VARIABLES ####################
        self.ListeVIN, self.ListeMRN, self.unique_VIN, self.unique_MRN = [], [], [], []
        self.IMG_PATH = path.join(path.dirname(__file__), 'images')
        # Patterne des VIN 
        # Explication : 
        # (?!FR) : Ne pas commencer par FR
        # [A-HJ-NPR-Z] : Lettre de A à Z sauf I, O et Q
        # Même chose pour les 15 caractères suivants
        # [A-HJ-NPR-Z\d] : Lettre de A à Z sauf I, O et Q ou chiffre
        self.VIN_PATTERN = r"(?!FR)[A-HJ-NPR-Z][A-HJ-NPR-Z0-9]{15}[A-HJ-NPR-Z\d]"
        # Patterne des MRN
        # Explication :
        # (?<!\S) : Ne pas commencer par un caractère
        # (?!MRN) : Ne pas commencer par MRN
        # (?=[A-Z0-9]{18}\b) : Commencer par 18 caractères alphanumériques
        # [A-Z0-9]*[A-Z][A-Z0-9]*\b : Finir par un caractère alphanumérique
        # (?!\S) : Ne pas finir par un caractère
        self.MRN_PATTERN = r"(?<!\S)(?!MRN)(?=[A-Z0-9]{18}\b)[A-Z0-9]*[A-Z][A-Z0-9]*\b(?!\S)"
        #################################################
        # Initialisation de la fenêtre principale
        super().__init__()
        self.icon = path.join(self.IMG_PATH, 'Icone.ico').replace('\\', '/')
        self.logo = path.join(self.IMG_PATH, 'Logo TEA FOS.png').replace('\\', '/')
        self.initUI()

    def initUI(self):
        # Initialisation de l'interface graphique
        # Paramètres de la fenêtre
        self.title("Extracteur VIN MRN")
        self.iconbitmap(self.icon)
        self.geometry("350x350")
        logoLabel = Label(self)
        logoImage = PhotoImage(file=self.logo)
        logoLabel.config(image=logoImage)
        logoLabel.image = logoImage
        logoLabel.grid(row=0, column=0, pady=10)
        # Frame
        frame = Frame(self)
        frame.grid(row=1, column=0)
        # Dossier PDF
        labelPDF = Label(frame, text="Dossier PDF:")
        labelPDF.grid(row=0, column=0, padx=5, pady=5)
        self.inputPDF = Entry(frame, width=30)
        self.inputPDF.grid(row=0, column=1, padx=5, pady=5)
        browsePDF = Button(frame, text="Parcourir", command=self.browsePDFClicked)
        browsePDF.grid(row=0, column=2, padx=5, pady=5)
        # Destination
        labelDest = Label(frame, text="Destination:")
        labelDest.grid(row=1, column=0, padx=5, pady=5)
        self.inputDest = Entry(frame, width=30)
        self.inputDest.grid(row=1, column=1, padx=5, pady=5)
        browseDest = Button(frame, text="Parcourir", command=self.browseDestClicked)
        browseDest.grid(row=1, column=2, padx=5, pady=5)
        # Boutons
        self.extractButton = Button(self, text="Extraction", command=self.launch_extraction)
        self.extractButton.grid(row=5, column=0, pady=10)
        # Options d'extraction
        self.extractOption = IntVar()
        # 1 = Extraction VIN
        self.vinRadio = Radiobutton(self, text="VIN", variable=self.extractOption, value=1)
        self.vinRadio.grid(row=2, column=0, padx=5, pady=5, sticky='n')
        # 2 = Extraction MRN
        self.mrnRadio = Radiobutton(self, text="MRN", variable=self.extractOption, value=2)
        self.mrnRadio.grid(row=3, column=0, padx=5, pady=5, sticky='n')
        # 3 = Extraction VIN et MRN
        self.vinMrnRadio = Radiobutton(self, text="VIN + MRN", variable=self.extractOption, value=3)
        self.vinMrnRadio.grid(row=4, column=0, padx=5, pady=5, sticky='n')
        # Barre de progression
        self.progress_bar = Progressbar(self, mode='determinate')
        self.progress_bar.grid(row=6, column=0, pady=5)
        self.progress_bar.grid_remove()

    # Désactiver les boutons
    def disable_buttons(self):
        self.extractButton.config(state="disabled")
        self.vinRadio.config(state="disabled")
        self.mrnRadio.config(state="disabled")
        self.vinMrnRadio.config(state="disabled")

    # Activer les boutons
    def enable_buttons(self):
        self.extractButton.config(state="normal")
        self.vinRadio.config(state="normal")
        self.mrnRadio.config(state="normal")
        self.vinMrnRadio.config(state="normal")

    # Fenêtre de sélection du dossier PDF
    def browsePDFClicked(self):
        # Selection du dossier contenant les PDF
        if directory := filedialog.askdirectory(title="Choisir un dossier"):
            self.inputPDF.delete(0, END)
            self.inputPDF.insert(0, directory)

    # Fenêtre de sélection du dossier de destination
    def browseDestClicked(self):
        # Selection du dossier de destination
        if directory := filedialog.askdirectory(title="Choisir un dossier"):
            self.inputDest.delete(0, END)
            self.inputDest.insert(0, directory)

    # Extraction des VIN
    def extract_vins(self, emplacement_pdf, destination):
        if not emplacement_pdf:
            return None
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = 'Liste des VIN'
        row_num = 2
        file_list = [pdf for pdf in listdir(emplacement_pdf) if pdf.lower().endswith('.pdf')]
        total_files = len(file_list)
        self.progress_bar['maximum'] = total_files
        self.progress_bar.grid()
        for file_count, pdf in enumerate(file_list, start=1):
            pdf = path.join(emplacement_pdf, pdf).replace("\\", "/")
            with open(pdf, 'rb') as file:
                lire_pdf = PdfReader(file)
                for page_num, page in enumerate(lire_pdf.pages):
                    texte = page.extract_text()
                    vins = findall(self.VIN_PATTERN, texte)
                    self.ListeVIN.extend(vins)
                    for vin_unique in set(self.ListeVIN):
                        self.unique_VIN.append(vin_unique)
                    for vin in self.unique_VIN:
                        sheet.cell(row=row_num, column=1, value=vin)
                        row_num += 1
                    self.ListeVIN.clear(), self.unique_VIN.clear()
            file_count += 1
            self.progress_bar['value'] = file_count
            self.update()
        if not destination:
            self.enable_buttons()
            return None
        file_name = "Extraction VIN EAD.xlsx".format(1)
        file_path = path.join(destination, file_name)
        i = 1
        while path.exists(file_path):
            i += 1
            file_name = f"Extraction VIN EAD {i}.xlsx"
            file_path = path.join(destination, file_name)
        workbook.save(file_path)
        return file_path
    
    # Extraction des MRN
    def extract_mrns(self, emplacement_pdf, destination):
        if not emplacement_pdf:
            return None
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = 'Liste des MRN'
        row_num = 2
        file_list = [pdf for pdf in listdir(emplacement_pdf) if pdf.lower().endswith('.pdf')]
        total_files = len(file_list)
        self.progress_bar['maximum'] = total_files
        self.progress_bar.grid()
        for file_count, pdf in enumerate(file_list, start=1):
            pdf = path.join(emplacement_pdf, pdf).replace("\\", "/")
            with open(pdf, 'rb') as file:
                lire_pdf = PdfReader(file)
                for page_num, page in enumerate(lire_pdf.pages):
                    texte = page.extract_text()
                    mrns = findall(self.MRN_PATTERN, texte)
                    self.ListeMRN.extend(mrns)
                    for mrn_unique in set(self.ListeMRN):
                        self.unique_MRN.append(mrn_unique)
                    for mrn in self.unique_MRN:
                        sheet.cell(row=row_num, column=1, value=mrn)
                        row_num += 1
                    self.ListeMRN.clear(), self.unique_MRN.clear()
            file_count += 1
            self.progress_bar['value'] = file_count
            self.update()
        if not destination:
            return None
        file_name = "Extraction MRN EAD.xlsx".format(1)
        file_path = path.join(destination, file_name)
        i = 1
        while path.exists(file_path):
            i += 1
            file_name = f"Extraction MRN EAD {i}.xlsx"
            file_path = path.join(destination, file_name)
        workbook.save(file_path)
        return file_path
    
    # Extraction des VIN et MRN
    def extract_vins_mrns(self, emplacement_pdf, destination):
        if not emplacement_pdf:
            return None
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = 'Liste des VIN'
        sheet['B1'] = 'Liste des MRN'
        row_num = 2
        file_list = [pdf for pdf in listdir(emplacement_pdf) if pdf.lower().endswith('.pdf')]
        total_files = len(file_list)
        self.progress_bar['maximum'] = total_files
        self.progress_bar.grid()
        for file_count, pdf in enumerate(file_list, start=1):
            pdf = path.join(emplacement_pdf, pdf).replace("\\", "/")
            with open(pdf, 'rb') as file:
                lire_pdf = PdfReader(file)
                for page_num, page in enumerate(lire_pdf.pages):
                    texte = page.extract_text()
                    vins = findall(self.VIN_PATTERN, texte)
                    mrns = findall(self.MRN_PATTERN, texte)
                    self.ListeVIN.extend(vins)
                    self.ListeMRN.extend(mrns)
                    for vin_unique in set(self.ListeVIN):
                        self.unique_VIN.append(vin_unique)
                    for vin in self.unique_VIN:
                        sheet.cell(row=row_num, column=1, value=vin)
                        sheet.cell(row=row_num, column=2, value=self.ListeMRN[0])
                        row_num += 1
                    self.ListeVIN.clear(), self.ListeMRN.clear(), self.unique_VIN.clear()
            file_count += 1
            self.progress_bar['value'] = file_count
            self.update()
        if not destination:
            return None
        file_name = "Extraction VIN MRN EAD.xlsx".format(1)
        file_path = path.join(destination, file_name)
        i = 1
        while path.exists(file_path):
            i += 1
            file_name = f"Extraction VIN MRN EAD {i}.xlsx"
            file_path = path.join(destination, file_name)
        workbook.save(file_path)
        return file_path
    
    # Mise à jour de la barre de progression
    def update_progress(self, value, maximum):
        self.progress_bar['value'] = value
        self.progress_bar['maximum'] = maximum
        self.update()
        
    # Chemin de fichier temporaire
    def extraction_temp(self, extraction_func, args, file_path_temp):
        file_path_temp[0] = extraction_func(*args)
        
    # Message d'erreur
    def error_message(self, arg0):
        messagebox.showerror('Erreur', arg0)
        self.enable_buttons()
        return None
    
    # Lancement de l'extraction
    def launch_extraction(self):
        self.disable_buttons()
        file_path_temp = [None]
        emplacement_pdf = self.inputPDF.get()
        destination = self.inputDest.get()
        file_list = [pdf for pdf in listdir(emplacement_pdf) if pdf.lower().endswith('.pdf')]
        total_files = len(file_list)
        if total_files == 0:
            return self.error_message(
                'Le dossier PDF sélectionné est vide'
            )
        extract_option = self.extractOption.get()
        if extract_option == 1:
            thread = threading.Thread(target=self.extraction_temp, args=(self.extract_vins, (emplacement_pdf, destination), file_path_temp))
        elif extract_option == 2:
            thread = threading.Thread(target=self.extraction_temp, args=(self.extract_mrns, (emplacement_pdf, destination), file_path_temp))
        else:
            thread = threading.Thread(target=self.extraction_temp, args=(self.extract_vins_mrns, (emplacement_pdf, destination), file_path_temp))
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
        file_path = file_path_temp[0]
        if not file_path:
            return self.error_message(
                'Pas de Dossier PDF / Destination choisi'
            )
        messagebox.showinfo('Extraction réussie', f"Fichier enregistré ici : {file_path}")
        startfile(file_path)
        self.progress_bar.grid_remove()
        self.enable_buttons()

#################################################
app = VINMRNExtractor()
app.resizable(False, False)
app.mainloop()
#################################################