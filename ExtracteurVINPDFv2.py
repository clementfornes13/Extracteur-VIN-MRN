# The VINExtractor class extracts VIN numbers from a PDF file and saves them in an Excel file.

import PySimpleGUI
from os import path, startfile
from re import findall
from PyPDF2 import PdfReader
from openpyxl import Workbook


class VINExtractor:
    def __init__(s):
        
        """
        This is the initialization function for a PySimpleGUI window that allows the user to browse for a
        PDF file and select a destination folder for VIN extraction.
        
        :param s: The parameter "s" is the instance of the class being created. It is commonly used as a
        reference to the instance's attributes and methods within the class
        """
        
        s.ListeVIN = []
        s.IMG_PATH = path.join(path.dirname(__file__), 'images')
        s.VIN_PATTERN = r"\b([A-HJ-NPR-Z0-9]{17})\b"
        s.icon = path.join(s.IMG_PATH, 'Icone.ico')
        s.logo = path.join(s.IMG_PATH, 'Logo TEA FOS.png')
        PySimpleGUI.theme('Reddit')
        s.layout = [
            [PySimpleGUI.Text("Fichier PDF :"), PySimpleGUI.Input(), PySimpleGUI.FileBrowse('Parcourir',file_types=(("Fichiers PDF", "*.pdf"),))],
            [PySimpleGUI.Text("Destination de  l'extraction :"), PySimpleGUI.Input(key='Destination'), PySimpleGUI.FolderBrowse('Parcourir')],
            [PySimpleGUI.Button("Extraction des VIN"), PySimpleGUI.Exit('Quitter'),PySimpleGUI.Image(s.logo, size=(214,50),pad=(50,0))]
        ]
        s.window = PySimpleGUI.Window("Extracteur VIN PDF EAD", s.layout, icon=s.icon)
    def run(s):
        
        """
        This function extracts VIN numbers from a PDF file and saves them in an Excel file.
        
        :param s: The parameter "s" is likely an instance of a class that contains the PySimpleGUI window
        and other variables and methods needed for the program to run
        """
        
        while True:
            event, values = s.window.read()
            if event == PySimpleGUI.WIN_CLOSED or event == "Quitter":
                break
            elif event == "Extraction des VIN":
                emplacement_pdf = values[0]
                if emplacement_pdf:
                    fichier_pdf = open(emplacement_pdf, 'rb')
                    lire_pdf = PdfReader(fichier_pdf)
                    workbook = Workbook()
                    sheet = workbook.active
                    sheet['A1'] = 'Liste des VIN'
                    for page_num in range(len(lire_pdf.pages)):
                        page = lire_pdf.pages[page_num]
                        texte = page.extract_text()
                        vins = findall(s.VIN_PATTERN, texte)
                        for vin in vins:
                            s.ListeVIN.append(vin)
                            s.ListeVIN = list(set(s.ListeVIN))
                    for vin in s.ListeVIN:
                        sheet['A' + str(sheet.max_row + 1)] = vin
                    file_name = "Extraction VIN EAD.xlsx".format(1)
                    if values['Destination']:
                        file_path = path.join(values['Destination'], file_name)
                        i = 1
                        while path.exists(file_path):
                            i += 1
                            file_name = "Extraction VIN EAD {}.xlsx".format(i)
                            file_path = path.join(values['Destination'], file_name)
                        workbook.save(file_path)
                        PySimpleGUI.popup("Fini!", f"Fichier enregistré ici : {file_path}", f"{len(s.ListeVIN)} VIN extraits", "Le fichier va s'ouvrir...")
                        startfile(file_path)
                        s.ListeVIN = []
                    else:
                        PySimpleGUI.popup('Pas de destination', title='Erreur')
                else:
                    PySimpleGUI.popup('Pas de fichier sélectionné, veuillez en sélectionner un !', title='Erreur', icon=s.icon)
                    s.ListeVIN = []
        s.window.close()
        
# This code block is checking if the current script is being run as the main program (as opposed to
# being imported as a module into another program). If it is being run as the main program, it creates
# an instance of the VINExtractor class and calls its run() method, which starts the PySimpleGUI
# window and runs the VIN extraction program.

if __name__ == '__main__':
    extracteur = VINExtractor()
    extracteur.run()