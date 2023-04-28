#Auto-generated comments by Mintlify Doc Writer

# These lines are importing necessary libraries/modules for the Python script.

import PySimpleGUI as sg
from os import path, startfile
from re import findall
from PyPDF2 import PdfReader
from openpyxl import Workbook

# `VIN_PATTERN` is a regular expression pattern that matches a string of 17 characters that contains
# only the characters A-H, J-N, P-R, Z, and 0-9. This pattern is specifically designed to match
# Vehicle Identification Numbers (VINs) which are unique identifiers assigned to motor vehicles. The
# `r` before the pattern indicates that it is a raw string, which means that backslashes are treated
# as literal backslashes rather than escape characters. The `\b` at the beginning and end of the
# pattern indicate word boundaries, which means that the pattern will only match if the VIN is
# surrounded by non-alphanumeric characters or the beginning/end of the string.

IMG_PATH = path.join(path.dirname(__file__), 'images')
VIN_PATTERN = r"\b([A-HJ-NPR-Z0-9]{17})\b" 
ListeVIN=[]
icon = path.join(IMG_PATH, 'Icone.ico')
logo = path.join(IMG_PATH, 'Logo TEA FOS.png')
# This code block is creating a GUI window using the PySimpleGUI library. The `sg.theme('Reddit')`
# sets the theme of the window to the Reddit theme. The `layout` variable defines the layout of the
# window, including text inputs, file and folder browsing buttons, and an image. The `sg.Window`
# function creates the window with the specified title, layout, and icon.

sg.theme('Reddit')
layout = [
    [sg.Text("Fichier PDF :"), sg.Input(), sg.FileBrowse('Parcourir',file_types=(("Fichiers PDF", "*.pdf"),))],
    [sg.Text("Destination de  l'extraction :"), sg.Input(key='Destination'), sg.FolderBrowse('Parcourir')],
    [sg.Button("Extraction des VIN"), sg.Exit('Quitter'),sg.Image(logo, size=(214,50),pad=(50,0))]
    ] 
window = sg.Window("Extracteur VIN PDF EAD", layout,icon=icon) 

# This code block is creating an infinite loop that continuously reads events from the GUI window
# created using PySimpleGUI library. The loop will continue until the user clicks the "Quitter" button
# or closes the window. If the user clicks the "Extraction des VIN" button, the script will extract
# Vehicle Identification Numbers (VINs) from a PDF file selected by the user and save them to an Excel
# file. The script also includes error handling to ensure that the user selects a PDF file and a
# destination folder for the Excel file.

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Quitter": 
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
                vins = findall(VIN_PATTERN, texte)
                for vin in vins:
                    ListeVIN.append(vin)
                    ListeVIN = list(set(ListeVIN))
            for vin in ListeVIN:
                sheet['A' + str(sheet.max_row + 1)] = vin
            file_name = "Extraction VIN EAD.xlsx".format(1)
            if values['Destination']:
                file_path = path.join(values['Destination'],file_name) 
                i = 1
                while path.exists(file_path):
                    i += 1
                    file_name = "Extraction VIN EAD {}.xlsx".format(i)
                    file_path = path.join(values['Destination'],file_name)
                workbook.save(file_path)
                sg.popup("Fini!", f"Fichier enregistré ici : {file_path}", f"{len(ListeVIN)} VIN extraits", "Le fichier va s'ouvrir...") 
                startfile(file_path)
                ListeVIN=[]
            else:
                sg.popup('Pas de destination')
        else:
            sg.popup('Pas de fichier sélectionné, veuillez en sélectionner un !',title='Erreur',icon=icon)
            ListeVIN=[]