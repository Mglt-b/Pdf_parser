import os
import pkg_resources.py2_warn
import csv
import glob
from parse_obj import parse_obj

ver = '1.1'
titre = 'Syno PDF vers FIELDWIRE'

print("########################")
print("Syno PDF vers FIELDWIRE")
print("Par MIGLIORATI Bastien")
print("Version : " + str(ver))
print("########################")
print("")

think_its_safe = input('Password : ')
      
#pour parse pdf
import pdfminer
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator

#pour dialogue
from tkinter import filedialog
import tkinter as tk
from tkinter import LabelFrame, Label, Tk
from tkinter.ttk import Notebook
import tkinter.messagebox

#pour export excel
#import xlrd
from xlwt import Workbook, Formula

#pour date et time
import time
import subprocess

#definition de la date
t = time.localtime()
current_time = time.strftime("%H_%M_%S", t)

#correlation temporelle des biblioteques
if str(think_its_safe) not in str(time.strftime("%H_%M", t)):
    print(str(bytes.fromhex('534f525259').decode('utf-8')))
    print(str(bytes.fromhex('57 52 4f 4e 47 20').decode('utf-8'))+str(bytes.fromhex('50 41 53 53 57 4f 52 44').decode('utf-8')))
    exit()

#tenter de lier les taches ?
link_task_sc = input('Entrez le mode complementaire (optionnel): \n o pour lier les taches du plan de tirage \n s pour lier les plans de soudure \n ')

#is_reverse
is_reverse = input('Besoin d inversion de coordonnées ? y0/y1/no : ')
if len(is_reverse) != 2:
    print('Mauvais parametre entré (f or e), fin. Exactement deux caractères svp.')
    exit()  
    if not (('y0' in str(is_reverse)) or ('y1' in str(is_reverse)) or ('no' in str(is_reverse))):
        print('Mauvais parametre entré (y0 or y1 or no), fin.')
        exit()

#emplacement et definition du fichier d'export
path = r"C:\Export_Syno_Fieldwire\export_" + current_time + ".xls"
# On créer un "classeur"
classeur = Workbook()
# On ajoute une feuille au classeur
feuille = classeur.add_sheet("Fieldwire_syno")
feuille2 = classeur.add_sheet("Listing_cables_boites")

#Interface graphique
#la fenetre
import sys
import PyQt5.QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel
root = Tk()
class Fenetre(QWidget):
    def __init__(self):
        QWidget.__init__(self)

        # creation du bouton   
        self.label = QLabel("")
        self.label1 = QLabel("Export terminé")
        self.label2 = QLabel("Par MIGLIORATI Bastien")
        self.label3 = QLabel("bastien.migliorati@gmail.com")
        self.label4 = QLabel("https://github.com/Migliorati/Pdf_parser")
             
        # creation du gestionnaire de mise en forme
        layout1 = QVBoxLayout()
        layout1.addWidget(self.label)
        layout1.addWidget(self.label1)
        layout1.addWidget(self.label2)
        layout1.addWidget(self.label3)
        layout1.addWidget(self.label4)
        self.setLayout(layout1) 
        self.setWindowTitle("PDF vers Fieldwire")

    #user choose path
    outputpath = filedialog.askopenfile()
    if not os.path.exists(r"C:\Export_Syno_Fieldwire"):
        os.makedirs(r"C:\Export_Syno_Fieldwire")

    #oon verifie si l'utilisateur a bien chosis un pdf
    if str(outputpath) == '' :
        print("Aucun fichier n'a été selectionné, fin.")
        exit()

    #securite si pas de pdf
    if not 'pdf' in str(outputpath):
        print("Veuillez selectionner un fichier pdf. Fin.")
        exit()
        
    if not str(outputpath.name) == '' :
        root.withdraw()
        from os.path import basename
        filename = basename(outputpath.name).replace('.pdf', '')

    # Open a PDF file.
    fp = open(str(outputpath.name), 'rb')

    # Create a PDF parser object associated with the file object.
    parser = PDFParser(fp)

    # Create a PDF document object that stores the document structure.
    document = PDFDocument(parser)
    print(str(document))
    # Check if the document allows text extraction. If not, abort.
    if not document.is_extractable:
        print('Non autorise')
        raise PDFTextExtractionNotAllowed
    print('Extraction possible')

    # Create a PDF resource manager object that stores shared resources.
    rsrcmgr = PDFResourceManager()
    print('Creation d un fichier de resources')
    # Create a PDF device object.
    device = PDFDevice(rsrcmgr)
    # BEGIN LAYOUT ANALYSIS: Set parameters for analysis.
    laparams = LAParams()
    print('Definition des parametres pour analyse')
    # Create a PDF page aggregator object.
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    # Create a PDF interpreter object.
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # loop over all pages in the document
    for page in PDFPage.create_pages(document):

        print('Lecture et mise en cache de tous les elements...')
        print('Veuillez patienter quelques minutes')
        # read the page into a layout object
        interpreter.process_page(page)
        layout = device.get_result()

        print(str(page))
        print('layout : ' + str(layout))
        print(str(device))
        
        #securite si tupple vide
        if not all(layout._objs):
            print('Le script n arrive pas a recuperer les données, essayer un autre PDF. FIN')
            exit()

        #recuperer le format de la page
        x1 = float(page.mediabox[2]) #paysage
        y1 = float(page.mediabox[3]) #paysage

        if float(y1) > float(x1) and str(is_reverse) == 'no':
            x1 = float(page.mediabox[3]) #portait
            y1 = float(page.mediabox[2]) #portait
        
        # extract text from this object
        parse_obj(layout._objs, x1, y1, filename, is_reverse, feuille, feuille2, path, link_task_sc, classeur)

app = QApplication.instance() 
if not app:
    app = QApplication(sys.argv)
    
fen = Fenetre()
fen.show()
