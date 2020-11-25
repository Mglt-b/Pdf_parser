import os
import pkg_resources.py2_warn
import csv
import glob
from parse_obj import parse_obj

ver = '1.5'
titre = 'Syno PDF vers FIELDWIRE'

print("########################")
print("Syno PDF vers FIELDWIRE")
print("Par MIGLIORATI Bastien")
print("Version : " + str(ver))
print("########################")
print("")
print("Veuillez privilégier des marges, en-têtes et pieds de pages à 0 lors de génération pdf depuis excel.")
print("")

#think_its_safe = input('Password : ')
      
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

#password
#if str(think_its_safe) != str(time.strftime("%H_%M", t)):
    #print(str(bytes.fromhex('534f525259').decode('utf-8')))
    #print(str(bytes.fromhex('57 52 4f 4e 47 20').decode('utf-8'))+str(bytes.fromhex('50 41 53 53 57 4f 52 44').decode('utf-8')))
    #exit()

#tenter de lier les taches ?
link_task_sc = input('Entrez le mode complementaire (optionnel): \n Tappez "o" pour lier les taches du plan de tirage (voir documentation) \n')

is_reverse = 'no'

''' n est plus utilisé, on essaye de deuire tout seul le reverse
is_reverse = 'input('Besoin d inversion de coordonnées X/Y ? Entrez un des deux paramètres : "y0" / "no" : ')'
if len(is_reverse) != 2:
    print('Mauvais parametre entré (f or e), fin. Exactement deux caractères svp.')
    exit()  
    if not (('y0' in str(is_reverse)) or ('no' in str(is_reverse))):
        print('Mauvais parametre entré (y0 or no), fin.')
        exit()
'''



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

    #emplacement et definition du fichier d'export
    path = r"C:\Export_Syno_Fieldwire\export_" + str(filename) + str('-') + current_time + ".xls"

    # Open a PDF file.
    fp = open(str(outputpath.name), 'rb')

    # Create a PDF parser object associated with the file object.
    parser = PDFParser(fp)

    # Create a PDF document object that stores the document structure.
    document = PDFDocument(parser, caching=False)
    #print(str(document))
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
    #laparams = LAParams()
    laparams = LAParams(detect_vertical=True)
    print('Definition des parametres pour analyse')
    # Create a PDF page aggregator object.
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    # Create a PDF interpreter object.
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # loop over all pages in the document

    page_num = 0
    compteur = 0
    # On créer un "classeur"
    classeur = Workbook()
    # On ajoute une feuille au classeur
    feuille = classeur.add_sheet("Fieldwire_syno")

    
    for page in PDFPage.create_pages(document):
        page_num = page_num + 1
        print('-'*20)
        print(filename + str('-') + str(page_num))

        print('Lecture et mise en cache de tous les elements')
        print("Veuillez patienter quelques instants (jusqu'à 5 minutes par page)...")
        # read the page into a layout object
        interpreter.process_page(page)
        layout = device.get_result()

        #print(str(page))
        print('layout : ' + str(layout))
        #print(str(device))
        
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
        compteur = parse_obj(layout._objs, x1, y1, filename, is_reverse, feuille, path, link_task_sc, classeur, page_num, current_time, compteur)

    
    classeur.save(path)

    print('Export enregistre sous : ')
    print(path)
    print('-'*20)
    subprocess.run(['explorer', os.path.realpath(path)])



app = QApplication.instance() 
if not app:
    app = QApplication(sys.argv)
    
fen = Fenetre()
fen.show()

