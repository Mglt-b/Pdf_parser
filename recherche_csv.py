import glob
import csv
import os
import sys

def recherche_csv(value_desired, type_tache):
    #on essate de recuperer le fichier texte le plus recent
    newest_f = max(glob.iglob(r'C:\Export_Syno_Fieldwire\Tache_associee\*.[c][s][v]'), key=os.path.getctime)
    #on lis le fichier
    with open(newest_f, "rt") as originFile:
        try:
            for row in csv.reader(originFile, delimiter=';', quoting=csv.QUOTE_NONE):
                if (str('[')+str(value_desired)+str(']')) in str(row[1]):
                    #travail si la tache est une boite
                    if type_tache == 'boite':
                        print('yes ', row[1])
                        return row[1]
        except csv.Error as e:
            print('file {}, row {}: {}'.format(newest_f, reader.line_num, e))
        
    originFile.close()