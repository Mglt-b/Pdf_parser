import glob
import csv
import os
import sys

def recherche_csv(value_desired, type_tache):
    #on essate de recuperer le fichier texte le plus recent
    pathname = os.path.dirname(sys.argv[0])  
    newest_f = max(glob.iglob(pathname + r'\Tache_associee\*.[c][s][v]'), key=os.path.getctime)
    i = 0
    #on lis le fichier
    with open(newest_f, "rt") as originFile:
        try:
            for row in csv.reader(originFile, delimiter=';', quoting=csv.QUOTE_NONE):
                if (str('[')+str(value_desired)+str(']')) in str(row[0]):
                    i = i + 1
                    #travail si la tache est une boite
                    if type_tache == 'boite':
                        return row[0]
                else :
                    if (str(value_desired)) in str(row[0]):
                        i = i + 1
                        #travail si la tache est une boite
                        if type_tache == 'boite':
                            return row[0]                    
        except csv.Error as e:
            print('file {}, row {}: {}'.format(newest_f, str(i), e))
        
    originFile.close()