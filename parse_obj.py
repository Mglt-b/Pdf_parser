import pdfminer
from recherche_csv import recherche_csv
import subprocess
import os
import xlrd
import sys, os




def parse_obj(lt_objs, x1, y1, filename, is_reverse, feuille, feuille2, path, link_task_sc, classeur):

    # Réouverture du classeur
    pathname = os.path.dirname(sys.argv[0])    
    classeur20 = xlrd.open_workbook(pathname + r'\parametrages.xlsx')
           
    

        
    # Récupération du nom de toutes les feuilles sous forme de liste
    nom_des_feuilles1 = classeur20.sheet_names()
        
    # Récupération de la première feuille
    feuille20 = classeur20.sheet_by_name(nom_des_feuilles1[0])

    ## PARTIE BOITE (acrire 'Variable Non utilisée' si la variable n'est pas utilisée)

    #Nombre d'underscores dans le nommage d'une boite
    #Par exemple, on renseigne 2 pour BPE_01266_00001
    nb_underscores_boites = format(feuille20.cell_value(20, 2))

    # Mes boites ne contiennent pas:

    # ET
    boite_var1_not_and = format(feuille20.cell_value(10, 2))
    # ET
    boite_var2_not_and = format(feuille20.cell_value(11, 2))
    # ET
    boite_var3_not_and = format(feuille20.cell_value(12, 2))
    # ET
    boite_var4_not_and = format(feuille20.cell_value(13, 2))
    # ET
    boite_var5_not_and = format(feuille20.cell_value(14, 2))
    # ET
    boite_var6_not_and = format(feuille20.cell_value(15, 2))
    # ET
    boite_var7_not_and = format(feuille20.cell_value(16, 2))
    # ET
    boite_var8_not_and = format(feuille20.cell_value(17, 2))
    # ET
    boite_var9_not_and = format(feuille20.cell_value(18, 2))
    # ET
    boite_var10_not_and = format(feuille20.cell_value(19, 2))

    # ET Mes boites contiennent :

    # ou
    boite_var1_or = format(feuille20.cell_value(0, 2))
    # ou
    boite_var2_or = format(feuille20.cell_value(1, 2))
    # ou
    boite_var3_or = format(feuille20.cell_value(2, 2))
    # ou
    boite_var4_or = format(feuille20.cell_value(3, 2))
    # ou
    boite_var5_or = format(feuille20.cell_value(4, 2))
    # ou
    boite_var6_or = format(feuille20.cell_value(5, 2))
    # ou
    boite_var7_or = format(feuille20.cell_value(6, 2))
    # ou
    boite_var8_or = format(feuille20.cell_value(7, 2))
    # ou
    boite_var9_or = format(feuille20.cell_value(8, 2))
    # ou
    boite_var10_or = format(feuille20.cell_value(9, 2))


    ## PARTIE CABLES (laisser variable vide si non utilisée°

    #Nombre d'underscores dans le nommage d'un cable
    #Par exemple, on renseigne 5 pour BPE_01266_00001_BPE_01266_00002
    nb_underscores_cables = format(feuille20.cell_value(20, 6))
    #ET

    #Mes cables contiennent :

    # ou
    cables_var1_or = format(feuille20.cell_value(0, 6))
    # ou
    cables_var2_or = format(feuille20.cell_value(1, 6))
    # ou
    cables_var3_or = format(feuille20.cell_value(2, 6))
    # ou
    cables_var4_or = format(feuille20.cell_value(3, 6))
    # ou
    cables_var5_or = format(feuille20.cell_value(4, 6))
    # ou
    cables_var6_or = format(feuille20.cell_value(5, 6))
    # ou
    cables_var7_or = format(feuille20.cell_value(6, 6))
    # ou
    cables_var8_or = format(feuille20.cell_value(7, 6))
    # ou
    cables_var9_or = format(feuille20.cell_value(8, 6))
    # ou
    cables_var10_or = format(feuille20.cell_value(9, 6))

    #Mes cables ne contiennent pas

    # ET
    cables_var1_not_and = format(feuille20.cell_value(10, 6))
    # ET
    cables_var2_not_and = format(feuille20.cell_value(11, 6))
    # ET
    cables_var3_not_and = format(feuille20.cell_value(12, 6))
    # ET
    cables_var4_not_and = format(feuille20.cell_value(13, 6))
    # ET
    cables_var5_not_and = format(feuille20.cell_value(14, 6))
    # ET
    cables_var6_not_and = format(feuille20.cell_value(15, 6))
    # ET
    cables_var7_not_and = format(feuille20.cell_value(16, 6))
    # ET
    cables_var8_not_and = format(feuille20.cell_value(17, 6))
    # ET
    cables_var9_not_and = format(feuille20.cell_value(18, 6))
    # ET
    cables_var10_not_and = format(feuille20.cell_value(19, 6))

    compteur = 0
    compteur_f2 = 0
    # loop over the object list
    print('Start loop')
    for obj in lt_objs:

        # if it's a textbox, print text and location
        if isinstance(obj, pdfminer.layout.LTTextBoxHorizontal):

            t_clean3 = obj.get_text()
            if 'FO' in str(t_clean3) and 'm / ' in str(t_clean3):

            
                compteur_f2 = compteur_f2 + 1
                feuille2.write(compteur_f2, 0, str(t_clean3))

            #clean du texte
            t_clean2 = str(t_clean3).replace('\n', '|')
            t_clean1 = str(t_clean2).replace('  ', '')
            t_clean = str(t_clean1).replace(' ', '|')
            
            #on detecte si c'est une boite, si oui, chaine de caractere avec 2 underscores
            nb_underscores_boites = int(str(nb_underscores_boites).replace('"',''))

            if t_clean.count("_") == nb_underscores_boites:

                #on supprime tout ce qu'il y a après un retour a la ligne dans le texte qui semble etre une boite
                if t_clean.split('|')[0].count('_') == nb_underscores_boites:
                    t_clean_b = t_clean.split('|')[0]
                elif t_clean.split('|')[1].count('_') == nb_underscores_boites:
                    t_clean_b = t_clean.split('|')[1]
                elif t_clean.split('|')[2].count('_') == nb_underscores_boites:
                    t_clean_b = t_clean.split('|')[2]
                elif t_clean.split('|')[3].count('_') == nb_underscores_boites:
                    t_clean_b = t_clean.split('|')[3]
                else:
                    t_clean_b = ''

                #essaye de recuperer uniquement les boites
                if not boite_var1_not_and in str(t_clean_b):
                    if not boite_var2_not_and in str(t_clean_b):
                        if not boite_var3_not_and in str(t_clean_b):
                            if not boite_var4_not_and in str(t_clean_b):
                                if not boite_var5_not_and in str(t_clean_b):
                                    if not boite_var6_not_and in str(t_clean_b):
                                        if not boite_var7_not_and in str(t_clean_b):
                                            if not boite_var8_not_and in str(t_clean_b):
                                                if not boite_var9_not_and in str(t_clean_b):
                                                    if not boite_var10_not_and in str(t_clean_b):                                
                                                        if boite_var1_or in str(t_clean_b) or boite_var2_or in str(t_clean_b) or boite_var3_or in str(t_clean_b) or boite_var4_or in str(t_clean_b) or boite_var5_or in str(t_clean_b) or boite_var6_or in str(t_clean_b) or boite_var7_or in str(t_clean_b) or boite_var8_or in str(t_clean_b) or boite_var9_or in str(t_clean_b) or boite_var10_or in str(t_clean_b):
                                                            compteur = compteur + 1
                                                                
                                                            x_percent = str(((float(obj.bbox[0])*100)/float(x1)))
                                                            y_percent = str(100-(float(obj.bbox[1])*100)/float(y1))
                                                                
                                                            if str(is_reverse) == 'y1':
                                                                x_percent = str(((float(obj.bbox[1])*100)/float(x1)))
                                                                y_percent = str(100-(float(obj.bbox[0])*100)/float(y1))
                                                                    
                                                            #on nourris l'export excel
                                                            feuille.write(compteur, 0, "[syno_b] " + str(t_clean_b))
                                                            feuille.write(compteur, 1, "1")
                                                            feuille.write(compteur, 2, "Tache_syno")
                                                            feuille.write(compteur, 3, str(os.getlogin()) + "@sogetrel.fr")
                                                            feuille.write(compteur, 6, str(filename))
                                                            feuille.write(compteur, 7, str(x_percent))
                                                            feuille.write(compteur, 8, str(y_percent))

                                                            if link_task_sc == 'o':
                                                                value_desired = str(t_clean_b)
                                                                type_tache = 'boite'
                                                                tache_assoc_b = recherche_csv(value_desired, type_tache)
                                                                if tache_assoc_b is not None:
                                                                    feuille.write(compteur, 13, str(tache_assoc_b))

            if t_clean.count("_") == int(str(nb_underscores_cables).replace('"','')):
                if not cables_var1_not_and in str(t_clean):
                    if not cables_var2_not_and in str(t_clean):
                        if not cables_var3_not_and in str(t_clean):
                            if not cables_var4_not_and in str(t_clean):
                                if not cables_var5_not_and in str(t_clean):
                                    if not cables_var6_not_and in str(t_clean):
                                        if not cables_var7_not_and in str(t_clean):
                                            if not cables_var8_not_and in str(t_clean):
                                                if not cables_var9_not_and in str(t_clean):
                                                    if not cables_var10_not_and in str(t_clean):                                
                                                        if cables_var1_or in str(t_clean) or cables_var2_or in str(t_clean) or cables_var3_or in str(t_clean) or cables_var4_or in str(t_clean) or cables_var5_or in str(t_clean) or cables_var6_or in str(t_clean) or cables_var7_or in str(t_clean) or cables_var8_or in str(t_clean) or cables_var9_or in str(t_clean) or cables_var10_or in str(t_clean):
                                                            
                                                            #on supprime tout ce qu'il y a après un retour a la ligne 
                                                            t_clean_c = t_clean.split('|')[0]
                                                            compteur = compteur + 1
                                                            x_percent = str((((float(obj.bbox[0]) + (float(obj.bbox[2])))/2)*100)/float(x1))
                                                            y_percent = str(100 - (float(obj.bbox[1])*100)/float(y1))
                                                            if str(is_reverse) == 'y1': 
                                                                x_percent = str((((float(obj.bbox[1]) + (float(obj.bbox[3])))/2)*100)/float(x1))
                                                                y_percent = str(100 - (float(obj.bbox[0])*100)/float(y1))                    


                                                            #on nourris l'export excel
                                                            feuille.write(compteur, 0, "[syno_c] " + str(t_clean_c))
                                                            feuille.write(compteur, 1, "1")
                                                            feuille.write(compteur, 2, "Tache_syno")
                                                            feuille.write(compteur, 3, str(os.getlogin()) + "@sogetrel.fr")
                                                            feuille.write(compteur, 6, str(filename))
                                                            feuille.write(compteur, 7, str(x_percent))
                                                            feuille.write(compteur, 8, str(y_percent))

    classeur.save(path)
    print('Export enregistre sous : ')
    print(path)
    subprocess.run(['explorer', os.path.realpath(path)])
