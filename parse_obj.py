import pdfminer
from recherche_csv import recherche_csv
import subprocess
import os
import xlrd
import sys, os
import time




def parse_obj(lt_objs, x1, y1, filename, is_reverse, feuille, path, link_task_sc, classeur, page_num, current_time, compteur):

    # Réouverture du classeur
    pathname = os.path.dirname(sys.argv[0])
    if os.path.isfile(pathname + r'\parametrages.xlsx') == False:
        print('Fichier de paramétrage absent, FIN')
        input('Fin')
        exit()
    classeur20 = xlrd.open_workbook(pathname + r'\parametrages.xlsx')

    

        
    # Récupération du nom de toutes les feuilles sous forme de liste
    nom_des_feuilles1 = classeur20.sheet_names()
        
    # Récupération de la première feuille
    feuille20 = classeur20.sheet_by_name(nom_des_feuilles1[0])

    ## PARTIE BOITE (acrire 'Variable Non utilisée' si la variable n'est pas utilisée)

    #Nombre d'underscores dans le nommage d'une boite
    # Mes boites ne contiennent pas:

    # Préfixes boites
    boite_prefix = format(feuille20.cell_value(0, 12))
    # Préfixes cables
    cable_prefix = format(feuille20.cell_value(1, 12))
    # mail_resp
    mail_resp = format(feuille20.cell_value(2, 12))
    # categorie boite
    cat_boite = format(feuille20.cell_value(3, 12))
    # categorie cable
    cat_cable = format(feuille20.cell_value(4, 12))
    # detection PR pour syane
    detect_pr = format(feuille20.cell_value(6, 12))
    # prefixes plans de boites
    prefixe_pdb = format(feuille20.cell_value(7, 12))
    # document multipage
    is_multipage = format(feuille20.cell_value(10, 12))
    # liste controle boite
    list_control_boite = format(feuille20.cell_value(12, 12))
    # liste controle cable
    list_control_cable = format(feuille20.cell_value(13, 12))



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

    is_reverse = 'no'
    xx = x1
    yy = y1

    #on loop pour savoir si un besoin de reverse est necessaire
    for obj in lt_objs:
        if isinstance(obj, pdfminer.layout.LTTextBoxHorizontal):
            x_percent = float(((float(obj.bbox[0])*100)/float(x1)))
            y_percent = float(100-(float(obj.bbox[1])*100)/float(y1))
            

                                                                        
            if ((x_percent >= 0) and (x_percent <= 100)) and ((y_percent >= 0) and (y_percent <= 100)):
                is_reverse = 'no'
            else :
                x1 = yy
                y1 = xx
                #is_reverse = 'y1'


    #on loop pour trouver le num PR
    num_pr = ''
    for obj in lt_objs:
        
        if isinstance(obj, pdfminer.layout.LTTextBoxHorizontal):
            le_texte3 = obj.get_text()
            if 'PR' in str(le_texte3):
                if str(detect_pr) == 'oui':
                    if str(le_texte3)[2].isnumeric() == True:
                        if not str('DemandeGC') in str(le_texte3):
                            num_pr = str(le_texte3)
    if num_pr == '':
        num_pr = 'ERREUR'


    #print('Start second loop')
    for obj in lt_objs:

        # if it's a textbox, print text and location
        if isinstance(obj, pdfminer.layout.LTTextBox) or isinstance(obj, pdfminer.layout.LTTextLine) or isinstance(obj, pdfminer.layout.LTFigure):

            try :
                    
                t_clean3 = obj.get_text()
                #clean du texte
                t_clean2 = str(t_clean3).replace('\n', '|')
                t_clean1 = str(t_clean2).replace('  ', '')
                #t_clean = str(t_clean1).replace(' ', '|')
                t_clean_0 = []
                t_clean_0.append(t_clean1)
                t_clean_b0 = t_clean_0[0].split('|')

                for t_clean_b in t_clean_b0:
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
                                                                    
                                                                
                                                                if is_reverse == 'no':
                                                                    x_percent = str(((float(obj.bbox[0])*100)/float(x1)))
                                                                    y_percent = str(100-(float(obj.bbox[1])*100)/float(y1))
                                                                if is_reverse == 'y1':
                                                                    x_percent = str(((float(obj.bbox[1])*100)/float(x1)))
                                                                    y_percent = str(100-(float(obj.bbox[0])*100)/float(y1))


                                                                        
                                                                #on nourris l'export excel
                                                                feuille.write(compteur, 0,  str(boite_prefix) + str(t_clean_b))
                                                                feuille.write(compteur, 1, "Priorité 1")
                                                                feuille.write(compteur, 2, str(cat_boite))
                                                                feuille.write(compteur, 3, str(mail_resp))
                                                                if is_multipage == 'oui':
                                                                    feuille.write(compteur, 6, str(filename) + str('_') + str(page_num))
                                                                else :
                                                                    feuille.write(compteur, 6, str(filename))
                                                                feuille.write(compteur, 7, str(x_percent))
                                                                feuille.write(compteur, 8, str(y_percent))
                                                                #associe les taches syno/cablage
                                                                value_desired = str(t_clean_b)
                                                                type_tache = 'boite'
                                                                tache_assoc_b = recherche_csv(value_desired, type_tache)
                                                                if tache_assoc_b is not None:
                                                                    feuille.write(compteur, 13, str(tache_assoc_b))
                                                                #associe les plans de boites
                                                                if not str(list_control_boite) == '':
                                                                    feuille.write(compteur, 14, str(list_control_boite))

                                                                if str(prefixe_pdb) == 'oui':
                                                                    feuille.write(compteur, 15, str(num_pr) + str('_') + str(value_desired) + str('.pdf'))
                                                                else:
                                                                    feuille.write(compteur, 15, str('_') + str(value_desired) + str('.pdf') )
                                                                

                    t_clean = t_clean_b
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

                                                                if is_reverse == 'no':
                                                                    x_percent = str((((float(obj.bbox[0]) + (float(obj.bbox[2])))/2)*100)/float(x1))
                                                                    y_percent = str(100 - (float(obj.bbox[1])*100)/float(y1))
                                                                if is_reverse == 'y1':
                                                                    x_percent = str((((float(obj.bbox[1]) + (float(obj.bbox[3])))/2)*100)/float(x1))
                                                                    y_percent = str(100 - (float(obj.bbox[0])*100)/float(y1))                    


                                                                #on nourris l'export excel
                                                                feuille.write(compteur, 0, str(cable_prefix) + str(t_clean_c))
                                                                feuille.write(compteur, 1, "Priorité 1")
                                                                feuille.write(compteur, 2, str(cat_cable))
                                                                feuille.write(compteur, 3, str(mail_resp))
                                                                if is_multipage == 'oui':
                                                                    feuille.write(compteur, 6, str(filename) + str('_') + str(page_num))
                                                                else:
                                                                    feuille.write(compteur, 6, str(filename))
                                                                feuille.write(compteur, 7, str(x_percent))
                                                                feuille.write(compteur, 8, str(y_percent))
                                                                if not str(list_control_cable) == '':
                                                                    feuille.write(compteur, 14, str(list_control_cable))
            except Exception as e:
                print(e)
    print("Fin de l'analise de la feuille")
    return compteur
    #if is_multipage == 'oui':
    #    path = r"C:\Export_Syno_Fieldwire\export_" + str(filename) + str(page_num) + str('-') + current_time + ".xls"

