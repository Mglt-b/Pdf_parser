import pdfminer
from recherche_csv import recherche_csv
import subprocess
import os

def parse_obj(lt_objs, x1, y1, filename, is_reverse, feuille, feuille2, path, link_task_sc, classeur):

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
            if t_clean.count("_") == 2:
                #on supprime tout ce qu'il y a après un retour a la ligne dans le texte qui semble etre une boite
                if t_clean.split('|')[0].count('_') == 2:
                    t_clean_b = t_clean.split('|')[0]
                elif t_clean.split('|')[1].count('_') == 2:
                    t_clean_b = t_clean.split('|')[1]
                elif t_clean.split('|')[2].count('_') == 2:
                    t_clean_b = t_clean.split('|')[2]
                elif t_clean.split('|')[3].count('_') == 2:
                    t_clean_b = t_clean.split('|')[3]
                else:
                    t_clean_b = ''

                #essaye de recuperer uniquement les boites
                if not 'FO' in str(t_clean_b):
                    if not 'NRO' in str(t_clean_b):
                        if not '/' in str(t_clean_b):
                            if not 'POSE' in str(t_clean_b):
                                if not 'LOVE' in str(t_clean_b):
                                    #if len(str(t_clean_b)) < 20: [c'est trop random pour utiliser ça, on va compter les underscrore plutot]
                                    if 'BPESYA' in str(t_clean_b) or 'NRO' in str(t_clean_b) or 'BRP' in str(t_clean_b) or 'BRD' in str(t_clean_b) or 'BRP' in str(t_clean_b) or 'BPI' in str(t_clean_b) or 'BRI' in str(t_clean_b) or 'BRA' in str(t_clean_b) or 'BRF' in str(t_clean_b):
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

            if 'POSE' in str(t_clean) or 'CABSYA' in str(t_clean) or '/|IMMEUBLE' in str(t_clean) or '_BR' in str(t_clean):
                #on supprime tout ce qu'il y a après un retour a la ligne dans le texte qui semble etre une boite
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
