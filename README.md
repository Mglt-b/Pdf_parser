# Pdf_parser <br/>
Taches (synoptique) pdf vers Fieldwire <br/>
Pour Sogetrel <br/> <br/>
Pourrait-on étendre cela au plan de cablage ? <br/>


# Dépendances <br/>
pip install pdfminer <br/>
pip install xlwt <br/>
pip install PyQt5  <br/>


# Dev avec Python 3.8 <br/>


# Mise à jour <br/>
18/08/2020 : Ajout d'un fichier de configuration avec 40 variables editables, en fonction des nominations cables/ boites des projets  <br/>
18/08/2020 : Creation d'un executable v1.1  <br/><br/>
19/08/2020 : Migration des parametrages dans un fichier excel  <br/>
19/08/2020 : Creation d'un executable v1.2  <br/><br/>
14/10/2020 : Adresse reponsable modulable <br/>
14/10/2020 : Categories taches modulables <br/>
14/10/2020 : Prefixes des noms de taches modulables <br/>
14/10/2020 : Prise en compte numero PR pour le syane, parametrable <br/>
14/10/2020 : Documents multipages compatibles <br/>
14/10/2020 : Verification de présence du fichier de paramétrage <br/>
14/10/2020 : Adapatation du code pour le pôle GC SIEA <br/>
15/10/2020 : Ajout des listes de controles fixes, modulables via paramètres <br/>
15/10/2020 : Corrections taches associées pour adapter au projet SYANE </br>
15/10/2020 : Creation d'un executable v1.3, ajout de la release <br/>

# Comment utiliser <br/>
Télécharger la dernière release (executable .exe et excel de paramétrage)  <br/>
Placer les deux fichiers dans le même dossier   <br/>
############### </br>
Pour utiliser la fonction de taches associées, réaliser un export fieldwire de toutes les taches en csv.  <br/>
Supprimer toutes les colonnes sauf la colonne titre, qui comprend les noms des taches.  <br/>
Cette colonne doit etre en colonne A.  <br/>
Supprimer tous les champs vides de cette colonne.  <br/>
Sauvegarder en csv (point virgules), dans le dossier de l'executable, dans le sous dossier "\Tache_associee".  <br/>
Le nommage n'est pas important, mais il faut un seul csv a cet emplacement.  <br/>
############### </br>
Executer le fichier .exe <br/>

# TO DO <br/>
 <br/>
 
 Procédure utilisation liaison de tâches
 Problème de la marge ! <br/>
 Detection problème inversion X / Y <br/>

