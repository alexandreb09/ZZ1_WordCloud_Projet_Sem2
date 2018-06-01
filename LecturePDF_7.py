#!/usr/bin/env python
################################################################################
#				               Projet ZZ1					                   #
#			        Bort Alexandre & Chomel Valentin			               #
#										                                       #
#	Projet TI - Technology Intelligence					                       #
#	Creation d'un nuage de mots cles a partir de rapports de stages		       #
#	ISIMA - 2018								                               #
#										                                       #
################################################################################

from functools import partial
from glob import glob
from os import listdir, path, remove, getcwd
from tkinter import (Button, Canvas, Checkbutton, Entry, Frame, IntVar, Label,
                     Listbox, Radiobutton, Scrollbar, StringVar, Tk)
from tkinter.filedialog import askdirectory, askopenfilename


from matplotlib.pyplot import imshow, axis,show
from numpy import array
from PIL import Image
from xlwt import Workbook

from PyPDF2 import PdfFileMerger, PdfFileReader
from textract import process
from wordcloud import STOPWORDS, ImageColorGenerator, WordCloud



global listeFichier,lionangue,methode,numPageDeb,numPageFin,nomFichier,dir_sauvegarde,dir_sauvegarde_display
global label_deb,entry_deb,label_fin,entry_fin,label_PpP,radiobutton_PpP
global frame,fenetre,fond0

Coul_fond = '#f1c40f'
Coul_bouton1 = '#eb984e'
Coul_fond_listeFic = '#f9e79f'
Coul_fond_option  = Coul_fond_listeFic
Coul_Demarrer = '#76d7c4'
Coul_Supp = '#eb984e'
Coul_fond_option2 = '#eb984e'

listeFichier = []
# d = path.dirname(__file__)



def LongueurChaine(chaine):
    """ Fonction qui renvoie la chaine de texte jusqu'à l'indice de fin des mots clés """
    if chaine[0] == ' ': chaine=chaine[1:]
    ind = 0
    while chaine[ind:ind+2] != '\n\n' and ind<len(chaine)-2:                    # tant qu'il n'y a pas 2 retours à la ligne consécutifs
        ind+=1
    return chaine[:ind+1]

def AffichageMotCle(Liste):
    """ Fonction qui affiche une liste de mots cles """
    print("\nListe des mots clés :")
    for elt in Liste:
        print(elt)

def RecherDebMotCle(text):
    """ Fonction qui recherche dans un texte l'apparition de mots clés
        Renvoie l'indice du début du mot 'mot-clé' """
    text = text.lower()
    stop = ["mots-clés","mots-cles","mots clés","mots cles","mot-clés","mot-cle","mot clé"]
    if langue.get() == 'Anglais':
        stop = ["keywords","keyword",'key words',"key word"]
    deb = -1
    numMot = 0
    while numMot < len(stop) and deb == -1:
        deb=text.find(stop[numMot])
        numMot+=1
    return deb

def remake(liste):
    """ Fonction qui creee une liste avec les mots cles les plus frequents et leurs nombre d'apparitions """
    newList = []
    for mot in set(liste):
        newList.append([mot,liste.count(mot)])
    return newList

def bytes2Txt(file_name):
    """ Fonction qui convertie un fichier PDF en texte """
    text = process(file_name, encoding='utf-8',method='pdfminer')
    return bytes.decode(text)

def pdf2Txt(file_name,deb,fin,nbPage,essai):
    """ Fonction qui creee un fichier PDF temporaire pour trouver les mots clés dedans """
    with open(file_name,"rb") as f:                                             # Ouverture du fichier
        pdf_reader = PdfFileReader(f)                                           # lecture du pdf
        merger = PdfFileMerger()
        if essai==1: merger.append(fileobj = f, pages = (deb,fin))              # En fonction du code d'essai
        else :                                                                  # On récupère le morceau de pdf correspondant
            if essai == 2:                                                      # ...
                merger.append(fileobj=f,pages=(1,deb))
                print(deb)
            else : merger.append(fileobj=f,pages=(fin,nbPage))
        file_name = "temp.pdf"
        with open(file_name, "wb") as output:
            merger.write(output)
    return bytes2Txt(file_name)

def SuppTemp():
    """Supprime le fichier temporaire : temp.pdf"""
    if 'temp.pdf' in listdir():
        remove('temp.pdf')

def tirets(liste):
    """ Rassemble les groupes de mots a l'aide de tirets
        Met tous les caracteres en minuscule
        Supprime les espaces inutiles"""
    listeret = []
    for mot in liste:                                                           # Pour chaque mot :
        mot = mot.lower()                                                       # Passage en minuscule
        if mot[0]==' ': mot=mot[1:]                                             # Suppression du 1er espace s'il y a
        while mot[-1]==' ': mot=mot[:-1]                                        # Suppression des espaces de fin
        mot = mot.replace('  ', ' ')                                            # Remplacement des doubles espaces par des simples
        mot = mot.replace(" ","_")                                              # Remplacement dievrs ...
        mot = mot.replace("'","_")
        mot = mot.replace("’","_")
        listeret.append(mot)
    return listeret

def methodeIntervalle(l1):
    """ Recherche des mots clés dans l'intervalle
        Si non trouvé : recherche dans le reste du document (debut puis fin)"""
    ListeMotCle = []
    for fichier in listeFichier:                                                # pour chaque fichier :
        print("Fichier étudié : ", fichier)
        debP = int(numPageDeb.get())                                            # Début de l'intervalle étudié
        finP = int(numPageFin.get())                                            # Fin de l'intervalle étudié
        nbPage = PdfFileReader(fichier).getNumPages()                           # Nombre de pages du pdf
        if not(0<=debP<=nbPage):                                                # Vérif intervalle saisie
            debP=0
        if not(0<=finP<=nbPage):
            finP=nbPage
        if not(numPageDeb.get()<numPageFin.get()):
            debP=0
            finP=nbPage
        deb=-1;i = 1                                                            # i : code d'essaie
        while i<4 and deb ==-1:                                                 # Tant que l'on a pas trouvé les mots clés
            text = pdf2Txt(fichier,debP,finP,nbPage,i).replace('  ',' ')        # On étudie l'intervalle de texte selectionné
            deb = RecherDebMotCle(text)                                         # on recherche le début
            i+=1

        if deb != 1 and text[deb+l1+2:] != "":                                  # Si on a trouvé les mots clés
            chaine = LongueurChaine(text[deb+l1+2:])                            # Traitement sur les mots
            chaine = chaine.replace('\n',' ').replace(', ',',')
            ListeMotCle = ListeMotCle + chaine.split(',')
    return ListeMotCle


def methodePageParPage (l1):
    """Recherche des mots clés page par page (s'arrête dès que trouvé)"""
    ListeMotCle=[]
    for fichier in listeFichier:                                                # Pour chaque fichier
        nbPage = PdfFileReader(fichier).getNumPages()                           # Nombre de pages
        i=0;deb=-1                                                              # i : indice courant
        while i<nbPage and deb==-1:                                             # tq l'on a pas trouvé les mots clés
            text = pdf2Txt(fichier,i,i+1,nbPage,2).replace('  ',' ')            # On travaille sur la i-eme page
            deb = RecherDebMotCle(text)                                         # Recherche des mots clés
            i+=1
        if deb != 1 and text[deb+l1+2:] != "":                                  # Si mots clés trouvés
            chaine = LongueurChaine(text[deb+l1+2:])                            # Traitement sur les mots
            chaine = chaine.replace('\n',' ').replace(', ',',')
            ListeMotCle = ListeMotCle + chaine.split(',')
    return ListeMotCle

def AddFichierCVS(listeMot):
    """Enregistres la liste des mots clés dans un fihcier xls"""
    book = Workbook()                                                           # création
    feuil1 = book.add_sheet('feuille 1')                                        # création de la feuille 1
    feuil1.write(0,0,'Keywords')                                                # ajout des en-têtes
    feuil1.write(0,1,'Nombre d\' occurence')

    for i,mot in enumerate(listeMot):                                           # Ecriture sur chaque ligne
        ligne = feuil1.row(i+1)
        ligne.write(0,mot[0])                                                   # Du mot
        ligne.write(1,mot[1])                                                   # De son nombre d'occurences

    feuil1.col(0).width = 10000                                                 # ajustement éventuel de la largeur d'une colonne
    nom = nomFichier.get()                                                      # Récupération nom saisi
    if len(nom)>3 and nom[len(nom)-4:] != '.xls':                               # Vérification nom fichier
        nom = nom+'.xls'
    book.save(dir_sauvegarde+'/'+nom);                                          # sauvegarde du fichier xsl


def main():
    """ Fonction principale, trouve les mots clés et genere le nuage de mot affiche avec matplotlib """
    global listeFichier
    if len(listeFichier)>0:
        l1 = 9
        if langue.get() == 'anglais': l1 = 7                                    # Définition de la longueur moyenne en fonction de la langue
        print(listeFichier)
        if methode.get() == 'Intervalle':                                       # Méthode intervalle
            ListeMotCle = methodeIntervalle(l1)
        else :                                                                  # Méthode Page par page
            ListeMotCle = methodePageParPage(l1)
        SuppTemp()                                                              # Suppression fichier temporaire
        ListeMotCle = tirets(ListeMotCle)                                       # Traitement sur les mots de la liste
        ListeMotCleAffichage = remake(ListeMotCle)                              # Ajout du nombre d'occurence de chaque mot dans la liste
        if export.get()!=0:                                                     # Si case export cochée
            AddFichierCVS(ListeMotCleAffichage)                                 # Exportation vers fichier xls
        AffichageMotCle(ListeMotCleAffichage)

        # Creation chaine des mots (avec repetitions)
        text = ""
        for mot in ListeMotCle:
        	text = text + " " + mot
        stopwords = set(STOPWORDS)
        stopwords.add("said")

        if path.exists(getcwd()+'/fond_wordcloud.png'):                                     # Si une imag de fond existe
            alice_coloring = array(Image.open(path.join(d, "fond_wordcloud.png")))       # Lecture de l'image
            wc = WordCloud(background_color="white", max_words=2000, mask=alice_coloring,
                   stopwords=stopwords, max_font_size=400, random_state=42).generate(text)  # Création du nuage de mots à partir de l'image
            image_colors = ImageColorGenerator(alice_coloring)                              # create coloring from image
            imshow(wc.recolor(color_func=image_colors), interpolation="bilinear")       # Affichage
        else :
            wc = WordCloud(background_color="white", max_words=2000,stopwords=stopwords,
                            max_font_size=400, random_state=42).generate(text)              # Génération du nuage de mot de manière automatique
            imshow(wc, interpolation="bilinear")                                        # Affichage
        axis("off")                                                                     # Desactivation axes
        show()                                                                          # Affichage fenetre

        MAJAffFic()

############################################################################
#                     FONCTION GESTION AFFICHAGE                           #
############################################################################

SuppTemp()
fenetre = Tk()

#Initialisation des variables
methode = StringVar(fenetre, 'Intervalle')
langue = StringVar(fenetre, 'Français')
numPageDeb = StringVar(fenetre,2)
numPageFin = StringVar(fenetre,6)
nomFichier = StringVar(fenetre,"NomDefault")
export = IntVar(fenetre,0)
dir_sauvegarde = getcwd()                                                       # Chemin absolu de sauvegarde
dir_sauvegarde_display = StringVar(fenetre,dir_sauvegarde.split('/')[-1])       # Dossier du sauvegarde


def MAJ_export():
    """Met à jour l'affichage du bloc export """
    if export.get()==0:                                                         # Si export désactivé
        label_NomFic.grid_forget()                                              # Suppresssion des différents éléments ...
        entry_export.grid_forget()
        label_save.grid_forget()
        Boutton_parcourir.grid_forget()
        label_indice.grid_forget()
    else:                                                                       # Sinon
        entry_export.grid(      row=7, column=3, sticky="we",padx=7)            # Ajout des éléments dans la grille
        label_NomFic.grid(      row=7, column=2, sticky="w", padx=25)
        label_save.grid(        row=8, column=3)
        label_indice.grid(       row=8,column=2)
        Boutton_parcourir.grid( row=8, column=4, sticky="e", padx=7)
    fenetre.update()                                                            # Actualisation du contenu

def update_methode(label):
    """ Met à jour la méthode """
    #methode = var.get()                                                        # Afficher la méthode selectionée
    #label.config(text='Méthode : ' + methode)                                  # à coté de "Méthode : "
    MAJaffichage_methode()

def MAJaffichage_methode():
    if methode.get() == 'Intervalle':                                           # Si méthode intervalle selectionnée
        label_deb.grid(row=11, column=2,sticky="w",padx = 40)                   # Ajout des différents éléments du bloc méthode intervalle
        entry_deb.grid(row=11, column=3,sticky="w",padx = 10)
        label_fin.grid(row=12,column=2,sticky="w",padx = 40)
        entry_fin.grid(row=12,column=3,sticky="w",padx = 10)
    else :                                                                      # sinon
        label_deb.grid_forget()                                                 # Suppression de ces éléments
        entry_deb.grid_forget()
        label_fin.grid_forget()
        entry_fin.grid_forget()
    fenetre.update()                                                            # Actualisationdu de l'affichage

def Path2Name(liste):
    """Prend une liste de chemins et renvoie une liste de nom du fichier"""
    newList = []
    for path in liste:
        newList.append(path.split('/')[-1])
    return newList

def MAJAffFic():
    """ Met à jour l'affichage des fichiers selectionnés"""
    global mylist
    frame = Frame(fenetre)
    frame.grid(row=3,column=1,rowspan=8,sticky = "sw",padx = 20,pady=20)
    scrollbar = Scrollbar(frame)
    scrollbar.pack(side='right',fill='y')
    mylist = Listbox(frame, yscrollcommand = scrollbar.set,width = 25,height=18)
    for file_name in Path2Name(listeFichier):
        mylist.insert('end', file_name)
    mylist.pack(expand=True,side='top', fill='y')
    scrollbar.config(command=mylist.yview )
    if len(listeFichier)>0:
        Button(fenetre, text="   Supprimer   ",command=Supp,bg=Coul_Supp,borderwidth=5,activebackground=Coul_fond_listeFic).grid(row = 7,column=1,sticky="e",padx=15)

def MAJListFichier():
    """ Actualise la liste des fichiers """
    new = askopenfilename()                                                     # Ouvre une boite de dialogue pour choisir le fichier
    if len(new)>4:                                                              # Si le fichier fait plus de 4 caractères
        if not(new in listeFichier) and new[len(new)-4:] == ".pdf":             # Et qu'il n'est pas dans la liste
            listeFichier.append(new)                                            # il est ajouté
            MAJAffFic()                                                         # l'affichage est actualisé

def MAJListDossier():
    chemin = askdirectory(title="Selectionner un dossier")                      # Demande le répertoire à l'utilisateur
    listeNew = glob(chemin + '/**/*.pdf', recursive=True)                       # Renvoie tous les fichiers "*.pdf" du dossier et sous dossier
    listeNew = list(set(listeNew))
    modif = False                                                               # Booleen : si modification = True, sinon = False
    for newFichier in listeNew:                                                 # Pour chaque fichier :
        if not(newFichier in listeFichier):                                     # S'il n'est pas dans la liste des anciens fichiers
            modif = True                                                        # Il y a une modification
            listeFichier.append(newFichier)                                     # On ajoute le fichier à la liste des fichiers
    if modif :                                                                  # S'il y a eu une modif dans la liste des fichiers
        MAJAffFic()                                                             # Actualisation de l'affichage

def MAJSaveFolder():
    """ Met à jour les chemins de sauvegardes du fichier xls"""
    global dir_sauvegarde
    dir_sauvegarde = askdirectory(title="Selectionner un dossier de sauvegarde")
    dir_sauvegarde_display.set(dir_sauvegarde.split('/')[-1])
    MAJ_export()

def Supp():
    """ Supprime un fichier de la liste"""
    global mylist
    temp = mylist.curselection()                                                # récupère l'élément selectionné
    if temp == (): temp = 'end'                                                 # s'il y en a pas, le dernier est selectionné
    fich = mylist.get(temp)                                                     # Suppression de l'élément
    fich = listeFichier[Path2Name(listeFichier).index(fich)]
    listeFichier.remove(fich)
    MAJAffFic()




############################################################################
#                                AFFICHAGE                                 #
############################################################################


fenetre.title("Synthèse PDF")

Canvas(fenetre,bg=Coul_fond,borderwidth=0,highlightthickness=0).grid(row=1,column=1,rowspan=15,columnspan=4,sticky='nswe')
Canvas(fenetre,bg=Coul_fond_listeFic).grid(row = 2,column = 1,rowspan=13,padx=5,sticky="nswe")
Canvas(fenetre,bg=Coul_fond_option, width=220).grid(row=2,column = 2,columnspan = 3,rowspan=13,padx=5,sticky="nswe")
Canvas(fenetre,bg=Coul_fond_option,highlightthickness=0,width=5,height=10).grid(row=9,column=2,padx=10,sticky="w")
Canvas(fenetre,bg=Coul_fond_option,highlightthickness=0,width=5,height=10).grid(row=10,column=2,padx=10,sticky="w")

Label(fenetre, text="Welcome !",font=("Times","24","bold"),bg=Coul_fond).grid(row=1,column=1,columnspan=3)
Label(fenetre, text="Options",font=("Times","20"),bg=Coul_fond_option).grid(row = 2,column=2,columnspan=2, padx = 5, pady = 5)
Label(fenetre, text="Liste fichiers",font=("Times","17",),bg=Coul_fond_listeFic).grid(row = 2,column=1, sticky="nw",padx = 15, pady = 10)

Button(fenetre, text="Ajouter fichiers",command = MAJListFichier,bg=Coul_bouton1,borderwidth=5,activebackground=Coul_fond_listeFic).grid(row=5,column=1,sticky="e",padx=15)
Button(fenetre, text="Ajouter dossier",command = MAJListDossier,bg=Coul_bouton1,borderwidth=5,activebackground=Coul_fond_listeFic).grid(row=6,column=1,sticky="e",padx=15)
Button(fenetre, text="DEMARRER",command = main,bg=Coul_Demarrer,borderwidth=5,activebackground="#d1f2eb",relief="raised").grid(row=15,column=1,columnspan=3,sticky='s',pady=5)


fenetre.rowconfigure(1,weight=40)
fenetre.rowconfigure(2,weight=20,minsize=20)
fenetre.rowconfigure(3,weight=15,minsize=20)
fenetre.rowconfigure(4,weight=10,minsize=20)
fenetre.rowconfigure(5,weight=10,minsize=20)
fenetre.rowconfigure(6,weight=15,minsize=20)
fenetre.rowconfigure(7,weight=10,minsize=20)
fenetre.rowconfigure(8,weight=10,minsize=20)
fenetre.rowconfigure(9,weight=10,minsize=20)
fenetre.rowconfigure(10,weight=10,minsize=20)
fenetre.rowconfigure(11,weight=10,minsize=20)
fenetre.rowconfigure(12,weight=10,minsize=20)
fenetre.rowconfigure(13,weight=10,minsize=20)



Boutton_parcourir = Button(fenetre, text="...",command = MAJSaveFolder,bg=Coul_bouton1,borderwidth=5,activebackground="#d1f2eb",relief="raised")

label_langue = Label(fenetre, text='Langue',font=("Times","12","bold"),bg=Coul_fond_option)
label_Fr = Label(fenetre, text='Français',bg=Coul_fond_option)
label_An = Label(fenetre, text='Anglais ',bg=Coul_fond_option)
label_Export = Label(fenetre, text = 'Exporter',font=("Times","12","bold"),bg=Coul_fond_option)
label_NomFic = Label(fenetre, text = 'Nom fichier',bg=Coul_fond_option)

label_save = Label(fenetre,textvariable=dir_sauvegarde_display,bg=Coul_fond_option)
label_indice = Label(fenetre,text="Sauvegardé dans :",bg=Coul_fond_option)

label_methode = Label(fenetre, text='Méthode',font=("Times","12","bold"),bg=Coul_fond_option)
label_In = Label(fenetre, text='Intervalle    ',bg=Coul_fond_option)
label_PpP = Label(fenetre, text='Page par page',bg=Coul_fond_option)
label_fin = Label(fenetre, text='Fin : ',bg=Coul_fond_option)
label_deb = Label(fenetre, text='Début : ',bg=Coul_fond_option)

checkbutton_export = Checkbutton(fenetre,variable=export,bg=Coul_fond_option2,
                             command=partial(MAJ_export))

radiobutton_Fr = Radiobutton(fenetre, variable=langue, value='Français',bg=Coul_fond_option2)
radiobutton_An = Radiobutton(fenetre, variable=langue, value='Anglais',bg=Coul_fond_option2)
radiobutton_In = Radiobutton(fenetre, variable=methode, value='Intervalle',bg=Coul_fond_option2,
                            command=partial(update_methode,label_methode))
radiobutton_PpP= Radiobutton(fenetre, variable=methode, value='Page par page',bg=Coul_fond_option2,
                            command=partial(update_methode,label_methode))

entry_deb = Entry(fenetre, textvariable=numPageDeb,width=3)
entry_fin = Entry(fenetre, textvariable=numPageFin,width=3)
entry_export = Entry(fenetre,textvariable=nomFichier,width=15)


label_langue.grid(      row=3, column=2,columnspan=2,sticky="w",padx=10)
label_Fr.grid(          row=4, column=2,sticky="w",padx=25)
radiobutton_Fr.grid(    row=4, column=3,padx=10)
label_An.grid(          row=5, column=2,sticky="w",padx=25)
radiobutton_An.grid(    row=5, column=3,padx=10)
label_Export.grid(      row=6, column=2,sticky="sw",padx=10)
checkbutton_export.grid(row=6, column=3,sticky="s",pady=2)

label_methode.grid(     row=10, column=2,columnspan=2,sticky="sw",padx=10)
label_PpP.grid(         row=11, column=2,sticky="w",padx=25)
radiobutton_PpP.grid(   row=11, column=3,padx = 10)
label_In.grid(          row=12, column=2,sticky="w",padx = 25)

radiobutton_In.grid(    row=12, column=3,padx = 10)
label_deb.grid(         row=13, column=2,sticky="w",padx=40)
entry_deb.grid(         row=13, column=3,padx = 10)
label_fin.grid(         row=14,column=2,sticky="w",padx=40)
entry_fin.grid(         row=14,column=3,padx = 10)


fenetre.mainloop()
