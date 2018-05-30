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

from textract import process
from os import listdir, chdir,remove
from PyPDF2 import PdfFileReader,PdfFileMerger
from tkinter import Tk, StringVar,IntVar, Label, Radiobutton, Canvas, Button,Frame,Entry,Scrollbar,Listbox,Checkbutton
from tkinter.filedialog import askopenfilename,askdirectory
from functools import partial
from wordcloud import WordCloud, STOPWORDS
from xlwt import Workbook
import matplotlib.pyplot as plt


global listeFichier,langue,methode,numPageDeb,numPageFin,nomFichier,export
global label_deb,entry_deb,label_fin,entry_fin,label_PpP,radiobutton_PpP
global frame,fenetre,fond0

Coul_fond = '#f1c40f'
Coul_bouton1 = '#eb984e'
orange  = '#ffb74d'
Coul_fond_listeFic = '#f9e79f'
bleu    = '#aed6f1'
Coul_bouton2   = '#f5b7b1'
Coul_fond_option  = Coul_fond_listeFic #'#efebe9'
Coul_Demarrer = '#76d7c4'
Coul_Supp = '#eb984e'
Coul_fond_option2 = '#eb984e'

listeFichier = []



def LongueurChaine(chaine):
    """ Fonction qui renvoie la chaine de texte jusqu'à l'indice de fin des mots clés """
    if chaine[0] == ' ': chaine=chaine[1:]
    ind = 0
    while chaine[ind:ind+2] != '\n\n' and ind<len(chaine)-2:  # tant qu'il n'y a pas 2 retours à la ligne consécutifs
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

def bytes2Txt(nomFichier):
    """ Fonction qui convertie un fichier PDF en texte """
    text = process(nomFichier, encoding='utf-8',method='pdfminer')
    return bytes.decode(text)

def pdf2Txt(nomFichier,deb,fin,nbPage,essai):
    """ Fonction qui creee un fichier PDF temporaire pour trouver les mots clés dedans """
    with open(nomFichier,"rb") as f:
        pdf_reader = PdfFileReader(f)
        merger = PdfFileMerger()
        if essai==1: merger.append(fileobj = f, pages = (deb,fin))
        else :
            if essai == 2:
                merger.append(fileobj=f,pages=(1,deb))
                print(deb)
            else : merger.append(fileobj=f,pages=(fin,nbPage))
        nomFichier = "temp.pdf"
        with open(nomFichier, "wb") as output:
            merger.write(output)
    return bytes2Txt(nomFichier)

def SuppTemp():
    if 'temp.pdf' in listdir():
        remove('temp.pdf')

def tirets(liste):
    """ Rassemble les groupes de mots a l'aide de tirets et met tous les caracteres en minuscule"""
    listeret = []
    for mot in liste:
        mot = mot.lower()
        #mot = mot.replace('  ', ' ')
        mot = mot.replace(" ","_")
        mot = mot.replace("'","_")
        mot = mot.replace("’","_")
        listeret.append(mot)
    return listeret

def methodeIntervalle(l1):
    ListeMotCle = []
    for fichier in listeFichier:
        print("Fichier étudié : ", fichier)
        debP = int(numPageDeb.get())
        finP = int(numPageFin.get())
        nbPage = PdfFileReader(fichier).getNumPages()
        if not(0<=debP<=nbPage):                       # Vérif intervalle saisie
            debP=0
        if not(0<=finP<=nbPage):
            finP=nbPage
        if not(numPageDeb.get()<numPageFin.get()):
            debP=0
            finP=nbPage
        deb=-1;i = 1
        while i<4 and deb ==-1:
            text = pdf2Txt(fichier,debP,finP,nbPage,i).replace('  ',' ')
            deb = RecherDebMotCle(text)
            i+=1

        if deb != 1 and text[deb+l1+2:] != "":
            chaine = LongueurChaine(text[deb+l1+2:])
            chaine = chaine.replace('\n',' ').replace(', ',',')
            ListeMotCle = ListeMotCle + chaine.split(',')
    return ListeMotCle


def methodePageParPage (l1):
    ListeMotCle=[]
    for fichier in listeFichier:
        print("Fichier étudié : ", fichier)
        nbPage = PdfFileReader(fichier).getNumPages()
        i=0;deb=-1
        while i<nbPage and deb==-1:                                             # tq l'on a pas trouvé les mots clés
            text = pdf2Txt(fichier,i,i+1,nbPage,2).replace('  ',' ')
            deb = RecherDebMotCle(text)
            i+=1
        if deb != 1 and text[deb+l1+2:] != "":
            chaine = LongueurChaine(text[deb+l1+2:])
            chaine = chaine.replace('\n',' ').replace(', ',',')
            ListeMotCle = ListeMotCle + chaine.split(',')
    return ListeMotCle

def AddFichierCVS(listeMot):
    book = Workbook()                                                           # création
    feuil1 = book.add_sheet('feuille 1')                                        # création de la feuille 1
    feuil1.write(0,0,'Keywords')                                                # ajout des en-têtes
    feuil1.write(0,1,'Nombre d\' occurence')

    print(listeMot)
    print(listeMot[0])

    for i,mot in enumerate(listeMot):
        ligne = feuil1.row(i+1)
        ligne.write(0,mot[0])
        ligne.write(1,mot[1])

    feuil1.col(0).width = 10000                                                 # ajustement éventuel de la largeur d'une colonne
    nom = nomFichier.get()
    if len(nom)>3 and nom[len(nom)-4:] != '.xls':
        nom = nom+'.xls'
    book.save(nom);                                                             # création matérielle du fichier résultant


def main():
    """ Fonction principale, trouve les mots clés et genere le nuage de mot affiche avec matplotlib """
    global listeFichier
    if len(listeFichier)>0:
        l1 = 9
        if langue.get() == 'anglais': l1 = 7
        print(listeFichier)
        if methode.get() == 'Intervalle':
            ListeMotCle = methodeIntervalle(l1)
        else :                                                                  # Méthode intervalle
            ListeMotCle = methodePageParPage(l1)
        SuppTemp()                                                              # Suppression fichier temporaire
        ListeMotCle = tirets(ListeMotCle)
        ListeMotCleAffichage = remake(ListeMotCle)
        if export.get()!=0:
            AddFichierCVS(ListeMotCleAffichage)
        AffichageMotCle(ListeMotCleAffichage)
        #Creation chaine des mots (avec repetitions)
        text = ""
        for mot in ListeMotCle:
        	text = text + " " + mot
        # génération du nuage de mots
        wordcloud = WordCloud().generate(text)
        # lower max_font_size
        wordcloud = WordCloud(max_font_size=40).generate(text)
        plt.figure()
        plt.imshow(wordcloud, interpolation="bilinear")
        plt.axis("off")
        plt.show()
        MAJAffFic()

############################################################################
#                       AFFICHAGE                                          #
############################################################################

SuppTemp()
fenetre = Tk()

methode = StringVar(fenetre, 'Intervalle')
langue = StringVar(fenetre, 'Français')
numPageDeb = StringVar(fenetre,2)
numPageFin = StringVar(fenetre,6)
nomFichier = StringVar(fenetre,"NomDefault")
export = IntVar(fenetre,0)


def update_export():
    if export.get()==0:
        label_NomFic.grid_forget()
        entry_export.grid_forget()
    else:
        entry_export.grid(      row=7, column=3,sticky="we",padx=7)
        label_NomFic.grid(      row=7, column=2,sticky="w",padx=25)
    fenetre.update()

def update_methode(label):
    """ Met à jour la méthode """
    #methode = var.get()
    #label.config(text='Méthode : ' + methode)
    MAJaffichage_methode()

def MAJaffichage_methode():
    if methode.get() == 'Intervalle':
        label_deb.grid(row=11, column=2,sticky="w",padx = 40)
        entry_deb.grid(row=11, column=3,sticky="w",padx = 10)
        label_fin.grid(row=12,column=2,sticky="w",padx = 40)
        entry_fin.grid(row=12,column=3,sticky="w",padx = 10)
    else :
        label_deb.grid_forget()
        entry_deb.grid_forget()
        label_fin.grid_forget()
        entry_fin.grid_forget()
    fenetre.update()

def Path2Name(liste):
    newList = []
    for path in liste:
        newList.append(path.split('/')[-1])
    return newList

def MAJAffFic():
    global mylist
    frame = Frame(fenetre)
    frame.grid(row=3,column=1,rowspan=8,sticky = "sw",padx = 20,pady=20)
    scrollbar = Scrollbar(frame)
    scrollbar.pack(side='right',fill='y')
    mylist = Listbox(frame, yscrollcommand = scrollbar.set,width = 25,height=18)
    for nomFichier in Path2Name(listeFichier):
        mylist.insert('end', nomFichier)
    mylist.pack(expand=True,side='top', fill='y')
    scrollbar.config(command=mylist.yview )
    if len(listeFichier)>0:
        Button(fenetre, text="   Supprimer   ",command=Supp,bg=Coul_Supp,borderwidth=5,activebackground=Coul_fond_listeFic).grid(row = 7,column=1,sticky="e",padx=15)

def MAJListFichier():
    new = askopenfilename()
    if len(new)>4:
        if not(new in listeFichier) and new[len(new)-4:] == ".pdf":
            listeFichier.append(new)
            MAJAffFic()

def MAJListDossier():
    chemin = askdirectory()
    listeNew = listdir(chemin)
    for i,fic in enumerate(listeNew):
        listeNew[i]=chemin+'/'+fic
    modif = 0
    for newFichier in listeNew:
        if len(newFichier)>4:
            if not(newFichier in listeFichier) and newFichier[len(newFichier)-4:] == ".pdf":
                modif = 1
                listeFichier.append(newFichier)
    if modif == 1:
        MAJAffFic()

def Supp():
    global mylist
    temp = mylist.curselection()
    if temp == (): temp = 'end'
    fich = mylist.get(temp)
    fich = listeFichier[Path2Name(listeFichier).index(fich)]
    listeFichier.remove(fich)
    MAJAffFic()


fenetre.title("Synthèse PDF")

Canvas(fenetre,bg=Coul_fond,borderwidth=0,highlightthickness=0).grid(row=1,column=1,rowspan=13,columnspan=4,sticky='nswe')
Canvas(fenetre,bg=Coul_fond_listeFic).grid(row = 2,column = 1,rowspan=11,padx=5,sticky="nswe")
Canvas(fenetre,bg=Coul_fond_option, width=220).grid(row=2,column = 2,columnspan = 2,rowspan=11,padx=5,sticky="nswe")
Canvas(fenetre,bg=Coul_fond_option,highlightthickness=0,width=5,height=10).grid(row=9,column=2,padx=10,sticky="w")
Canvas(fenetre,bg=Coul_fond_option,highlightthickness=0,width=5,height=10).grid(row=10,column=2,padx=10,sticky="w")

Label(fenetre, text="Welcome !",font=("Times","24","bold"),bg=Coul_fond).grid(row=1,column=1,columnspan=3)
Label(fenetre, text="Options",font=("Times","20"),bg=Coul_fond_option).grid(row = 2,column=2,columnspan=2, padx = 5, pady = 5)
Label(fenetre, text="Liste fichiers",font=("Times","17",),bg=Coul_fond_listeFic).grid(row = 2,column=1, sticky="nw",padx = 15, pady = 10)

Button(fenetre, text="Ajouter fichiers",command = MAJListFichier,bg=Coul_bouton1,borderwidth=5,activebackground=Coul_fond_listeFic).grid(row=5,column=1,sticky="e",padx=15)
Button(fenetre, text="Ajouter dossier",command = MAJListDossier,bg=Coul_bouton1,borderwidth=5,activebackground=Coul_fond_listeFic).grid(row=6,column=1,sticky="e",padx=15)
Button(fenetre, text="DEMARRER",command = main,bg=Coul_Demarrer,borderwidth=5,activebackground="#d1f2eb",relief="raised").grid(row=13,column=1,columnspan=3,sticky='s',pady=5)


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


label_langue = Label(fenetre, text='Langue',font=("Times","12","bold"),bg=Coul_fond_option)
label_Fr = Label(fenetre, text='Français',bg=Coul_fond_option)
label_An = Label(fenetre, text='Anglais ',bg=Coul_fond_option)
label_Export = Label(fenetre, text = 'Exporter',font=("Times","12","bold"),bg=Coul_fond_option)
label_NomFic = Label(fenetre, text = 'Nom fichier',bg=Coul_fond_option)

label_methode = Label(fenetre, text='Méthode',font=("Times","12","bold"),bg=Coul_fond_option)
label_In = Label(fenetre, text='Intervalle    ',bg=Coul_fond_option)
label_PpP = Label(fenetre, text='Page par page',bg=Coul_fond_option)
label_fin = Label(fenetre, text='Fin : ',bg=Coul_fond_option)
label_deb = Label(fenetre, text='Début : ',bg=Coul_fond_option)

checkbutton_export = Checkbutton(fenetre,variable=export,bg=Coul_fond_option2,
                                 command=partial(update_export))

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
# entry_export.grid(      row=7, column=3,sticky="we",padx=7)
# label_NomFic.grid(      row=7, column=2,sticky="w",padx=25)

label_methode.grid(     row=8, column=2,columnspan=2,sticky="sw",padx=10)
label_PpP.grid(         row=9, column=2,sticky="w",padx=25)
radiobutton_PpP.grid(   row=9, column=3,padx = 10)
label_In.grid(          row=10, column=2,sticky="w",padx = 25)

radiobutton_In.grid(    row=10, column=3,padx = 10)
label_deb.grid(         row=11, column=2,sticky="w",padx=40)
entry_deb.grid(         row=11, column=3,padx = 10)
label_fin.grid(         row=12,column=2,sticky="w",padx=40)
entry_fin.grid(         row=12,column=3,padx = 10)


fenetre.mainloop()
