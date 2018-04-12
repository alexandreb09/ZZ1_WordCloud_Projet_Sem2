from textract import process
from os import listdir, chdir,remove
from PyPDF2 import PdfFileReader,PdfFileMerger
from tkinter import Tk, StringVar, Label, Radiobutton, Canvas, Button,Frame,Entry,Scrollbar,Listbox
from tkinter.filedialog import askopenfilename,askdirectory
from functools import partial

def LongueurChaine(chaine):                                   # Donne l'indice de fin des mots clés
    if chaine[0] == ' ': chaine=chaine[1:]
    ind = 0
    while chaine[ind:ind+2] != '\n\n' and ind<len(chaine)-2:  # tant qu'il n'y a pas 2 retours à la ligne consécutifs
        ind+=1
    return chaine[:ind+1]

def AffichageMotCle(Liste):
    print("\nListe des mots clés :")
    for elt in Liste:
        print(elt)

def RecherDebMotCle(text, langue='fr'):
    text = text.lower()
    stop = ["mots-clés","mots-cles","mots clés","mots cles"]
    if langue == 'an':
        stop = ["keywords","keyword",'key words']
    lg = len(stop[0])
    deb = -1
    numMot = 0
    while numMot < len(stop) and deb == -1:
        deb=text.find(stop[numMot])
        numMot+=1
    return deb


def convertPDF2Txt(nomFichier):
    text = process(nomFichier, encoding='utf-8',method='pdfminer')
    return bytes.decode(text)

def remake(liste):
    newList = []
    for mot in set(liste):
        newList.append([mot,liste.count(mot)])
    return newList

def creerFichierPdfTemp(nomFichier,numPageDeb,numPageFin):
    with open(nomFichier,"rb") as f:
        pdf_reader = PdfFileReader(f)
        if numPageFin < pdf_reader.getNumPages() and numPageDeb > 0 :
            merger = PdfFileMerger()
            merger.append(fileobj = f, pages = (numPageDeb,numPageFin))
            nomFichier = "temp.pdf"
            with open(nomFichier, "wb") as output:
                merger.write(output)
    return nomFichier

def SuppTemp():
    if 'temp.pdf' in listdir():
        remove('temp.pdf')


def main(listeFichier,langue='fr',numPageDeb=2,numPageFin=6):
    l1 = 9
    if langue == 'an': l1 = 7
    ListeMotCle = []
    for fichier in listeFichier:
        print("Fichier étudié : ", fichier)
        fichier = creerFichierPdfTemp(fichier,numPageDeb,numPageFin)
        text = convertPDF2Txt(fichier)
        text = text.replace('  ',' ')
        deb = RecherDebMotCle(text)
        if deb!= -1:
            chaine = LongueurChaine(text[deb+l1+2:])
            chaine = chaine.replace('\n',' ')                                       # suppresion des \n
            chaine = chaine.replace(', ',',')
            ListeMotCle = ListeMotCle + chaine.split(',')
    ListeMotCle = remake(ListeMotCle)
    SuppTemp()
    return(ListeMotCle)


############################################################################
#                       AFFICHAGE                                          #
############################################################################

global listeFichier,langue,methode,numPageDeb,numPageFin
global label_deb,entry_deb,label_fin,entry_fin,label_PpP,radiobutton_PpP,fenetre,fond0
global frame

vert0 = '#daf7a6'
vert1 = '#81c784'
orange = '#ffb74d'
jaune = '#f9e79f'
bleu = '#aed6f1'
rouge = '#f5b7b1'
marron = '#efebe9'
marron2='#bcaaa4'

SuppTemp()
listeFichier = []
fenetre = Tk()

methode = StringVar(fenetre, 'Intervalle')
langue = StringVar(fenetre, 'Français')
numPageDeb = StringVar(fenetre,2)
numPageFin = StringVar(fenetre,6)


def update_langue(label, var):
    """ Met à jour de la langue """
    text = var.get()
    if text == 'Intervalle': text+= '   '
    label.config(text='Langue : ' + text)

def update_methode(label, var):
    """ Met à jour de la langue """
    methode = var.get()
    label.config(text='Méthode : ' + methode)
    MAJaffichage(methode)

def MAJaffichage(methode):
    if methode == 'Intervalle':
        label_deb.grid(row=9,column=2)
        entry_deb.grid(row=9,column=3)
        label_fin.grid(row=10,column=2)
        entry_fin.grid(row=10,column=3)
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
    frame.grid(row=3,column=1,rowspan=8,sticky = "nw",padx = 20)
    scrollbar = Scrollbar(frame)
    scrollbar.pack(side='right',fill='y')
    mylist = Listbox(frame, yscrollcommand = scrollbar.set,width = 40,height=18)
    for nomFichier in Path2Name(listeFichier):
        mylist.insert('end', nomFichier)
    mylist.pack(expand=True,side='top', fill='y')
    scrollbar.config(command=mylist.yview )
    if len(listeFichier)>0:
        Button(fenetre, text="   Supprimer   ",command=Supp,bg=rouge).grid(row = 7,column=1,sticky="e",padx=15)

def MAJListFichier():
    new = askopenfilename().split('/')[-1]
    if not(new in listeFichier):
        listeFichier.append(new)
        MAJAffFic()

def MAJListDossier():
    listeNew = listdir(askdirectory())
    modif = 0
    for newFichier in Path2Name(listeNew):
        if not(newFichier in listeFichier):
            modif = 1
            listeFichier.append(newFichier)
    if modif ==1:
        MAJAffFic()

def Supp():
    global mylist
    temp = mylist.curselection()
    if temp == (): temp = 'end'
    fich = mylist.get(temp)
    listeFichier.remove(fich)
    MAJAffFic()


fenetre.title("Synthèse PDF")

Canvas(fenetre,bg=vert0,borderwidth=0,highlightthickness=0).grid(row=1,column=1,rowspan=12,columnspan=4,sticky='nswe')
Canvas(fenetre, width=550, height=350, bg=jaune).grid(row = 2,column = 1,rowspan=9,pady=5,padx = 10)
Canvas(fenetre, width=220,bg=marron).grid(row=2,column = 2,columnspan = 2,rowspan=9,pady=5,padx = 5,sticky="nswe")

Label(fenetre, text="Welcome !",font=("Times","24","bold"),bg=vert0).grid(row=1,column=1,columnspan=3)
Label(fenetre, text="Options",font=("Times","20","underline"),bg=marron).grid(row = 2,column=2,columnspan=2, padx = 5, pady = 5)
Label(fenetre, text="Liste fichiers : ",font=("Times","17",),bg=jaune).grid(row = 2,column=1, sticky="nw",padx = 15, pady = 10)

Button(fenetre, text="Ajouter fichiers",command = MAJListFichier,bg=vert1).grid(row=5,column=1,sticky="e",padx=15)
Button(fenetre, text="Ajouter dossier",command = MAJListDossier,bg=vert1).grid(row=6,column=1,sticky="e",padx=15)
Button(fenetre, text="DEMARRER",command = partial(main,listeFichier),bg=rouge).grid(row=12,column=1,sticky='s',pady=5)



label_langue = Label(fenetre, text='Langue : ' + langue.get() + '   ',bg=marron)
label_Fr = Label(fenetre, text='Français',bg=marron)
label_An = Label(fenetre, text='Anglais ',bg=marron)
label_methode = Label(fenetre, text='Méthode : ' + methode.get(),bg=marron)
label_In = Label(fenetre, text='Intervalle',bg=marron)
label_PpP = Label(fenetre, text='Page par page',bg=marron)
label_fin = Label(fenetre, text='Fin : ',bg=marron)
label_deb = Label(fenetre, text='Début : ',bg=marron)

radiobutton_Fr = Radiobutton(fenetre, variable=langue, value='Français',bg=marron2,
                            command=partial(update_langue,label_langue,langue))
radiobutton_An = Radiobutton(fenetre, variable=langue, value='Anglais',bg=marron2,
                            command=partial(update_langue,label_langue,langue))
radiobutton_In = Radiobutton(fenetre, variable=methode, value='Intervalle',bg=marron2,
                            command=partial(update_methode,label_methode,methode))
radiobutton_PpP= Radiobutton(fenetre, variable=methode, value='Page par page',bg=marron2,
                            command=partial(update_methode,label_methode,methode))

entry_deb = Entry(fenetre, textvariable=numPageDeb,width=4)
entry_fin = Entry(fenetre, textvariable=numPageFin,width=4)


label_langue.grid(      row=3, column=2,sticky="w",padx=10)
label_Fr.grid(          row=4, column=2,sticky="w",padx=25)
radiobutton_Fr.grid(    row=4, column=3,sticky="w",padx = 10)
label_An.grid(          row=5, column=2,sticky="w",padx =25)
radiobutton_An.grid(    row=5, column=3,sticky="w",padx = 10)


label_methode.grid(     row=6, column=2,sticky="w",padx=10)
label_PpP.grid(         row=7, column=2,sticky="w",padx=25)
radiobutton_PpP.grid(   row=7, column=3,sticky="w",padx = 10)

label_In.grid(          row=8, column=2,sticky="w",padx = 25)
radiobutton_In.grid(    row=8, column=3,sticky="w",padx = 10)
label_deb.grid(         row=9, column=2)
entry_deb.grid(         row=9, column=3)
label_fin.grid(         row=10,column=2)
entry_fin.grid(         row=10,column=3)


fenetre.mainloop()
