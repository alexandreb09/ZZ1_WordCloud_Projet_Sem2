from textract import process
from os import listdir, chdir,remove
from PyPDF2 import PdfFileReader,PdfFileMerger

def LongueurChaine(chaine):                             # Donne l'indice de fin des mots clés
    if chaine[0] == ' ': chaine=chaine[1:]
    ind = 0
    while chaine[ind:ind+2] != '\n\n' and ind<len(chaine)-2:                    # tant qu'il n'y a pas 2 retours à a ligne consécutifs
        ind+=1
    return chaine[:ind+1]

def AffichageMotCle(Liste):
    print("\nListe des mots clés :")
    for elt in Liste:
        print(elt)

"""
def AjoutMotsCles(chaine,ListeMotCle):          #fonction split (ici sans espace)
    print('ma chaine : ',chaine)
    cour = debMot = 0
    while debMot + cour  < fin:
        if chaine[debMot + cour] == ',':
            print(chaine[debMot:debMot+cour])
            ListeMotCle.append(chaine[debMot:debMot+cour])
            debMot = debMot + cour + 2
            cour = 0
        else : cour+=1
    return(ListeMotCle)
"""

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

langue='fr'
chdir('./Documents')
ListeMotCle = []
l1 = 9
if langue == 'an': l1 = 7
SuppTemp()
listeFichier = listdir()

for fichier in listeFichier:
    print("Fichier étudié : ", fichier)
    fichier = creerFichierPdfTemp(fichier,2,6)
    text = convertPDF2Txt(fichier)
    text = text.replace('  ',' ')
    deb = RecherDebMotCle(text)
    if deb!= -1:
        chaine = LongueurChaine(text[deb+l1+2:])
        chaine = chaine.replace('\n',' ')                                       # suppresion des \n
        chaine = chaine.replace(', ',',')
        ListeMotCle = ListeMotCle + chaine.split(',')
ListeMotCle = remake(ListeMotCle)
AffichageMotCle(ListeMotCle)

SuppTemp()
