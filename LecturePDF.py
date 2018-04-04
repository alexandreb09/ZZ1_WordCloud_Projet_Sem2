import textract


nomFichier = "Rapport1.pdf"
print("Fichier étudié : ",nomFichier)
ListeMotCle = []


def LongueurChaine(chaine):                             # Donne l'indice de fin des mots clés
    if chaine[0] == ' ': chaine=chaine[1:]
    nbMot = nbLigne = 0                                 # et la chaine sans les répétitions
    ind = 0
    while nbMot < 10 and nbLigne < 2:
        if chaine[ind] == ',': nbMot+=1
        if chaine[ind] == '\n': nbLigne+=1
        ind+=1
    return (ind,chaine[:ind+1])

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

def RecherDebMotCle(text):
    deb = 0
    while deb < longueur - l2 and text[deb:deb+l2]!=stop:       #recherche de l'occurence
        deb +=1

"""


stop = "Mots-clés"
l2 = len(stop)

text = textract.process(nomFichier, encoding='utf-8',method='pdfminer')
text = bytes.decode(text)


deb = text.index(stop)
fin,chaine = LongueurChaine(text[deb+l2+2:])
chaine = chaine.replace('\n',' ')                                   # suppresion des \n
chaine = chaine.replace(', ',',')
ListeMotCle = chaine.split(',')
AffichageMotCle(ListeMotCle)
