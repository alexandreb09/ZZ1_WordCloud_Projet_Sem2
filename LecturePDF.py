import PyPDF2

fichier = "monfichier.pdf"

def get_pdf_content_lines(fich):
    with open(fich, "rb") as f:                                     # with open ... as : ouverture fichier dans le bloc et fermeture en sortie
        pdf_reader = PyPDF2.PdfFileReader(f)
        for page in pdf_reader.pages:
            for line in page.extractText().splitlines():
                yield line

for ligne in get_pdf_content_lines(fichier):
    print(ligne)
