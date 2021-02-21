from PyPDF2 import PdfFileReader, PdfFileWriter
import re

# open the pdf file
object = PdfFileReader(r'C:\Users\cmeneses\Desktop\Facturacion_claro_Septiembre.pdf')

# get number of pages
NumPages = object.getNumPages()

# define keyterms
String = "número celular 3218140266"

# define obj pdf writer
pdf_writer = PdfFileWriter()

# extract text and do the search
for page in range(NumPages):
    PageObj = object.getPage(page)
    # print("this is page " + str(i)) 
    Text = PageObj.extractText() 
    findText = re.findall(String, Text.lower())
    if findText:
        pdf_writer.addPage(PageObj)
        with open("doc1.pdf", 'wb') as file1:
            pdf_writer.write(file1)