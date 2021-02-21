import PyPDF2
pdfFileObj = open(r'C:\Users\cmeneses\Desktop\Facturacion_claro_Septiembre.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
search_word = "Número Celular 3218140266"
search_word_count = 0
for pageNum in range(1, pdfReader.numPages):
    pageObj = pdfReader.getPage(pageNum)
    text = pageObj.extractText().encode('utf-8')
    search_text = text.lower().split()
    for word in search_text:
        if search_word in word.decode("utf-8"):
            search_word_count += 1
            print("The page was found in {} ".format(pageNum))
        
print("The word {} was found {} times".format(search_word, search_word_count))