from datetime import date

import aspose.words as aw

# load first document
doc = aw.Document(r"C:\Users\MohanMadhuri\PycharmProjects\thermofisher\doc1.docx")

# load second document
doc2 = aw.Document(r"C:\Users\MohanMadhuri\PycharmProjects\thermofisher\doc2.docx")

# compare documents
doc.compare(doc2, "user", date.today())

# save the document to get the revisions
if (doc.revisions.count > 0):
    doc.save("compared.docx")
else:
    print("Documents are equal")