##This was made to merge multiple pdf scans into one file for my boss

import os
from PyPDF2 import PdfFileMerger

##get all pdf files in directory
pdflist = [a for a in os.listdir() if a.endswith(".pdf")]
 
merger = PdfFileMerger()
 
for pdf in pdflist:
    merger.append(open(pdf, 'rb'))
 
with open("result.pdf", "wb") as fout:
    merger.write(fout)
