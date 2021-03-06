# -*- coding: utf-8 -*-

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
from glob import glob

rsrcmgr = PDFResourceManager()
retstr = StringIO()
codec = 'utf-8'
laparams = LAParams()
laparams.detect_vertical = True
device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)

fp = open('pdf/HZN_001_20150701_81_01.pdf', 'rb')
interpreter = PDFPageInterpreter(rsrcmgr, device)
maxpages = 0
caching = True
pagenos=set()
fstr = ''

for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages,caching=caching, check_extractable=True):
    interpreter.process_page(page)

    str = retstr.getvalue()
    fstr += str

fp.close()
device.close()
retstr.close()