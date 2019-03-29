# -*- coding: utf-8 -*-

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import glob

file_list = glob.glob('pdf/*')
for input_path in file_list:
    output_path = 'text/' + input_path[4:-4] + '.txt'
    rsrcmgr = PDFResourceManager()
    codec = 'utf-8'
    params = LAParams()
    with open(output_path, "wb") as output:
        device = TextConverter(rsrcmgr, output, codec=codec, laparams=params)
        with open(input_path, 'rb') as input:
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            for page in PDFPage.get_pages(input):
                interpreter.process_page(page)
        device.close()