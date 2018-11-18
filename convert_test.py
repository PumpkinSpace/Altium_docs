# Altium convert pdf to text test.

import pyPdf
from functools import partial
import time
import pdfminer
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO

def convert_pdf_to_txt(path):
    """
    Function to extract the text from a pdf that contains embedded text.
    
    Based on Chianti5's code from:
    stackoverflow.com/questions/40031622/pdfminer-error-for-one-type-of-pdfs-
         too-many-vluae-to-unpack

    @param[in]    path:          The file path of the pdf to read
                                 (string).
    @return       (string)       The extracted text. 
    """ 
    
    # create a PDF resource manager object that stores shared resources
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    
    # set parameters for analysis
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = file(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()
    text = []

    # process each page in the pdf
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, \
                                  password=password,caching=caching, \
                                  check_extractable=True):
        interpreter.process_page(page)
        # extract the text
        text.append(retstr.getvalue())
    # end for

    # close all files
    fp.close()
    device.close()
    retstr.close()
    
    # return the text
    return text[0]
# end def

def converter_by_james(path):
    return 'test_string'
# end



# test code
start_time = time.time()
old_text = convert_pdf_to_txt('00337E0-8.pdf')
old_time = start_time - time.time()

start_time = time.time()
new_text = converter_by_james('00337E0-8.pdf')
new_time = start_time - time.time()

print "The new converter is " + str(100*(old_time-new_time)/old_time) + "% faster"

if 'Modification' not in new_text:
    print "However the new converter is missing important text information in the PDF"
    
else:
    print "The new converter appears to be parsing the file correctly"
# end if
    