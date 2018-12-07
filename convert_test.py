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
import multiprocessing
import os




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


def split_pages(pdf_path, pdf_prefix):
    page_list = []
    # split the pdf into pages
    with open(pdf_path, "rb") as schematic_file:
        schematic = pyPdf.PdfFileReader(schematic_file)
        
        # write each page to a separate pdf file
        for page in xrange(schematic.numPages):
            # add page to the output stream
            output = pyPdf.PdfFileWriter()
            output.addPage(schematic.getPage(page))
            # format the filename 
            file_name = pdf_prefix + '-' + str(page+1) + '.pdf'
            
            with open(file_name, "wb") as outputStream:
                # write the page
                output.write(outputStream)
            # end with
            page_list.append(file_name)
        # end for
    # end with   
    return page_list
# end def

def cleanup_pdfs(pdf_list):
    for filename in pdf_list:
        os.remove(filename)
    # end for
# end def    
    

def create_process_threads(process, arg_list):
    thread_list = []
    for arg in arg_list:
        # define the thread to perform the writing
        thread = multiprocessing.Process(name=('renaming-' + str(i)),
                                         target = process, 
                                         args=arg)
        # start the thread
        thread.start()
        
        thread_list.append(thread)
    # end for   
    
    # wait for all the threads to finish
    while any([t.is_alive() for t in thread_list]):
        time.sleep(0.1)
    # end while    
# test code

# blocking for multiprocessing
if __name__ == '__main__':

    print "Single page test"
    start_time = time.time()
    old_text = convert_pdf_to_txt('00337E0-8.pdf')
    old_time = time.time() - start_time
    
    start_time = time.time()
    new_text = converter_by_james('00337E0-8.pdf')
    new_time = time.time() - start_time
    
    print "\told_time = " + str(old_time) + "s, new_time = " + str(new_time) + "s"
    print "\tThe new converter is " + str(100*(old_time-new_time)/old_time) + "% faster"
    
    if 'Modification' not in new_text:
        print "\tHowever the new converter is missing important text information in the PDF"
        
    else:
        print "\tThe new converter appears to be parsing the file correctly"
    # end if
    
    
    
    print "\nSchematic - single thread test"
    start_time = time.time()
    old_text = convert_pdf_to_txt('Schematic.pdf')
    old_time = time.time() - start_time
    
    start_time = time.time()
    new_text = converter_by_james('Schematic.pdf')
    new_time = time.time() - start_time
    
    print "\told_time = " + str(old_time) + "s, new_time = " + str(new_time) + "s"
    print "\tThe new converter is " + str(100*(old_time-new_time)/old_time) + "% faster"
    
    
    
    
    print "\nSchematic - multi thread test"
    
    page_list = split_pages('Schematic.pdf', 'page')
    arg_list = [(i,) for i in page_list]
    
    start_time = time.time()  
    create_process_threads(convert_pdf_to_txt, arg_list)
    old_time = time.time() - start_time
    
    start_time = time.time()
    create_process_threads(converter_by_james, arg_list)
    new_time = time.time() - start_time
    
    cleanup_pdfs(page_list)
    
    print "\told_time = " + str(old_time) + "s, new_time = " + str(new_time) + "s"
    print "\tThe new converter is " + str(100*(old_time-new_time)/old_time) + "% faster"
    

    
    
    #print "\nLayers - single thread test"
    #start_time = time.time()
    #old_text = convert_pdf_to_txt('Layers with Text.PDF')
    #old_time = time.time() - start_time
    
    #start_time = time.time()
    #new_text = converter_by_james('Layers with Text.PDF')
    #new_time = time.time() - start_time
    
    #print "\told_time = " + str(old_time) + "s, new_time = " + str(new_time) + "s"
    #print "\tThe new converter is " + str(100*(old_time-new_time)/old_time) + "% faster"
    
    
    
    
    print "\nLayers - multi thread test"
    
    page_list = split_pages('Layers with Text.PDF', 'layer')
    arg_list = [(i,) for i in page_list]
    
    start_time = time.time()  
    create_process_threads(convert_pdf_to_txt, arg_list)
    old_time = time.time() - start_time
    
    start_time = time.time()
    create_process_threads(converter_by_james, arg_list)
    new_time = time.time() - start_time
    
    cleanup_pdfs(page_list)
    
    print "\told_time = " + str(old_time) + "s, new_time = " + str(new_time) + "s"
    print "\tThe new converter is " + str(100*(old_time-new_time)/old_time) + "% faster"    
    
# end if