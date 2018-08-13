#!/usr/bin/env python
###########################################################################
#(C) Copyright Pumpkin, Inc. All Rights Reserved.
#
#This file may be distributed under the terms of the License
#Agreement provided with this software.
#
#THIS FILE IS PROVIDED AS IS WITH NO WARRANTY OF ANY KIND,
#INCLUDING THE WARRANTY OF DESIGN, MERCHANTABILITY AND
#FITNESS FOR A PARTICULAR PURPOSE.
###########################################################################
"""
@package Altium_OCR.py

Package that manages to OCR of pdfs, and subsequent pdf actions as needed by 
the Altium Documentation module.
"""

__author__ = 'David Wright (david@asteriaec.com)'
__version__ = '0.2.0' #Versioning: http://www.python.org/dev/peps/pep-0386/


#
# -------
# Imports

import os
import sys
sys.path.insert(1, 'src\\')
import shutil
import subprocess
import pyPdf
from functools import partial
import time
import Altium_helpers
import pdfminer
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO
import multiprocessing

# This program also requires the following installed packages:  
# pypdfocr 
# imagemagik
# Pillow
# reportlab
# watchdog
# pypdf2
# ghostscript
#import reportlab
#import watchdog
#import PyPDF2

#
# ----------------
# Public Functions 

def log_error(get = False):
    """
    Function to log errors within this module.

    @param[in]    get:        True  = return no_errors without logging an error
                              False = log an error and return nothing (bool)
    @attribute    no_errors:  Whether there have been no errors logged
    @return       (bool)      True  = no errors have been logged.
                              False = Errors have been logged.
    """  
    
    # determine which action to take
    if get:
        # return the state
        return log_error.no_errors
    
    else:
        # log an error
        log_error.no_errors = False
    # end if
# end def

# set the initial value
log_error.no_errors = True


def log_warning(get = False):
    """
    Function to log warnings within this module.

    @param[in] get:          True  = return no_warnings without logging a warning
                             False = log a warning and return nothing (bool)
    @attribute no_warnings:  Whether there have been no errors logged
    @return    (bool)        True  = no errors have been logged.
                             False = Errors have been logged.
    """    
    
    # determine which action to take
    if get:
        # return the state
        return log_warning.no_warnings
    
    else:
        # log a warning
        log_warning.no_warnings = False
    # end if
# end def

# set the inital state
log_warning.no_warnings = True


def adjust_layer_filename(starting_dir):
    """
    Function to adjust the file name of the layers pdf to the desired filename 
    from one of the possible output options.

    @param[in]    starting_dir:    The Altium project directory (full path) 
                                   (string).
    @return       (mod_date)       The modification of the layers pdf. 
    """      
    
    # get the root file list
    root_file_list = os.listdir(starting_dir)
    
    # Find the layers pdf file
    if (('Layers.pdf' not in root_file_list) 
        and ('layers.pdf' not in root_file_list) 
        and ('PCB Prints.pdf' not in root_file_list)):
        # Could not find layers.pdf
        print('***  Error: No layers.pdf or PCB Prints.pdf file found  ***')
        log_error()
        return None
    # end if
    
    # get the modification date of the file and then re-name it if need be
    if ('Layers.pdf' in root_file_list):
        mod_date = Altium_helpers.mod_date(os.path.getmtime(starting_dir+'\\Layers.pdf'), 
                                           'layers.pdf')
        os.rename(starting_dir+'\\Layers.pdf', starting_dir+'\\layers.pdf')
    
    elif ('PCB Prints.pdf' in root_file_list):
        mod_date = Altium_helpers.mod_date(os.path.getmtime(starting_dir+'\\PCB Prints.pdf'),
                                           'layers.pdf')
        os.rename(starting_dir+'\\PCB Prints.pdf', starting_dir+'\\layers.pdf')
    
    else:
        mod_date = Altium_helpers.mod_date(os.path.getmtime(starting_dir+'\\layers.pdf'),
                                           'layers.pdf')
    # end if  
    
    return mod_date
# end def


def get_filename_init():
    """
    Initialise the static variables for the get_filename function to facilitate 
    error detection.
    """
    
    get_filename.layer = 0
    get_filename.MECHDWG = False
    get_filename.ADB = False
    get_filename.ADT = False
    get_filename.SST = False
    get_filename.SMT = False
    get_filename.SSB = False
    get_filename.SMB = False
    get_filename.DD = False
    get_filename.SPB = False
    get_filename.SPT = False
# end def


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

    # process each page in the pdf
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, \
                                  password=password,caching=caching, \
                                  check_extractable=True):
        interpreter.process_page(page)
    # end

    # extract the text
    text = retstr.getvalue()

    # close all files
    fp.close()
    device.close()
    retstr.close()
    
    # return the text
    return text
# end
    
    
def perform_Altium_OCR(exe_OCR, starting_dir, num_layers, 
                       silence = True, with_threads = False):
    """
    Function to perform OCR on the layers pdf and then extract the text from 
    them and split and rename the pages accordingly.

    @param[in]  exe_OCR:            Whether to use the .exe file for OCR (bool)
    @param[in]  starting_dir:       The Altium project directory (full path) 
                                    (string).    
    @param[in]  num_layers:         The number of layers in the PCB (int).
    @param[in]  silence:            Whether to silence the output of the OCR 
                                    engine (bool).
    @param[in]  with_threads:       Split OCR and page renaming into threads for
                                    execution (bool).
    @return     (list of mod_dates) The modification dates of the files used.
    """     
    
    # correct the filename of the layers pdf if required
    modified_dates = [adjust_layer_filename(starting_dir)]
    
    # initialise the filename checker variables
    get_filename_init()  
    
    pdf_dir = Altium_helpers.get_pdf_dir(starting_dir)
    
    if ocr_is_required(starting_dir):
        
        print 'Starting OCR on Layers file...'
        
        # determine execution type
        if with_threads:
            # run OCR as threads
            thread_OCR_init(starting_dir, exe_OCR, silence)
            
        else:
            # run OCR on the layers pdf
            run_OCR(starting_dir, exe_OCR, silence)
            
            # if OCR did not return the right pdf then it must have failed
            if not os.path.isfile(starting_dir +'\\layers_ocr.pdf'):
                log_error()
                print '*** Error: OCR was unsuccessful ***'
                return None, None
            # end if
            
            print '\tRenaming the layer PDFs...'
            
            # split the OCR pdf into multiple pages and then rename the appropriately
            split_OCR_pages(starting_dir)
        # end if
    
    else:
        
        print 'Splitting Layers file without OCR...'
        
        # open pdf file and split into pages
        try:
            with open(starting_dir+'\\layers.pdf', "rb") as layers_file:
                layers_pdf = pyPdf.PdfFileReader(layers_file)
                
                # write each page to a separate pdf file
                for page in xrange(layers_pdf.numPages):
                    # add page to the output stream
                    output = pyPdf.PdfFileWriter()
                    output.addPage(layers_pdf.getPage(page))
                    # format the filename 
                    file_name = pdf_dir + '\\layer--' + str(page+1) + '.pdf'
                    
                    with open(file_name, "wb") as outputStream:
                        # write the page
                        output.write(outputStream)
                    # end with
                # end for
            # end with
            
        except:
            print('***   Error: Could not open Layers.pdf document   ***')
            log_error()
            return None        
        # end try
        
        print '\tComplete!'
        
        print '\tRenaming the PDFs...'
        
        # rename the sheets with threads or without
        if with_threads:
            # initialise list of threads
            thread_list = []
            
            thread_queue = multiprocessing.Queue()
            
            # start a thread to rename each page
            for i in range(1,page+2):
                # define the thread to perform the writing
                thread = multiprocessing.Process(name=('renaming-' + str(i)),
                                                 target = rename_layer, 
                                                 args=(starting_dir,i,thread_queue))
                # start the thread
                thread.start()
                
                thread_list.append(thread)
            # end for       
            
            # wait for all the threads to finish
            while any([t.is_alive() for t in thread_list]):
                time.sleep(0.01)
            # end while
            
            # read all data from the queue
            thread_data = []
            
            # retrieve data from the queue until empty
            while True:
                try:
                    thread_data.append(thread_queue.get(block=False))
                
                except:
                    break
            # end while
                               
            if any([q == False for q in thread_data]):
                # an error occurred
                log_error()
            # end if
            
            for name in [item for item in thread_data if (type(item) == str)]:
                if name == 'MECHDWG.pdf':
                    get_filename.MECHDWG = True
            
                elif name == 'ADB0230.pdf':
                    get_filename.ADB = True
            
                elif name ==  'ADT0127.pdf':
                    get_filename.ADT = True
            
                elif name == 'SST0126.pdf':
                    get_filename.SST = True
            
                elif name == 'SMT0125.pdf':
                    get_filename.SMT = True
        
                elif name == 'SSB0229.pdf':
                    get_filename.SSB = True
            
                elif name == 'SMB0223.pdf':
                    get_filename.SMB = True
            
                elif name == 'DD0124.pdf':
                    get_filename.DD = True
        
                elif name == 'SPB0223.pdf':
                    get_filename.SPB = True
            
                elif name == 'SPT0123.pdf':
                    get_filename.SPT = True
            
                elif name.startswith('ART'):
                    get_filename.layer += 1
                # end if
            # end for            
            
        else:
            # rename the pdfs with the correct filenames
            for i in range(1,page+2):
                rename_layer(starting_dir, i)
            # end for
        # end if        
    # end if
    
    # check the output files to ensure there is nothing wrong
    check_OCR_outputs(num_layers)    
    
    print '\tComplete!\nComplete!\n'
    
    return modified_dates
# end def  


def check_DRC(starting_dir):
    """
    Checks the design rule check output PDF to see if there are any errors

    @param:    starting_dir   The full path of the Altium Folder (string).
    @return:   (mod_date)     The modification date of the Design Rule Check
    """  
    print '\nChecking the Design Rule Check...'
    file_list = os.listdir(starting_dir)
    
    
    if 'Design Rules Check.PDF' not in file_list:
        print '*** Error: No design rule check has been completed ***'
        log_error()
        return None
    # end if
    
    # get the modification date of the file
    DRC_date = Altium_helpers.mod_date(os.path.getmtime(starting_dir+'\\Design Rules Check.PDF'), 
                                       'Design Rules Check.PDF')
    
    # extract text and remove whitespace
    DRC_text = "".join(convert_pdf_to_txt(starting_dir+'\\Design Rules Check.PDF').split())
    
    if 'Warnings0' not in DRC_text:
        print '*** Warning: Warnings were raised during Altium DRC ***'
        log_warning()
    # end if
    
    if 'Violations0' not in DRC_text:
        print '*** Warning: Rule Violations were found during Altium DRC ***'
        log_warning()
    # end if
    
    print 'Complete!\n'
    
    return DRC_date
# end def


def check_ERC(starting_dir):
    """
    Checks the electrical rule check output PDF to see if there are any errors

    @param:    starting_dir   The full path of the Altium Folder (string).
    @return:   (mod_date)     The modification date of the Electrical Rule Check
    """  
    print '\nChecking the Electrical Rule Check...'
    file_list = os.listdir(starting_dir)
    
    
    if 'Electrical Rules Check.PDF' not in file_list:
        print '*** Error: No electrical rule check has been completed ***'
        log_error()
        return None
    # end if
    
    # get the modification date of the file
    ERC_date = Altium_helpers.mod_date(os.path.getmtime(starting_dir+'\\Electrical Rules Check.PDF'), 
                                       'Electrical Rules Check.PDF')
    
    # extract text and remove whitespace
    ERC_text = "".join(convert_pdf_to_txt(starting_dir+'\\Electrical Rules Check.PDF').split())
    
    if 'Warning' in ERC_text:
        print '*** Warning: Warnings were raised during Altium ERC ***'
        log_warning()
    # end if
    
    if 'Error' in ERC_text:
        print '*** Warning: Rule Violations were found during Altium ERC ***'
        log_warning()
    # end if
    
    print 'Complete!\n'
    
    return ERC_date
# end def


#
# ----------------
# Private Functions 

def rename_layer(starting_dir, sheet_number, queue = None):
    """
    Function to rename a layer document based on information contained within 
    it.

    @param[in]   starting_dir:        The Altium project directory (full path) 
                                      (string).
    @param[in]   sheet_number:        The number of the sheet to rename (int)
    @param[out]  queue:               The queue to return data to if running
                                      as a thread (Queue).
    """  
    # get the pdf directory
    pdf_dir = Altium_helpers.get_pdf_dir(starting_dir)
       
    old_filename = pdf_dir + '\\layer--' + \
        str(sheet_number) + '.pdf'
            
    # find new name
    new_filename = get_filename(old_filename, queue)
    
    if (new_filename != None):       
        # rename the file
        try:
            os.rename(old_filename, pdf_dir + '\\' + new_filename)
            
        except:
            print('***   Error: Could rename pdf document   ***')
            
            if queue == None:
                log_error()
                
            else:
                queue.put(False)
            # end if
        # end try
        
    else:
        os.remove(old_filename)
    # end if
# end def


def ocr_is_required(starting_dir):
    """
    Checks to see if the layers pdf has embedded text. if it doesn't then 
    OCR is required.

    @param:    starting_dir   The full path of the Altium Folder (string).
    @return:   (boolean)      True if there is no embedded text.
    """
    
    layers_text = convert_pdf_to_txt(starting_dir+'\\layers.pdf')
    try:
        return not ('layer' in layers_text.lower())
    
    except:
        return False
    # end try
# end def


def thread_OCR_init(starting_dir, exe_OCR, silence = True):
    """
    Function to perform OCR on the layers pdf and then extract the text from 
    them and split and rename the pages accordingly.

    @param[in]  starting_dir:       The Altium project directory (full path) 
                                    (string).    
    @param[in]  exe_OCR:            Whether to use the .exe file for OCR (bool)
    @param[in]  silence:            Whether to silence the output of the OCR 
                                    engine (bool).
    """  
    
    # get ocr directory
    ocr_dir = get_OCR_dir(exe_OCR)
    
    # open pdf file and split into pages and put them in the ocr_directory
    try:
        with open(starting_dir+'\\layers.pdf', "rb") as layers_file:
            layers_doc = pyPdf.PdfFileReader(layers_file)
            
            # write each page to a separate pdf file
            for page in xrange(layers_doc.numPages):
                # add page to the output stream
                output = pyPdf.PdfFileWriter()
                output.addPage(layers_doc.getPage(page))
                # format the filename 
                file_name = ocr_dir + '\\layer--' + str(page+1) + '.pdf'
                
                with open(file_name, "wb") as outputStream:
                    # write the page
                    output.write(outputStream)
                # end with
            # end for
        # end with
        
    except:
        print('***   Error: Could not open layers document   ***')
        log_error()      
        return None
    # end try   
    
    # get the list of layer files in the ocr_directory
    layer_files = [filename for filename in os.listdir(ocr_dir) if filename.startswith('layer--')]
    
    # initialise threading trackers
    thread_queue = multiprocessing.Queue()
    thread_list = []
    
    # iterate through the layer files
    for layer in layer_files:
        # define the thread to perform the writing
        thread = multiprocessing.Process(name=('OCR-' + layer),
                                         target = OCR_thread, 
                                         args=(starting_dir,exe_OCR, layer,
                                               thread_queue, silence))
        # start the thread
        thread.start()
    
        thread_list.append(thread)    
    # end for
    
    # wait for all the threads to finish
    while any([t.is_alive() for t in thread_list]):
        time.sleep(0.01)
    # end while
    
    # read all data from the queue
    thread_data = []
    
    # retrieve data from the queue until empty
    while True:
        try:
            thread_data.append(thread_queue.get(block=False))
        
        except:
            break
    # end while
    
    if any([q == False for q in thread_data]):
        # an error occurred
        log_error()
    # end if
    
    for name in [item for item in thread_data if (type(item) == str)]:
        if name == 'MECHDWG.pdf':
            get_filename.MECHDWG = True
    
        elif name == 'ADB0230.pdf':
            get_filename.ADB = True
    
        elif name ==  'ADT0127.pdf':
            get_filename.ADT = True
    
        elif name == 'SST0126.pdf':
            get_filename.SST = True
    
        elif name == 'SMT0125.pdf':
            get_filename.SMT = True

        elif name == 'SSB0229.pdf':
            get_filename.SSB = True
    
        elif name == 'SMB0223.pdf':
            get_filename.SMB = True
    
        elif name == 'DD0124.pdf':
            get_filename.DD = True

        elif name == 'SPB0223.pdf':
            get_filename.SPB = True
    
        elif name == 'SPT0123.pdf':
            get_filename.SPT = True
    
        elif name.startswith('ART'):
            get_filename.layer += 1
        # end if
    # end for
# end def
        
        
def OCR_thread(starting_dir, exe_OCR, filename, queue, silence = True):
    """
    Function to perform OCR on a single pdf page as 

    @param[in]  starting_dir:       The Altium project directory (full path) 
                                    (string).    
    @param[in]  exe_OCR:            Whether to use the .exe file for OCR (bool).
    @param[in]  filename:           The filename on which toi perform OCR (string).
    @param[out] queue:              The queue to write output to (Queue).
    @param[in]  silence:            Whether to silence the output of the OCR 
                                    engine (bool).
    """     
    
    # get the correct directory in which to perform OCR
    ocr_dir = get_OCR_dir(exe_OCR)
    
    # get the pdf directory
    pdf_dir = Altium_helpers.get_pdf_dir(starting_dir)
    
    # determine which type of OCR to do
    if exe_OCR:
        
        if silence:
            with open(os.devnull, 'w') as out:
                # perform OCR on the layers pdf
                cmd = subprocess.Popen(['pypdfocr.exe', filename], 
                                       cwd=ocr_dir, shell=True, 
                                       stdout=out, stderr=out)    
            # end with
        
        else:
            # perform OCR on the layers pdf
            cmd = subprocess.Popen(['pypdfocr.exe', filename], cwd=ocr_dir, shell=True)
        # end if
        
        # wait for analysis to complete
        cmd.wait()
        
    else:
        if silence:
            with open(os.devnull, 'w') as out:
                # perform OCR on the layers pdf
                cmd = subprocess.Popen(['python', 'pypdfocr.py', filename], 
                                       cwd=ocr_dir, stdout=out, stderr=out)   
            # end with
        
        else:        
            # perform OCR on the layers pdf
            cmd = subprocess.Popen(['python', 'pypdfocr.py', filename], 
                                   cwd=ocr_dir)
        # end if
        
        # wait for analysis to complete
        cmd.wait()
    # end if
        
    # determine the name of the OCR'ed file
    ocr_filename = '_ocr.'.join(filename.split('.'))
    
    try:
        # return OCR file from OCR directory and clean the OCR directory
        os.remove(ocr_dir + '\\' + filename)
        shutil.move(ocr_dir + '\\' + ocr_filename, pdf_dir + '\\' + filename)      
        
    except:
        queue.put(False)
        print '*** Error: OCR was unsuccessful on ' + filename + ' ***'
        return None
    # end try 
    
    # find the desired filename for the new file
    new_filename = get_filename(pdf_dir + '\\' + filename, queue)
    
    # performs appropriate actions on the file
    if new_filename != None:
        # This file is desired so rename with the correct name
        try:
            os.rename(pdf_dir + '\\' + filename, pdf_dir + '\\' + new_filename)
            
        except:
            print '*** Error, ' + new_filename + ' already exists ***'
            queue.put(False)
        # end try
        
    else:
        # file is not wanted so remove it
        os.remove(pdf_dir + '\\' + filename)
    # end if   
# end def
    

def get_filename(path, queue = None):
    """
    Determine the correct name for a pdf file by looking in the text for 
    certain substrings.

    @param[in]  path:       The path of the pdf file to rename (string). 
    @param[out] queue:      Queue to write output to in threading operations 
                            (Queue).
    @return     (string)    The correct filename for that file.
    """      
    # extract the text from the pdf
    pdf_text = beautify(convert_pdf_to_txt(path))
    
    # look for certain substrings to determin the correct filename
    if (beautify('number') in pdf_text)\
       and (beautify('round') in pdf_text):
        # This is a Mechanical Drawing file
        if queue == None:
            get_filename.MECHDWG = True
            
        else:
            queue.put('MECHDWG.pdf')
        # end if
        
        return 'MECHDWG.pdf'
    
    elif (beautify('Bottom') in pdf_text)\
         and (beautify('silk') in pdf_text)\
         and (beautify('Solder') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Bottom Assembly Drawing file
        if queue == None:
            get_filename.ADB = True
            
        else:
            queue.put('ADB0230.pdf')
        # end if        

        return 'ADB0230.pdf'
    
    elif (beautify('Top') in pdf_text)\
         and (beautify('silk') in pdf_text)\
         and (beautify('Solder') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Top Assembly Drawing file
        if queue == None:
            get_filename.ADT = True
            
        else:
            queue.put('ADT0127.pdf')
        # end if         
        
        return 'ADT0127.pdf'
    
    elif (beautify('Top') in pdf_text)\
         and (beautify('silk') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Top Silkscreen File
        if queue == None:
            get_filename.SST = True
            
        else:
            queue.put('SST0126.pdf')
        # end if         
        return 'SST0126.pdf' 
    
    elif (beautify('Top') in pdf_text)\
         and (beautify('mask') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Top Soldermask file
        if queue == None:
            get_filename.SMT = True
            
        else:
            queue.put('SMT0125.pdf')
        # end if         

        return 'SMT0125.pdf'     
    
    elif (beautify('Bottom') in pdf_text)\
         and (beautify('silk') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Bottom Silkscreen file
        if queue == None:
            get_filename.SSB = True
            
        else:
            queue.put('SSB0229.pdf')
        # end if         

        return 'SSB0229.pdf'  
    
    elif (beautify('Bottom') in pdf_text)\
         and (beautify('mask') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Bottom Soldermask file
        if queue == None:
            get_filename.SMB = True
            
        else:
            queue.put('SMB0223.pdf')
        # end if         

        return 'SMB0223.pdf'  
    
    elif (beautify('Drill') in pdf_text)\
         and (beautify('Drawing') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Drill Drawing File
        if queue == None:
            get_filename.DD = True
            
        else:
            queue.put('DD0124.pdf')
        # end if         

        return 'DD0124.pdf' 
    
    elif (beautify('Bottom') in pdf_text)\
         and (beautify('Solder') in pdf_text)\
         and (beautify('Paste') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Bottom Solder Paste file   
        if queue == None:
            get_filename.SPB = True
            
        else:
            queue.put('SPB0223.pdf')
        # end if         

        return 'SPB0223.pdf'  
    
    elif (beautify('Top') in pdf_text)\
         and (beautify('Solder') in pdf_text)\
         and (beautify('Paste') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Top Solder Paste file
        if queue == None:
            get_filename.SPT = True
            
        else:
            queue.put('SPT0123.pdf')
        # end if         

        return 'SPT0123.pdf'   
    
    elif ((beautify('Layer') in pdf_text) or (beautify('Plane') in pdf_text)) \
         and (beautify('COMPO') not in pdf_text):
        # This is a Layer Artwork file
        #name = 'ART' + format(get_filename.layer, '02') + '.pdf'
        if queue == None:
            get_filename.layer += 1
            
        else:
            queue.put(get_layer_number(pdf_text))
        # end if         
        
        #return name
        return get_layer_number(pdf_text)
    
    else:
        # no appropriate name was found so return unsuccessful
        return None
    # end if
# end def


def get_layer_number(page_text):
    """
    Determine the correct name for a layer pdf file by looking in the text for 
    certain substrings.

    @param[in]  pdf_text:       The text from the pdf (string). 
    @return     (string)        The correct filename for that file.
    """      
    if (beautify('Layer 1') in page_text):
        return 'ART01.pdf'
    
    elif (beautify('Layer 2') in page_text):
        return 'ART02.pdf'
    
    elif (beautify('Layer 3') in page_text):
        return 'ART03.pdf'
    
    elif (beautify('Layer 4') in page_text):
        return 'ART04.pdf'

    elif (beautify('Layer 5') in page_text):
        return 'ART05.pdf'
    
    elif (beautify('Layer 6') in page_text):
        return 'ART06.pdf'

    elif (beautify('Layer 7') in page_text):
        return 'ART07.pdf'
    
    elif (beautify('Layer 8') in page_text):
        return 'ART08.pdf'
    
    else:
        # renaiming this layer was unsuccessful
        print "this layer did not get named"
        print beautify('Layer 1')
        return "unnamed layer.pdf"
    #end if
#end def


def split_OCR_pages(starting_dir):
    """
    Function to split the pages of the OCR'ed pdfs into appropriately named 
    pdfs.

    @param[in]  starting_dir:       The Altium project directory (full path) 
                                    (string).    
    """    
    
    # open the OCR'ed pdf
    try:
        # read the OCR'ed file
        layers_pdf = pyPdf.PdfFileReader(open(starting_dir+'\\layers_ocr.pdf', "rb"))
        
    except:
        print '***  Error: Could not open layers_ocr.pdf ***'
        log_error()
        return None
    # end try
    
    # get the directory to put the split pages into
    pdf_dir = Altium_helpers.get_pdf_dir(starting_dir)
    
    # write each page to a separate pdf file
    for page in xrange(layers_pdf.numPages):
        # add page to the ouput writer
        output = pyPdf.PdfFileWriter()
        output.addPage(layers_pdf.getPage(page))
        
        # filename to write for each layer
        file_name = pdf_dir + '\\' + 'layer-' + str(page+1) + '.pdf'
        
        with open(file_name, "wb") as outputStream:
            # write file
            output.write(outputStream)
        # end with
        
        # find the desired filename for the new file
        new_filename = get_filename(file_name)
        if new_filename != None:
            # This file is desired so rename with the correct name
            try:
                os.rename(file_name, pdf_dir + '\\' + new_filename)
                
            except:
                print '*** Error, ' + new_filename + ' already exists ***'
                log_error()
            # end try
            
        else:
            # file is not wanted so remove it
            os.remove(file_name)
        # end if
    # end for
# end def


def beautify(text):
    """
    Function to replace common mis-read letter resulting from the OCR process. 
    This subsitiutes letters to minimise the number of errors.

    @param[in]  text:       The text to clean up (string).
    @return     (string)    The cleaned up text
    """      
    # If the text is a list then join it together
    text = ''.join(text.split())
    
    # lower case
    text = text.lower()
    
    # remove all non-alphanumeric characters
    text = ''.join([c for c in text if c.isalnum()])
    
    # letter replacement
    text = text.replace('5', 's')
    text = text.replace('1', 'l')
    text = text.replace('i', 'l')
    text = text.replace('p', 'r')
    text = text.replace('0', 'o')
    #text = text.replace('j', 'l')
    #text = text.replace('g', 'y')
    #text = text.replace('h', 'a')
    #text = text.replace('u', 'w')
    text = text.replace('f', 'r')
    text = text.replace('\'', '')
    
    # return the cleaned up text
    return text
# end def


def run_OCR(starting_dir, exe_OCR, silence = True):
    """
    Function to perform OCR on the layers pdf.

    @param[in]  starting_dir:       The Altium project directory (full path) 
                                    (string).    
    @param[in]  exe_OCR:            Whether to use the .exe file for OCR (bool)
    @param[in]  silence:            Whether to silence the output of the OCR 
                                    engine (bool).
    """    
    
    # get the correct directory in which to perform OCR
    ocr_dir = get_OCR_dir(exe_OCR)
    
    # determine which type of OCR to do
    if exe_OCR:
        # copy the layers pdf into the ocr directory to allow ocr to be performed
        shutil.copy(starting_dir+'\\layers.pdf', ocr_dir +'\\layers.pdf')
        
        if silence:
            with open(os.devnull, 'w') as out:
                # perform OCR on the layers pdf
                cmd = subprocess.Popen(['pypdfocr.exe', 'layers.pdf'], 
                                       cwd=ocr_dir, shell=True, 
                                       stdout=out, stderr=out)    
            # end with
        
        else:
            # perform OCR on the layers pdf
            cmd = subprocess.Popen(['pypdfocr.exe', 'layers.pdf'], cwd=ocr_dir, shell=True)
        # end if
        
        # wait for analysis to complete
        cmd.wait()
        
    else:
        
        # copy the layers pdf into the ocr directory to allow ocr to be performed
        shutil.copy(starting_dir+'\\layers.pdf', ocr_dir +'\\layers.pdf')
        
        if silence:
            with open(os.devnull, 'w') as out:
                # perform OCR on the layers pdf
                cmd = subprocess.Popen(['python', 'pypdfocr.py', 'layers.pdf'], 
                                       cwd=ocr_dir, stdout=out, stderr=out)   
            # end with
        
        else:        
            # perform OCR on the layers pdf
            cmd = subprocess.Popen(['python', 'pypdfocr.py', 'layers.pdf'], 
                                   cwd=ocr_dir)
        # end if
        
        # wait for analysis to complete
        cmd.wait()
    # end if
        
    try:
        # return OCR file from OCR directory and clean the OCR directory
        os.remove(ocr_dir +'\\layers.pdf')
        shutil.move(ocr_dir +'\\layers_ocr.pdf', starting_dir +'\\layers_ocr.pdf')      
        
    except:
        pass
    # end try
# end def


def check_OCR_outputs(num_layers):
    """
    Check to see if anything is missing in the output files

    @param[in]  num_layers:         The number of layers in the PCB (int).
    """    
    
    # Generate warnings for pecuiliar outputs
    if (get_filename.layer != num_layers):
        print '*** WARNING wrong number of layers printed ***'
        log_warning()
    # end
    if (get_filename.MECHDWG == False):
        print '*** WARNING No MECHDWG file output ***'
        log_warning()
    # end
    if (get_filename.ADB == False):
        print '*** WARNING No ADB0230 file output ***'
        log_warning()
    # end
    if (get_filename.ADT == False):
        print '*** WARNING No ADT0127 file output ***'
        log_warning()
    # end
    if (get_filename.SST == False):
        print '*** WARNING No SST0126 file output ***'
        log_warning()
    # end
    if (get_filename.SMT == False):
        print '*** WARNING No SMT0125 file output ***'
        log_warning()
    # end
    if (get_filename.SSB == False):
        print '*** WARNING No SSB0229 file output ***'
        log_warning()
    # end
    if (get_filename.SMB == False):
        print '*** WARNING No SMB0223 file output ***'
        log_warning()
    # end
    if (get_filename.DD == False):
        print '*** WARNING No DD0124 file output ***'
        log_warning()
    # end
    if (get_filename.SPB == False): 
        print '*** WARNING No SPB0223 file output ***'
        log_warning()
    # end
    if (get_filename.SPT == False):
        print '*** WARNING No SPT0123 file output ***'
        log_warning()
    # end    
# end def


def get_OCR_dir(exe_OCR):
    """
    Get the path to the correct directory in which to perform OCR

    @param[in]  exe_OCR:            Whether to use the .exe file for OCR (bool)
    @return     (string)            The full path to the OCR directory.
    """    
    
    # is this to be run an an executable?
    if exe_OCR:
        return 'C:\\Pumpkin\\Altium_docs\\OCR'
        
    else:
        return 'C:\\Python27\\Lib\\site-packages\\pypdfocr'
    # end if
# end def


def test():
    """
    Test code for this module.
    """
    
    if not Altium_helpers.clear_output(os.getcwd() + '\\test folder', False):
        log_error()
    # end if
    
    perform_Altium_OCR(False, os.getcwd() + '\\test folder', 8)
    
    print [log_error(get=True), log_warning(get=True)]
    
    #no_errors = True
    #no_warnings = True
    #perform_Altium_OCR(no_errors, no_warnings, True, os.getcwd() + '\\test folder', 8)
    
    #print [no_errors, no_warnings]    
    
#end def

if __name__ == '__main__':
    # if this code is not running as an imported module run test code
    test()
# end if