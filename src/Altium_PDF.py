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
@package Altium_PDF.py

Package that reads text from and manages pdfs, and subsequent pdf actions as needed by 
the Altium Documentation module.
"""

__author__ = 'David Wright (david@asteriaec.com)'
__version__ = '0.3.0' #Versioning: http://www.python.org/dev/peps/pep-0386/


#
# -------
# Imports

import os
import sys
sys.path.insert(1, 'src\\')
import shutil
import PyPDF2
import time
import Altium_helpers
import pdfminer
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
from io import BytesIO
# end try

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
    retstr = BytesIO()
    codec = 'utf-8'
    
    # set parameters for analysis
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
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
    # end

    # close all files
    fp.close()
    device.close()
    retstr.close()
    
    # return the text
    return text[0]
# end

def manage_Altium_PDFs(pdf_dir, output_pdf_dir, num_layers, 
                       silence = True):
    """
    Analyse the output PDFs from Altium and then extract the text from 
    them and split and rename the pages accordingly.

    @param[in]  pdf_dir:            The location of the pdf files (full path) 
                                    (string).  
    @param[in]  output_pdf_dir:     The location to move the pdfs to (full path) 
                                    (string).    
    @param[in]  num_layers:         The number of layers in the PCB (int).
    @param[in]  silence:            Whether to silence the output of the OCR 
                                    engine (bool).
    @return     (list of mod_dates) The modification dates of the files used.
    """     

    # correct the filename of the layers pdf if required
    modified_dates = []
    
    file_list = os.listdir(pdf_dir)
    
    print("Moving PDF documents")
    
    layer_count = 0
    
    for filename in file_list:
        modified_dates.append(Altium_helpers.mod_date(os.path.getmtime(pdf_dir + '\\' + filename), 
                                                              filename))        
        if filename.startswith("layers."):
            # split the layers file into its pages and write them to the output
            with open(pdf_dir + '\\layers.pdf', "rb") as layers_file:
                layers_pdf = PyPDF2.PdfFileReader(layers_file)                

                # write each page to a separate pdf file
                for page in range(layers_pdf.numPages):
                    # open the output stream
                    output = PyPDF2.PdfFileWriter()

                    # add the layers page to the output stream.
                    output.addPage(layers_pdf.getPage(page))
                    # format the filename 
                    file_name = output_pdf_dir + '\\ART%02d.pdf' % (page+1)

                    with open(file_name, "wb") as outputStream:
                        # write the page
                        output.write(outputStream)
                        layer_count += 1
                    # end with                     
                # end for
        # end with            
            
        elif (('Check' not in filename) and ('layer' not in filename) and ('MOD' not in filename)):
            shutil.copy(pdf_dir+'\\'+ filename, output_pdf_dir + '//' + filename)
        # end if
    # end for
    
    # check outputs
    # get a list of filenames without extensions
    file_list = [f_name.split('.')[0] for f_name in os.listdir(pdf_dir)]
    
    # Generate warnings for pecuiliar outputs
    if (layer_count != num_layers):
        print('\t*** WARNING wrong number of layers printed ***')
        log_warning()
    # end
    if ("MECHDWG" not in file_list):
        print('\t*** WARNING No MECHDWG file output ***')
        log_warning()
    # end
    if ("ADB0230" not in file_list):
        print('\t*** WARNING No ADB0230 file output ***')
        log_warning()
    # end
    if ("ADT0127" not in file_list):
        print('\t*** WARNING No ADT0127 file output ***')
        log_warning()
    # end
    if ("SST0126" not in file_list):
        print('\t*** WARNING No SST0126 file output ***')
        log_warning()
    # end
    if ("SMT0125" not in file_list):
        print('\t*** WARNING No SMT0125 file output ***')
        log_warning()
    # end
    if ("SSB0229" not in file_list):
        print('\t*** WARNING No SSB0229 file output ***')
        log_warning()
    # end
    if ("SMB0223" not in file_list):
        print('\t*** WARNING No SMB0223 file output ***')
        log_warning()
    # end
    if ("DD0124" not in file_list):
        print('\t*** WARNING No DD0124 file output ***')
        log_warning()
    # end
    if ("SPB0223" not in file_list): 
        print('\t*** WARNING No SPB0223 file output ***')
        log_warning()
    # end
    if ("SPT0123" not in file_list):
        print('\t*** WARNING No SPT0123 file output ***')
        log_warning()
    # end      

    print('Complete!\n')

    return modified_dates
# end def  
    
def check_DRC(pdf_dir):
    """
    Checks the design rule check output PDF to see if there are any errors

    @param:    pdf_dir        The full path of the Altium pdf Folder (string).
    @return:   (mod_date)     The modification date of the Design Rule Check
    """  
    print('\nChecking the Design Rule Check...')
    
    DRC_text = ''
    
    if os.path.isdir(pdf_dir):
        # get the file list of the root directory
        file_list = os.listdir(pdf_dir)   
        
        if 'Design Rules Check.PDF' not in file_list:
            print('*** Error: No design rule check has been completed ***')
            log_error()
            return None
        # end if
        
        # get the modification date of the file
        DRC_date = Altium_helpers.mod_date(os.path.getmtime(pdf_dir+'\\Design Rules Check.PDF'), 
                                           'Design Rules Check.PDF')
        
        # extract text and remove whitespace
        DRC_text = "".join(str(convert_pdf_to_txt(pdf_dir+'\\Design Rules Check.PDF')).split())        
        
    else:
        print('***  Error: Folder structure not compliant with current Outjob file   ***\n\n')
        return None     
    # end if
    
    if 'Warnings0' not in DRC_text:
        print('*** Warning: Warnings were raised during Altium DRC ***')
        log_warning()
    # end if
    
    if 'Violations0' not in DRC_text:
        print('*** Warning: Rule Violations were found during Altium DRC ***')
        log_warning()
    # end if
    
    print('Complete!')
    
    return DRC_date
# end def


def check_ERC(pdf_dir):
    """
    Checks the electrical rule check output PDF to see if there are any errors

    @param:    pdf_dir        The full path of the Altium pdf Folder (string).
    @return:   (mod_date)     The modification date of the Electrical Rule Check
    """  
    print('\nChecking the Electrical Rule Check...')
    DRC_text = ''
    
    if os.path.isdir(pdf_dir):
        # get the file list of the root directory
        file_list = os.listdir(pdf_dir)   
        
        if 'Electrical Rules Check.PDF' not in file_list:
            print('*** Error: No electrical rule check has been completed ***')
            log_error()
            return None
        # end if
        
        # get the modification date of the file
        ERC_date = Altium_helpers.mod_date(os.path.getmtime(pdf_dir +'\\Electrical Rules Check.PDF'), 
                                           'Electrical Rules Check.PDF')
        
        # extract text and remove whitespace
        ERC_text = "".join(str(convert_pdf_to_txt(pdf_dir+'\\Electrical Rules Check.PDF')).split())       
        
    else:
        print('***  Error: Folder structure not compliant with current Outjob file   ***\n\n')
        log_error()
        return None  
    # end if
    
    if 'Warning' in ERC_text:
        print('*** Warning: Warnings were raised during Altium ERC ***')
        log_warning()
    # end if
    
    if 'Error' in ERC_text:
        print('*** Warning: Rule Violations were found during Altium ERC ***')
        log_warning()
    # end if
    
    print('Complete!\n')
    
    return ERC_date
# end def


#
# ----------------
# Private Functions
    

def test():
    """
    Test code for this module.
    """
    
    if not Altium_helpers.clear_output(os.getcwd() + '\\test folder', False):
        log_error()
    # end if
    
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