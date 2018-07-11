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
@package Altium_Files.py

Package that manages files in the Altium Documentation module
"""

__author__ = 'David Wright (david@asteriaec.com)'
__version__ = '0.2.0' #Versioning: http://www.python.org/dev/peps/pep-0386/


#
# -------
# Imports

import os
import sys
sys.path.insert(1, 'src\\')
import Altium_Excel
import Altium_helpers
import shutil
import pyPdf
import re
import Altium_OCR
import threading
import multiprocessing
import time

#
# -------
# Constants

# desired file extensions
altium_ext = ['Outjob', 'cam', 'PrjPcb', 'PrjPcbStructure', \
              'PcbDoc', 'SchDoc', 'Harness', 'Outjob']

# rejected file extensions
bad_gerber_ext = ['zip', 'ods', 'xlsx', 'Report.Txt', '2.txt', \
                  '4.txt', '6.txt', '8.txt', 'drc', 'html', '~lock', '_Previews']

# file extensions that represent layers
layer_gerber_list = ['.GTL', '.GBL', '.G1', '.G2', '.G3', '.G4', '.G5', \
                     '.G6', '.GP1', '.GP2', '.GP3', '.GP4', '.GP5', '.GP6']

required_gerber_list = ['.apr', '.DRR', '.EXTREP', '.GBL', '.GBO', '.GBP', '.GBS',
                        '.GKO', '.GTL', '.GTO', '.GTP', '.GTS', '.LDP', '.RUL',
                        '.REP', '.APR_LIB', '.xls']

# Dictionary of lines that can be added to Readme file
Readme_dictionary = {'X': '.X                      Dielectric X file                               Gerber\n',
                     'apr': '.apr    		Aperture file           			Gerber\n',
                     'DRR': '.DRR     		Drill file            				ASCII\n',
                     'EXTREP': '.EXTREP     		Layer Information file				ASCII\n',
                     'GX': '.GX			Artwork file for internal signal Layer X	Gerber\n',
                     'GBL': '.GBL			Artwork File for Layer X			Gerber\n',
                     'GBO': '.GBO    		Silkscreen files for layer X			Gerber\n',
                     'GBP': '.GBP     		SMD paste files for layer X     		Gerber\n',
                     'GBS': '.GBS     		Soldermask files for layer X    		Gerber\n',
                     'GKO': '.GKO            	Keepout file					Gerber\n',
                     'GMX': '.GMX     		Mechanical Layer X file 	   		Gerber\n',
                     'GTL': '.GTL            	Artwork File for Layer 1			Gerber\n',
                     'GML': '.GML            	Altium .GML file				Gerber\n',
                     'GTO': '.GTO			Silkscreen files for layer 1			Gerber\n',
                     'GTP': '.GTP			SMD paste files for layer 1     		Gerber\n',
                     'GTS': '.GTS			Soldermask files for layer 1    		Gerber\n',
                     'GPX': '.GPX			Artwork file for internal plane Layer X		Gerber\n',
                     'LDP': '.LDP			Altium .LDP file\n',
                     'REP': '.REP			Altium Report file				ASCII\n',
                     'RUL': '.RUL			Altium .RUL file\n',		
                     'TXT': '.TXT			Altium Drill file				ASCII\n',
                     'txt': '.txt                       Altium Pick and Place file                      ASCII\n',
                     'csv': '.CSV			Pick and Place File				ASCII\n',
                     'APR_LIB': '.APR_LIB	        Altium Aperture Library 		        Gerber\n',
                     'xls': '.xls			BOM File					ASCII\n'}


# List in which the readme file gets built, pre loaded with header              
Readme_lines = ['================== PCB Fabrication Information ==================================\n',
                'PUMPKIN, Inc.         750 Naples Street     San Francisco, CA 94112\n',
                'tel (415) 584-6360    fax (415) 585-7948    web http://www.pumpkininc.com\n',
                '== PCB Type =====================================================================\n',
                'Mixed through-hole and surface-mount components\n',
                '== PCB Layout Package Used ======================================================\n',
                'Altium Desgner 16\n',
                '== Layers =======================================================================\n',
                'Layer 1\t\t\tTop of board (component side)\n',
                '== File Extensions ==== Description ====================================Format ==\n']

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

# set the initial value
log_warning.no_warnings = True


def get_part_number(starting_dir):
    """
    Function to extract the part number form the starting directory folder name.

    @param[in]   starting_dir:     The Altium project directory (full path) 
                                   (string).
    @return      (string)          The extracted part number.
    """       
    
    # strip the part number out and return it.
    part_list = starting_dir.split('(')
    return part_list[1][:-1]    
# end def

def set_assembly_number(starting_dir):
    """
    Function to prompt the user for the assembly number and then store it.

    @param[in]   starting_dir:     The Altium project directory (full path) 
                                   (string).
    """       
    
    not_valid = 1
    while (not_valid == 1):
        assy_number = raw_input("Please enter the project assembly Number : 710-")
        
        if (len(assy_number) == 5) and assy_number.isdigit():
            set_assembly_number.assy_number = assy_number
            not_valid = 0
        
        else:
            print "that is not a valid Assembly Number!"
        # end if 
    # end while
# end def

# set the initial value
set_assembly_number.assy_number = None

def get_assembly_number(starting_dir):
    """
    Function to prompt the user for the assembly number and then store it.

    @param[in]   starting_dir:     The Altium project directory (full path) 
                                   (string).
    """       
    
    if (set_assembly_number.assy_number == None):
        print "*** Error: No Assembly Number has been set ***"
        log_error()
        
    else:
        return set_assembly_number.assy_number
    # end if 
# end def


def move_Altium_files(starting_dir):
    """
    Function to move all of the altium files to the deliverable directory.

    @param[in]   starting_dir:        The Altium project directory (full path) 
                                      (string).
    @return      (list of mod_dates)  List of the modification dates of the 
                                      Altium files
    """       
    print 'Moving Altium Files...'
    
    # initialise dates list
    modified_dates = []
    
    # get the deliverable directory
    andrews_dir = Altium_helpers.get_Andrews_dir(starting_dir)
    
    # create directory for the altium files
    altium_dir = andrews_dir + '\Altium Files'
    os.makedirs(altium_dir)
    
    # get file list of root directory
    root_file_list = os.listdir(starting_dir)
    
    # copy desired files
    for filename in root_file_list:
        # compare each filename with the desired set of extensions
        for ext in altium_ext:
            if filename.endswith(ext):
                # This file is desired to copy it across
                try:
                    shutil.copyfile(starting_dir+'\\'+filename, altium_dir+'\\'+filename)  
                    
                except:
                    print '*** Error: could not move ' + filename + ' ***'
                    log_error()
                # end try
                
                if (ext == 'PrjPcb'):
                    # this is the project file which creates the pdf filename
                    # so move the similarly named pdf file
                    pdf_filename = filename.split('.')[0] + '.pdf'
                    
                    try:
                        shutil.copyfile(starting_dir+'\\'+pdf_filename, altium_dir+'\\'+pdf_filename)  
                        
                    except:
                        print '*** Error: could not move ' + pdf_filename + ' ***'
                        log_error()
                    # end try   
                # end if
                
                modified_dates.append(Altium_helpers.mod_date(os.path.getmtime(starting_dir+'\\'+filename),
                                                               filename))
            # end if
        # end for
    # end for
    
    # zip the resulting folder and remove the temporary directory
    try:
        shutil.make_archive(andrews_dir+'\\Altium Files', 'zip', altium_dir)  
        
    except:
        print '*** Error: could not create Altium files.zip ***'
        log_error()
    # end try
    
    try:
        shutil.rmtree(altium_dir)     
        
    except:
        print '*** Error: could not remove Altium Files folder ***'
        log_error()
    # end try
    
    
    print 'Complete! \n'    
    
    return modified_dates
# end def

                
def move_gerbers(starting_dir):
    """
    Function to move all of the gerber files to the deliverable directory.

    @param[in]   starting_dir:        The Altium project directory (full path) 
                                      (string).
    @return      (list of mod_dates)  List of the modification dates of the 
                                      Gerbers.
    @return      (int)                The number of layers in the design.
    """    
    print 'Moving Gerber Files...'
    
    # initialiase list of modified dates to return
    modified_dates = []
    
    # get the project outputs directory
    outputs_dir = Altium_helpers.get_output_dir(starting_dir)
    
    # make gerber directory
    gerbers_dir = Altium_helpers.get_gerbers_dir(starting_dir)
    
    # get andrews directory
    andrews_dir = Altium_helpers.get_Andrews_dir(starting_dir)
    
    # get list of gerber files
    gerber_file_list = os.listdir(outputs_dir)
    
    # if there are simly too few gerber files to be acceptible
    if len(gerber_file_list) < 10:
        # No gerbers exist
        print '***   Error: No Gerbers have been generated   ***\n\n'
        log_error()
        return None, None
    # end 
    
    # layer counter
    layers = 0
    
    # iterate through the files in the gerber directory
    for filename in gerber_file_list:
        good_filename = True
        
        # see if the selected gerber is a layer artwork file
        for ext in layer_gerber_list:
            if filename.endswith(ext):
                # is a valid gerber layer
                layers+=1
    
            if filename.endswith('G7') or filename.endswith('GP7'):
                # This indicates that there are more layers than this code was 
                # developed to handle
                print '***  Error: Script needs to be extended to ' + \
                      'handle this many layers   ***\n\n'
                log_error()
            # end
        # end
        
        # see if the file is one of the gerber files to be ignored
        for ext in bad_gerber_ext:
            if (filename.endswith(ext) and ("Pick Place" not in filename)):
                good_filename = False
            # end if
        # end for
        
        # if the filename is desired
        if good_filename:
            if filename.endswith('.xls'):
                if ('DNP' not in filename):
                    # this is the BOM file
                    
                    # get the part number for the project
                    part_number = get_part_number(starting_dir)
                    
                    # copy the bom to the deliverable
                    shutil.copyfile(outputs_dir + '\\' + filename, \
                                            gerbers_dir + '\\' + \
                                            part_number + '_BOM.xls')
                    
                    # get it's modification date
                    modified_dates.append(Altium_helpers.mod_date(os.path.getmtime(outputs_dir + '\\' +\
                                                           filename), filename))    
                # end if
            
            else:
                # attempt to copy the gerber file to the deliverables
                try:
                    shutil.copyfile(outputs_dir+'\\'+filename, 
                                    gerbers_dir+'\\'+filename)
                    
                except:
                    print '*** Error: could not move ' + filename + ' ***'
                    log_error()
                # end try
                
                # get the modification date of the file
                modified_dates.append(Altium_helpers.mod_date(os.path.getmtime(outputs_dir+'\\'+ \
                                                       filename), filename))     
            # end if
        # end if
    # end for
    
    # create the readme for the gerbers directory
    create_readme(starting_dir, layers)
    
    # check that all required gerbers are in the directory
    check_gerber_folder(gerbers_dir)
    
    # get the part number for the project
    part_number = get_part_number(starting_dir)
    
    # move all gerbers into the desired zip archive
    try:
        shutil.make_archive(andrews_dir+'\\'+part_number, 'zip', gerbers_dir)
        shutil.rmtree(gerbers_dir, ignore_errors=True) 
    except:
        print '*** Error: could not create gerber.zip archive ***'
        log_error()
    # end
    
    print 'Complete! \n'  
    
    return modified_dates, layers
# end def


def move_documents(starting_dir, exe_OCR, layers):
    """
    Function to move all the documents to the deliverable directory.

    @param[in]   starting_dir:        The Altium project directory (full path) 
                                      (string).
    @param[in]   exe_OCR:             Whether to use the .exe file for OCR 
                                      conversion (bool).
    @param[in]   layers:              The number of layers in the PCB (int).
    @return      (list of mod_dates)  Modification dates of the documents.
    """     
    
    # find the Schematic and BOM documents
    no_schematic = True
    
    # create required directories
    andrews_dir = Altium_helpers.get_Andrews_dir(starting_dir)
    pdf_dir = Altium_helpers.get_pdf_dir(starting_dir)
    
    # move the BOM
    # this is now done by move_gerbers()
    # modified_dates = move_bom(starting_dir)
    
    # manage the schematic document
    # modified_dates = [manage_schematic(starting_dir)]
    modified_dates = [manage_schematic(starting_dir, with_threads = True)]
    
    # construct the assembly doc
    modified_dates.extend(Altium_Excel.construct_assembly_doc(starting_dir))
    
    # move the xps file
    modified_dates.append(move_xps(starting_dir))
    
    # get the file list for the starting directory
    root_file_list = os.listdir(starting_dir)
    
    # get the part number
    part_number = get_part_number(starting_dir)
    
    # search for ASSY_REV document in root folder
    for filename in root_file_list:
        if (filename.endswith('xlsx')) and ('ASSY' in filename):
            # ASSY REV doc found
            try:
                shutil.copyfile(starting_dir+'\\'+filename, \
                                pdf_dir + '\\' + part_number + '_ASSY_REV.xlsx')
                
            except:
                print('***   Error: could not move ASSY_REV document   ***')
                log_error()     
            # end try
            
            break
        # end   
    # end    
    
    # perform OCR on the layers pdf file   
    #modified_dates.extend(Altium_OCR.perform_Altium_OCR(exe_OCR, starting_dir, layers))
    modified_dates.extend(Altium_OCR.perform_Altium_OCR(exe_OCR, starting_dir, layers, with_threads=True))
    
    
    # check for errors and warnings following OCR
    if not (Altium_OCR.log_error(get=True) and Altium_Excel.log_error(get=True)):
        log_error()
    # end if
    
    if not (Altium_OCR.log_warning(get=True) and Altium_Excel.log_warning(get=True)):
        log_warning()
    # end if   
    
    # make the pdf archive
    shutil.make_archive(andrews_dir+'\\'+get_part_number(starting_dir)+ \
                        'PD', 'zip', pdf_dir)
    
    # remove the temp directory
    shutil.rmtree(pdf_dir)     
    
    return modified_dates
# end def


def zip_step_file(starting_dir):
    """
    Function to move the step file to the deliverable directory.

    @param[in]   starting_dir:        The Altium project directory (full path) 
                                      (string).
    @return      (datetime)           Modification date of the step file.
    """      
    
    print 'Zipping Step File...'
    
    # completion flags
    step_file_found = False
    modified_date = None
    
    # get andrews directory
    andrews_dir = Altium_helpers.get_Andrews_dir(starting_dir)
    
    # get the file list of the starting directory
    root_file_list = os.listdir(starting_dir)
    
    # search for step file
    for filename in root_file_list:
        if filename.endswith('.step'):
            # step file has been found
            
            # get it's modification date
            modified_date = Altium_helpers.mod_date(os.path.getmtime(starting_dir+'\\'+filename), 
                                                    filename)
            
            # Shrinking the step file would happen here....
            
            # make a folder to put the step file in temporarily
            step_dir = starting_dir + '\\step_temp'
            if os.path.exists(step_dir):
                shutil.rmtree(step_dir)
            # end
            
            # create temporary directory in which to place all of the files needed
            os.makedirs(step_dir)    
            
            # get the part number
            part_number = get_part_number(starting_dir)
            
            # copy file into temp directory
            shutil.copy(starting_dir+'\\'+filename, step_dir +'\\'+part_number+'.step')
            
            # make archive
            shutil.make_archive(andrews_dir+'\\'+part_number+'_step', 'zip', step_dir)
            
            # delete the temp directory
            shutil.rmtree(step_dir, ignore_errors=True)        
            
            step_file_found = True
        # end if
    # end for
    
    if not step_file_found:
        print '*** WARNING: No step file found ***'
        log_warning()
    # end if
            
    print 'Complete! \n'
    
    return modified_date
# end def


#
# ----------------
# Private Functions 

def move_xps(starting_dir):
    """
    Function to move the xps file to the deliverable directory.

    @param[in]   starting_dir:        The Altium project directory (full path) 
                                      (string).
    @return      (datetime)           Modification date of the xps file.
    """      
    
    # get the file list of the root directory
    root_file_list = os.listdir(starting_dir)
    
    # get andres direstory
    andrews_dir = Altium_helpers.get_Andrews_dir(starting_dir)
    
    # initialise variables
    xps_file = ''
    modified_date = None
    
    # search through the file list fot the xps file
    for filename in root_file_list:
        if filename.endswith('xps'):
            # store the filename
            xps_file = filename
            
            # store it's modification date
            modified_date = Altium_helpers.mod_date(os.path.getmtime(starting_dir+'\\'+filename), 
                                                    filename)
        # end if
    # end for
    
    if xps_file == '':
        print '*** Error: no .xps file found ***'
        log_error()
        return None
    # end if
    
    # get the extension opf the xps file (incase it is .oxps)
    xps_ext = xps_file.split('.')[1]
    
    # define the xps directory
    xps_dir = andrews_dir + '//xps_temp'
    
    # create the xps directory if required
    if not os.path.exists(xps_dir):
        os.makedirs(xps_dir)
    # end if
    
    # get the part number
    part_number = get_part_number(starting_dir)
    
    # copy file into temp directory
    shutil.copy(starting_dir+'\\'+xps_file, xps_dir +'\\'+part_number + '.' + xps_ext)

    try:
        # make archive
        shutil.make_archive(andrews_dir+'\\'+part_number+'_xps', 'zip', xps_dir)
        
    except:
        print('***   Error: could not zip .xps file ***')
        log_error()     
    # end try        
            
    try:
        # delete temp directory
        shutil.rmtree(xps_dir, ignore_errors=True)     
    
    except:
        print('***   Error: could not remove temporary xps folder file ***')
        log_error()     
    # end try        
    
    return modified_date
# end def


def manage_schematic(starting_dir, with_threads = False):
    """
    Function to move the schematic to the deliverable directory.

    @param[in]   starting_dir:        The Altium project directory (full path) 
                                      (string).
    @param[in]   with_threads:        Use threads for this process (bool).
    @return      (mod_date)           Modification dates of the schematic.
    """     
    print 'Finding Schematic Document...'
    
    # initialise the return value
    modified_date = None
    
    # create the pdf directory
    pdf_dir = Altium_helpers.get_pdf_dir(starting_dir)
    
    # get the file list of the root directory
    root_file_list = os.listdir(starting_dir)
    
    # get the part number
    part_number = get_part_number(starting_dir)    
    
    no_schematic = True    
    # search for a schematic document in the root directory
    for filename in root_file_list:
        if (filename.lower().endswith('pdf') and 
            (len(filename) > 12) and 
            (filename != 'PCB Prints.pdf') and 
            ('layers' not in filename)):
            # schematic found
            no_schematic = False
            modified_date = Altium_helpers.mod_date(os.path.getmtime(starting_dir+'\\'+filename),
                                                    filename)
            break
        # end
    # end
    
    if no_schematic:
        # No schematic was found
        print('***   Error: No Schematic Document was found   ***')
        log_error()
        return None
    # end
    
    print '\tReading the Schematic file...'
    
    # open pdf file and split into pages
    try:
        with open(starting_dir+'\\'+filename, "rb") as schematic_file:
            schematic = pyPdf.PdfFileReader(schematic_file)
            
            # write each page to a separate pdf file
            for page in xrange(schematic.numPages):
                # add page to the output stream
                output = pyPdf.PdfFileWriter()
                output.addPage(schematic.getPage(page))
                # format the filename 
                file_name = pdf_dir + '\\' + part_number + '--' + str(page+1) + '.pdf'
                
                with open(file_name, "wb") as outputStream:
                    # write the page
                    output.write(outputStream)
                # end with
            # end for
        # end with
        
    except:
        print('***   Error: Could not open schematic document   ***')
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
                                             target = rename_sheet, 
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
                
        if any([q == True for q in thread_data]):
            # mod doc found
            found_mod_doc()
        # end if
        
        if any([q == False for q in thread_data]):
            # an error occurred
            log_error()
        # end if
        
    else:
        # rename the pdfs with the correct filenames
        for i in range(1,page+2):
            rename_sheet(starting_dir, i)
        # end for
    # end if
    
    if not found_mod_doc(get=True):
        print('***   Warning: No Modification information found in schematic   ***')
        Altium_Excel.set_assy_options(starting_dir, [], [])
        log_warning()
    # end if    
    
    print 'Complete! \n'    
    
    return modified_date
# end def

def rename_sheet(starting_dir, sheet_number, queue = None):
    """
    Function to rename a schematic sheet based on information contained within 
    it.

    @param[in]   starting_dir:        The Altium project directory (full path) 
                                      (string).
    @param[in]   sheet_number:        The number of the sheet to rename (int)
    @param[out]  queue:               The queue to return data to if running
                                      as a thread (Queue).
    """  
    # get the pdf directory
    pdf_dir = Altium_helpers.get_pdf_dir(starting_dir)
    
    # get the part number
    part_number = get_part_number(starting_dir)
    
    old_file_name = pdf_dir + '\\' + part_number + '--' + \
        str(sheet_number) + '.pdf'
            
    # find page number
    [page_number, is_assy] = get_page_number(old_file_name, part_number, 
                                             starting_dir)
    
    # if this page is the assembly_revision document
    if is_assy:
        if queue == None:
            found_mod_doc()
            
        else:
            queue.put(True)
        # end if
    # end if
        
    # rename the file
    try:
        os.rename(old_file_name, pdf_dir + '\\' + \
                  part_number + '-' + page_number + '.pdf')
        
    except:
        print('***   Error: Could rename pdf document   ***')
        
        if queue == None:
            log_error()
            
        else:
            queue.put(False)
        # end if
    # end try
# end def


def found_mod_doc(get = False):
    """
    Function to indicate that the modification document was found.

    @param[in]    get:        True  = return mod_found
                              False = indicate that the document was found (bool)
    @attribute    mod_found:  Whether the document has been found
    @return       (bool)      True  = The document has been found
                              False = The document has not been found
    """  
    
    # determine which action to take
    if get:
        # return the state
        return found_mod_doc.mod_found
    
    else:
        # Mark as found
        found_mod_doc.mod_found = True
    # end if
# end def

# set the initial value
found_mod_doc.mod_found = False 


def move_bom(starting_dir):
    """
    Function to move the BOM to the deliverable directory.

    @param[in]   starting_dir:        The Altium project directory (full path) 
                                      (string).
    @return      (list of datetimes)  List of the modification dates of the 
                                      BOM
    """     
    print 'Moving BOM...'
    
    # success flag
    no_bom = True
    
    # initialise the return list
    modified_dates = []
    
    # get the project part number
    part_number = get_part_number(starting_dir)
    
    # get the outputs directory
    outputs_dir = Altium_helpers.get_output_dir(starting_dir)
    
    # get list of gerber files
    gerber_file_list = os.listdir(outputs_dir)    
    
    # create the PDF Directory
    pdf_dir = Altium_helpers.get_pdf_dir(starting_dir)    
    
    # search for BOM in project outputs folder
    for filename in gerber_file_list:
        if filename.endswith('xls'):
            # BOM found
            shutil.copyfile(outputs_dir + '\\' + filename, \
                           pdf_dir + '\\' + part_number + '_BOM.xls')
            modified_dates.append(Altium_helpers.mod_date(os.path.getmtime(outputs_dir+'\\'+filename),
                                                          filename))
            
            no_bom = False
            break
        # end if
    # end for
    
    if no_bom:
        # No BOM was found
        print('***   No BOM was found in project outputs   ***')
        log_error()
    # end
    
    print 'Complete! \n'
    
    return modified_dates
# end def


def check_gerber_folder(gerber_dir):
    """
    Function to check the gerbers directory to see if anything is missing.

    @param[in]   gerber_dir:        The directory that the gerbers have been put
                                    into (fullp path) (string).
    """     
    
    # get the file list for the gerbers dir
    file_list = os.listdir(gerber_dir)
    
    # iterate through the list required file extensions
    for ext in required_gerber_list:
        # flag to record whether or not the extension was found
        ext_found = False
        
        # iterate through the files in the directory to see if the extenstion 
        # is present
        for filename in file_list:
            if filename.endswith(ext):
                ext_found = True
                break
            # end if
        # end for
        
        if not ext_found:
            print '*** Error: no ' + ext + ' file output to gerbers ***'
            log_error()
        # end if
    # end for
    
    # check for files that are not differentiable by extension
    readme_found = False
    pick_found = False
    
    # iterate through the file list again
    for filename in file_list:
        if 'README' in filename:
            readme_found = True
            
        elif 'Pick Place' in filename:
            pick_found = True
        # end if
    # end for
    
    if not readme_found:
        print '*** Error: no readme file output to gerbers ***'
        log_error()
    # end if
    
    if not pick_found:
        print '*** Error: no pick and place file output to gerbers ***'
        log_error()
    # end if    
# end def


def create_readme(starting_dir, layers):
    """
    Function to create the readme file for the gerber file delivery.

    @param[in]   starting_dir:        The Altium project directory (full path) 
                                      (string).
    @param[in]   layers:              The number of layers in the PCB (int).
    """     
    print '\tGenerating Readme File...'
    
    # list to store new lines in
    new_readme_lines = []
    
    # list of extensions already used to prevent repeats
    extensions_used = []
    
    # get the output directory
    outputs_dir = Altium_helpers.get_output_dir(starting_dir)
    
    # get the gerbers directory
    gerbers_dir = Altium_helpers.get_gerbers_dir(starting_dir)
    
    # get list of gerber files
    gerber_file_list = os.listdir(outputs_dir)    
    
    # Iterate through every file
    for filename in gerber_file_list:
        good_filename = True
    
        # reject unwanted file extensions
        for ext in bad_gerber_ext:
            if filename.endswith(ext) or ("Pick Place" in filename):
                good_filename = False
            # end if
        # end for
        
        if good_filename:
            # get file extension
            filename_list = filename.split('.')
            extension = filename_list[1]
    
            if len(extension) == 1:
                # is a dielectric file -> number accordingly
                new_line = Readme_dictionary['X']
                new_line2 = new_line.replace('X', extension)
                new_readme_lines.append(new_line2)
    
            elif len(extension) == 2:
                # is a internal signal layer file -> number accordingly
                new_line = Readme_dictionary['GX']
                number = extension[-1]    
                new_line2 = new_line.replace('X', number)
                new_readme_lines.append(new_line2)
    
            elif extension.startswith('GP'):
                # is a internal plane file -> number accordingly
                new_line = Readme_dictionary['GPX']
                number = extension[-1]    
                new_line2 = new_line.replace('X', number)
                new_readme_lines.append(new_line2)   
    
            elif any(char.isdigit() for char in extension):
                # is a mechanical layer file -> number accordingly
                new_line = Readme_dictionary['GMX']
                number = re.findall(r'\d+', extension)   
                new_line2 = new_line.replace('X', number[0])
                new_readme_lines.append(new_line2)
    
            elif extension not in extensions_used:
                # this extension has not been repeated
    
                if (extension != 'TXT') and (extension != 'EXTREP'):
                    # auto numbering will not affect file extension
                    new_line = Readme_dictionary[extension]
                    new_line2 = new_line.replace('X', str(layers))
                    new_readme_lines.append(new_line2)
    
                else:
                    new_readme_lines.append(Readme_dictionary[extension])
                # end if
    
                extensions_used.append(extension)               
            # end if
        # end for
    # end for  
    
    # order gerber file lines
    new_readme_lines.sort()
    
    # add line for the readme file
    new_readme_lines.append('README'+ str(layers) + '.TXT              '+ \
                            'This file                       		ASCII\n')
    
    # add new lines to readme line list
    Readme_lines.extend(new_readme_lines)
    
    # insert layer descriptions into readme
    for layer in range(2,layers+1):
        if layer != layers:
            # Is an internal layer
            Readme_lines.insert(7+layer, 'Layer '+str(layer)\
                                +'\t\t\tInternal Layer '+str(layer-1) + '\n')
    
        else:
            # is the bottom layer
            Readme_lines.insert(7+layer, 'Layer '+str(layer)\
                                +'\t\t\tBottom of board (solder side)\n')
        # end       
    # end    
    
    # open and write to tex file for readme
    try:
        readme_file = open(gerbers_dir+'\\'+'README'+str(layers)+'.TXT', 'w')
        readme_file.writelines(Readme_lines)
        readme_file.close()    
        
    except:
        print '*** could not write readme file ***'
        log_error()
    # end try
    
    print '\tComplete!'
# end def


def get_page_number(path, pn, starting_dir):
    """
    Function to extract the text from a schematic document and assign it the 
    correct page number as read from the document.
    
    @param[in]   path:             The path of the pdf file to be read (string)
    @param[in]   pn:               The part number of the project (string)
    @param[in]   starting_dir:     The Altium project directory (full path) 
                                   (string).
    @return      (string)          The page number of this page
    @return      (bool)            True is this page is the assembly revision 
                                   page. False otherwise
    """        
    
    # extract the text from a pdf page
    pdf_text = Altium_OCR.convert_pdf_to_txt(path)
    
    # check to see if this document is the Assembly revision document
    if 'ASSY' in pdf_text:
        # if it is, extract that information and process it
        extract_assy_info(pdf_text, starting_dir)
        is_assy = True
        
    else:
        is_assy = False
    # end if
    
    page_number = ''
        
    # find the location of the of string that separates the page numers
    of_index = pdf_text.rfind(' of ')
    if of_index == -1:
        # if of is not found return an error
        page_number = 'error'
        
    elif (pdf_text[of_index-2].isdigit() and 
          (pdf_text[of_index-1] != '5') and 
          (pdf_text[of_index-6:of_index-1] != '94112') and 
          (pdf_text[of_index-6:of_index-1] != pn[:-1]) and 
          (pdf_text[of_index-4:of_index-1] != ' 01')):
        # there are two digits in the page number, return both
        page_number = pdf_text[of_index-2:of_index]
    
    else:
        # return the single digit page number
        page_number = pdf_text[of_index-1]
    # end if
    
    if not page_number.isdigit():
        print '*** Error: ' + page_number + ' is not a valid page number ***'
        log_error()
    # end if
    
    return page_number, is_assy
# end def


def extract_assy_info(pdf_text, starting_dir):
    """
    Function to extract the assembly revision information from the text of a 
    schematic sheet.

    @param[in]   pdf_text:         The text of the schematic page
    @param[in]   starting_dir:     The Altium project directory (full path) 
                                   (string).
    """     
    
    # split the text into blocks around the word case
    assy_blocks = pdf_text.split('case')
    
    # if there are not the expected 3 blocks, throw an error
    if len(assy_blocks) != 3:
        print "*** Warning, too many ASSY_REV blocks. ***"
        log_warning()
    # end if
    
    # split the third block into it's parts
    list_1 = assy_blocks[2].split(';')[:-1]
    
    # remove garbage
    while ((assy_blocks[1][0].isalpha() == False) and (assy_blocks[1] != '')):
        assy_blocks[1] = assy_blocks[1][1:]
    # end while
    
    # split the second block into it's parts
    list_0 = assy_blocks[1].split(';')[:-1]
    
    # if the lists are not the same length post a warning
    if len(list_0) != len(list_1):
        print "*** Warning, missmatched ASSY_REV blocks ***"
        log_warning()
        return None
    # end if        
    
    # insert the gathered information into the ASSY_REV document
    if not Altium_Excel.set_assy_options(starting_dir, list_0, list_1):
        log_error()
    # end if
# end def
