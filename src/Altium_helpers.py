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
@package Altium_helpers.py

Package that contains helper functions for the Altium Documentation module
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
import datetime
import Altium_Files
import Altium_OCR

#
#
# ----------------
# Classes

class Logger(object):
    def __init__(self, filename):
        self.terminal = sys.stdout
        self.log = open(filename, "w")
        # write a header to the log file
        self.log.write("=========   Deliverable Log   ===========\n")
        project_string = "Project: \t\t" + filename.split('\\')[-2] + '\n'
        self.log.write(project_string)
        time_string = "Created at: \t" + datetime.datetime.now().strftime("%Y-%m-%d %H:%M") + '\n\n'
        self.log.write(time_string)
    # end def

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)  
    # end def

    def flush(self):
        #this flush method is needed for python 3 compatibility.
        #this handles the flush command by doing nothing.
        #you might want to specify some extra behavior here.
        pass    
    # end def
    
    def close(self):
        self.log.close()
# end class


class mod_date:
    """ 
    Class to store a files modification information.
    
    @attribute     date:     The modification date of the file (datetime).
    @attribute     text:     Text associated with this modification date 
                             (string).
    """
    def __init__(self, modified_date, filename):
        """
        Initialise the mod_date class
        
        @param[in]     modified_date:   The date the file was last modified 
                                        (datetime).
        @param[in]     filename:        Test associated with this file, usually 
                                        the filename (string).
        """
        self.date = modified_date
        self.text = filename
    # end def
# end class


#
# ----------------
# Public Functions 

def get_output_dir(starting_dir):
    """
    Function to get the path to the Project outputs directory of an Altium 
    Project given the path of the project.

    @param[in]    starting_dir:    The Altium project directory (full path) 
                                   (string).
    @return       (string)         The Full path of the Altium project outputs 
                                   folder for the project being delivered, 
                                   =None if no folder is found.
    """    
    
    # get the file list of the starting directory
    root_file_list = os.listdir(starting_dir)
    
    # find project Outputs folder
    for filename in root_file_list:
        if filename.startswith('Project Outputs'):
            return starting_dir + '\\' + filename
        # end
    # end
    
    # if this code was reached, then no folder was found
    print '***  Error: No Project Outputs Directory Found   ***\n\n'
    return None
# end def 


def get_Andrews_dir(starting_dir):
    """
    Function to get the path to the Andrews Format directory that has been 
    created to temporarly house all of the files to be delivered while this 
    module is running. 
    This function will create said directory if it does not already exist.

    @param[in]    starting_dir:    The Altium project directory (full path) 
                                   (string).
    @return       (string)         The Full path of the Andrews Format folder 
                                   =None if no folder is found.
    """   
    
    # define the string of the full path to the folder
    andrews_dir = starting_dir + '\\Andrews Format'
    
    # if the path does not already exist, create it
    if not os.path.exists(andrews_dir):
        os.makedirs(andrews_dir)
    # end if
    
    # return the full path
    return andrews_dir
#end def


def get_pdf_dir(starting_dir):
    """
    Function to get the path to the pdf directory that has been 
    created to temporarly house all of the pdfs to be delivered while this 
    module is running. 
    This function will create said directory if it does not already exist.

    @param[in]    starting_dir:    The Altium project directory (full path) 
                                   (string).
    @return       (string)         The Full path of the pdfs folder 
    """       
    
    # get the Andrews directory path
    andrews_dir = get_Andrews_dir(starting_dir)
    
    # define the string for the pdfs folder
    pdf_dir = andrews_dir+'\\'+'pdfs'
    
    # if the path does not yet exist, create it
    if not os.path.exists(pdf_dir):
        os.makedirs(pdf_dir)
    # end if
    
    # return the full path
    return pdf_dir
# end def


def get_gerbers_dir(starting_dir):
    """
    Function to get the path to the gerbers directory that has been 
    created to temporarly house all of the gerbers to be delivered while this 
    module is running. 
    This function will create said directory if it does not already exist.

    @param[in]    starting_dir:    The Altium project directory (full path) 
                                   (string).
    @return       (string)         The Full path of the gerbers folder 
    """       
    
    # retrieve the path to the Andrews directory
    andrews_dir = get_Andrews_dir(starting_dir)
    
    # define the path fror the gerbers folder
    gerbers_dir = andrews_dir + '\\Gerbers'
    
    # if the path does not already exist, create it
    if not os.path.exists(gerbers_dir):
        os.makedirs(gerbers_dir)
    # end if
    
    # return the full path
    return gerbers_dir
#end def
    
        
def clear_output(starting_dir, exe_OCR):
    """
    Function to delete all remnants of previously executed iterations of this
    code to create a clean slate.

    @param[in]    starting_dir:    The Altium project directory (full path) 
                                   (string).
    @param[in]    exe_OCR:         Whether or not to use the .exe for OCR (bool).
    @return       (bool)           True  = no errors occurred in this function,
                                   False = errors occurred in this function
    """       
    
    # get Andrews directory path
    andrews_dir = starting_dir + '\\Andrews Format'
    
    # If that path exits, delete it
    if os.path.exists(andrews_dir):
        # attempt to delete the directory
        try:
            shutil.rmtree(andrews_dir)
            
        except:
            print '***  Error: Could not remove previous ' + \
                  'Andrews Format Folder  ***\n\n'
            return False
        # end try
    # end if
    
    # if there is a test.xlsx file, delete it.
    if os.path.isfile(starting_dir + '\\test.xlsx'):
        try:
            os.remove(starting_dir + '\\test.xlsx')
            
        except:
            print '***  Error: Could not remove previous test.xlsx file  ***\n\n'
            return False 
        # end try  
    # end if
    
    # delete any files in the starting sirectory that are .zip files or are 
    # the result of previous executions of OCR.
    for filename in os.listdir(starting_dir):
        if filename.endswith('.zip') or filename.endswith('_ocr.pdf'):
            try:
                os.remove(starting_dir + '\\' + filename)
                
            except:
                print '***  Error: Could not remove previous file  ***\n\n'
                return False 
            # end try                  
        # end if      
    # end for
    
    # remove previous step_temp directory
    if os.path.exists(starting_dir + '\\step_temp'):
        try:
            shutil.rmtree(starting_dir + '\\step_temp')
            
        except:
            print '***  Error: Could not remove previous step directory  ***\n\n'
            return False           
        # end try
    # end if
    
    # get file list of the ocr directory
    ocr_dir = Altium_OCR.get_OCR_dir(exe_OCR)
    ocr_list = os.listdir(ocr_dir)
    
    # search through the files for previous ocr inputs and outputs
    for filename in ocr_list:
        if filename.endswith('.pdf'):
            # this is a previous file so delete it
            try:
                os.remove(ocr_dir + '//' + filename)
                
            except:
                print '***  Error: Could not remove previous ocr_files  ***\n\n'
                return False       
            # end try
        # end if
    # end for
    
    # If this code is reached then no errors occurred
    return True
# end def


def check_modified_dates(modified_dates): 
    """
    Check a lit of dates to see if some files are too old and are therefore
    potentially a risk to the deliverable.

    @param[in]    modified_dates:  List of the modified dates of all the files 
                                   delivered (list of datetimes).
    @return       (bool)           True  = no errors occurred in this function,
                                   False = errors occurred in this function
                                   =None if no folder is found.    
    """
    
    # initialise the limiting dates
    min_time = 0
    max_time = 0
    
    # iterate through the dates storing max and min values
    for date in modified_dates:
        # ignore missing dates
        if date != None:
            if (min_time == 0):
                min_time = date.date
                max_time = date.date      

            elif date.date < min_time:
                # new minimum time
                min_time = date.date
                
            elif date.date > max_time:
                # new maximum time
                max_time = date.date
            # end if
        # end if
    # end for
    
    # detect old files
    if ((max_time - min_time) > 1200):
        # there is more than 10 mins between the oldest and youngest file dates
        print '*** WARNING possibly delivering old files ***'
        
        # print all filenames that are old and their dates.
        for date in [d for d in modified_dates if d != None]:
            if (max_time - date.date) > 1200:
                early_date = datetime.datetime.fromtimestamp(date.date)
                formatted_date = early_date.strftime('%Y-%m-%d at %H:%M:%S')                
                print '\t' + date.text + ' modified on ' + formatted_date
            # end if
        # end for
        
        return False
    # end if
    
    return True
# end def

def construct_root_archive(starting_dir):
    """
    Construct an archive of all of the files to be delivered to Pumpkin.

    @param[in]    starting_dir:    The Altium project directory (full path) 
                                   (string).   
    """    
    print '\n\nConstructing Archive...'
    
    # get the project part number
    part_number = Altium_Files.get_part_number(starting_dir)
    
    # get the path to Andrews directory
    andrews_dir = get_Andrews_dir(starting_dir)
    
    zip_filename = part_number + '_Folder'
    
    # make the .zip archive
    shutil.make_archive(starting_dir+'\\'+zip_filename, 
                        'zip', andrews_dir)
    
    # remove the left over folder
    shutil.rmtree(andrews_dir, ignore_errors=True)
    
    # indicate completion
    print '*** Directory ' + part_number + '_Folder.zip' + \
          ' has been generated successfully ***'
    
    zip_filename = zip_filename + '.zip'
    return zip_filename
# end def