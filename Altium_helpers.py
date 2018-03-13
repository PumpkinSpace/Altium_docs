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
import shutil
import datetime
import Altium_Files

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
    
        
def clear_output(starting_dir):
    """
    Function to delete all remnants of previously executed iterations of this
    code to create a clean slate.

    @param[in]    starting_dir:    The Altium project directory (full path) 
                                   (string).
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
    min_time = modified_dates[0]
    max_time = modified_dates[0]
    
    # iterate through the dates storing max and min values
    for time in modified_dates:
        # ignore missing dates
        if time != None:
            if time < min_time:
                # new minimum time
                min_time = time
                
            elif time > max_time:
                # new maximum time
                max_time = time
            # end if
        # end if
    # end for
    
    # detect old files
    if ((max_time - min_time) > 600):
        # there is more than 10 mins between the oldest and youngest file dates
        early_date = datetime.datetime.fromtimestamp(min_time)
        formatted_date = early_date.strftime('%Y-%m-%d %H:%M:%S')
        print '*** WARNING possibly delivering old files, date ' + \
              formatted_date + ' ***'
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
    print 'Constructing Archive...'
    
    # get the project part number
    part_number = Altium_Files.get_part_number(starting_dir)
    
    # get the path to Andrews directory
    andrews_dir = get_Andrews_dir(starting_dir)
    
    # make the .zip archive
    shutil.make_archive(starting_dir+'\\'+part_number+'_Folder', 
                        'zip', andrews_dir)
    
    # remove the left over folder
    shutil.rmtree(andrews_dir, ignore_errors=True)
    
    # indicate completion
    print '\n*** Directory ' + part_number + '_Folder.zip' + \
          ' has been generated successfully ***'
# end def