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
import time

#
#
# ----------------
# Classes

class Logger(object):
    def __init__(self, filename):
        self.terminal = sys.stdout
        if os.path.isfile(filename):
            os.remove(filename)
        # end if
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

def get_part_number(starting_dir):
    """
    Function to determine the part number for the folders contained in the 
    starting directory.

    @param[in]    starting_dir:    The Altium project directory (full path) 
                                   (string).
    @return       list of strings: The part number prefix, part number and 
                                   part revision for the project.
    """    
    
    # get the file list of the starting directory
    root_file_list = os.listdir(starting_dir)
    
    # find project Outputs folder
    for filename in root_file_list:
        if (('.' not in filename) and
            (filename.startswith('7')) and 
            (not filename.endswith('PD'))):
            part_prefix = filename[0:3]
            part_number = filename[4:9]
            part_revision = filename[9:11]
    
            return [part_prefix, part_number, part_revision]
        # end
    # end    
    
    # if this code was reached, then no folder was found
    print('***  Error: Folder structure not compliant with current Outjob file   ***\n\n')
    return None
# end def    

def get_assy_number(starting_dir):
    """
    Function to determine the assembly number for the folders contained in the 
    starting directory.

    @param[in]    starting_dir:    The Altium project directory (full path) 
                                   (string).
    @return       list of strings: The assembly number prefix, assembly number and 
                                   asssembly revision for the project.
    """        
    
    # get the file list of the starting directory
    root_file_list = os.listdir(starting_dir)
    
    # find project Outputs folder
    for filename in root_file_list:
        if (('.' not in filename) and
            (filename.startswith('7')) and 
            (filename.endswith('PD'))):
            assy_prefix = filename[0:3]
            assy_number = filename[4:9]
            assy_revision = filename[9:11]
    
            return [assy_prefix, assy_number, assy_revision]
        # end
    # end    
    
    # if this code was reached, then no folder was found
    print('***  Error: Folder structure not compliant with current Outjob file   ***\n\n')
    return None
# end def   

        
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
            print('***  Error: Could not remove previous ' + \
                  'Andrews Format Folder  ***\n\n')
            return False
        # end try
    # end if
    
    # if there is a test.xlsx file, delete it.
    if os.path.isfile(starting_dir + '\\test.xlsx'):
        try:
            os.remove(starting_dir + '\\test.xlsx')
            
        except:
            print('***  Error: Could not remove previous test.xlsx file  ***\n\n')
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
                print('***  Error: Could not remove previous file  ***\n\n')
                return False 
            # end try                  
        # end if      
    # end for
    
    # remove previous step_temp directory
    if os.path.exists(starting_dir + '\\step_temp'):
        try:
            shutil.rmtree(starting_dir + '\\step_temp')
            
        except:
            print('***  Error: Could not remove previous step directory  ***\n\n')
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
        print('*** WARNING possibly delivering old files ***')
        
        # print all filenames that are old and their dates.
        for date in [d for d in modified_dates if d != None]:
            if (max_time - date.date) > 1200:
                early_date = datetime.datetime.fromtimestamp(date.date)
                formatted_date = early_date.strftime('%Y-%m-%d at %H:%M:%S')                
                print('\t' + date.text + ' modified on ' + formatted_date)
            # end if
        # end for
        
        return False
    # end if
    
    return True
# end def

def construct_root_archive(output_dir, part_number):
    """
    Construct an archive of all of the files to be delivered to Pumpkin.

    @param[in]   output_dir:          The packaging outputs directory (full path) 
                                      (string).   
    @param[in]   part_number:         The part number for the design
                                      (string).                 
    """    
    print('\n\nConstructing Archive...')
    
    zip_filename = output_dir + '\\' + part_number + '_Folder'
    
    # make a folder to put the step file in temporarily
    temp_dir = output_dir + '\\temp'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    # end
    
    # get the list of files in the directory
    file_list = os.listdir(output_dir)
    
    # create temporary directory in which to place all of the files needed
    os.makedirs(temp_dir)    
    
    # copy all files into temp directory
    for filename in file_list:
        if '.' in filename:
            # copy files
            shutil.copy(output_dir +'\\' + filename, temp_dir + '\\' + filename)
            
        else:
            # copy folders
            #shutil.copytree(output_dir +'\\' + filename, temp_dir + '\\' + filename)
            shutil.make_archive(temp_dir +'\\' + filename, 'zip', output_dir +'\\' + filename)
        # end if        
    # end for 
    
    # make archive
    shutil.make_archive(zip_filename, 'zip', temp_dir)
    
    time.sleep(0.5)
    
    try:
        # delete the temp directory
        shutil.rmtree(temp_dir)       
        
    except:
        print('*** WARNING: Could not delete temporary output directory ***')
        log_warning()     
    # end try    
    
    # indicate completion
    print('*** Directory ' + part_number + '_Folder.zip' + \
          ' has been generated successfully ***')
    
    zip_filename = zip_filename + '.zip'
    return zip_filename
# end def