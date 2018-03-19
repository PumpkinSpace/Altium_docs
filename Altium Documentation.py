# PCB document backup creater
# Pumpkin Inc.
# David Wright 2018
#
# In order to work the starting directory must contain:
#     A project Outputs Folder contiaining:
#           Full Gerber outputs
#           A BOM in .xls format
#     A pdf of the schematics with a name at least 12 characters long
#     A pdf of all the layers in the design (Default Prints) called layers.pdf
#     Circuit Board design of up to 12 layers (maybe less depending on plane/signal layer breakdown

# protect from subprocessing module
if __name__ == '__main__':
    
    # Required imports
    import os
    import Altium_OCR
    import Altium_Excel
    import Altium_helpers
    import Altium_Files
    
    
    #################### Change this for each implementation #######################
    # directory where the Circuit board files are stored
    starting_dir = 'C:\Users\Asteria\Dropbox\Satellite\Pumpkin PCBs\ADCS Interface Module 2 (01845B)'
    
    # should the executable be used to perform OCR, otherwise use the 
    # installed pypdfocr
    exe_OCR = False
    
    # store the execution directory
    Altium_Excel.set_directory(os.getcwd())
    
    # go to desired working directory
    os.chdir(starting_dir)
    
    # attempt to clear previous files from the directory
    if not Altium_helpers.clear_output(starting_dir, exe_OCR):
        print '*** Error: Previous output could not be deleted ***'
    # end if
    
    # warning racker
    no_warnings = False
    
    # move master ASSY REV document
    Altium_Excel.copy_assy_rev(starting_dir)
    
    # create list to load file modified dates into.
    modified_dates = []
    
    # Move all of the Altium files into their folder
    modified_dates.extend(Altium_Files.move_Altium_files(starting_dir))
    
    # Move the gerber files and create a readme file for them
    [gerber_dates, layers] = Altium_Files.move_gerbers(starting_dir)
    
    # add the gerber modified dates to the list
    modified_dates.extend(gerber_dates)
    
    # move all of the other documents
    modified_dates.extend(Altium_Files.move_documents(starting_dir, 
                                                      exe_OCR, layers))
    
    # zip the step file
    modified_dates.append(Altium_Files.zip_step_file(starting_dir))
    
    # find the oldest and newest files used.
    no_warnings = Altium_helpers.check_modified_dates(modified_dates)
    
    # construct the final zip file and remove un-needed directories
    Altium_helpers.construct_root_archive(starting_dir)
    
    # check for warnings
    if not (no_warnings and 
            Altium_Excel.log_warning(get=True) and 
            Altium_OCR.log_warning(get=True) and 
            Altium_Files.log_warning(get=True)):
        print '\n*** Warnings were raised so please reveiw ***'
    # end if
    
    # check for errors
    if not (Altium_Excel.log_error(get=True) and 
            Altium_OCR.log_error(get=True) and 
            Altium_Files.log_error(get=True)):
        print '\n*** Errors occurred so please reveiw ***'
    # end if
    
# end if
