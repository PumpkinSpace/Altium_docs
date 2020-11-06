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
    import sys
    
    if len(sys.argv) != 3:
        #################### Change this for each implementation #######################
        # directory where the Circuit board files are stored
        starting_dir = 'C:\\Pumpkin\\Altium_docs\\test folder (02190A)'
        
        # should the executable be used to perform OCR, otherwise use the 
        # installed pypdfocr
        exe_OCR = False
        
        src_path = os.getcwd()+'\\src\\'
        sys.path.insert(1, src_path)
        import Altium_Excel 
        import Altium_GS
        
        # store the execution directory
        Altium_Excel.set_directory(os.getcwd())
        
    else:
        # this code is running from the command line
        
        dir_path = '\\'.join(sys.argv[0].split('\\')[:-1])
        sys.path.insert(1, dir_path + '\\src\\')
        import Altium_Excel    
        import Altium_GS
        
        # the first argument is the full path of the script
        Altium_Excel.set_directory(dir_path)
        
        # the second argument is the directory this had been called from
        starting_dir = sys.argv.pop(1)
        
        # check for additional depricated elements from old code to clear the buffer
        try:
            sys.argv.pop(1)
        except:
            pass
        # end try
    # end if
    
    import Altium_PDF
    import Altium_helpers
    import Altium_Files  
    import time
    import zipfile
    import shutil
    
    # direct all output to a log file as well
    log_filename = starting_dir + '\\Deliverable_log.txt'
    sys.stdout = Altium_helpers.Logger(log_filename)
    
    # go to desired working directory
    os.chdir(starting_dir)
    
    # get part number
    try:
        [part_prefix, part_number, part_revision] = Altium_helpers.get_part_number(starting_dir)
        
    except:
        # the output is not up to date with the current structure
        sys.exit()
    # end try
    
    # get part number
    [assy_prefix, assy_number, assy_revision] = Altium_helpers.get_assy_number(starting_dir)    
    
    # define directories
    pdf_dir = starting_dir + '\\' + assy_prefix + '-' + assy_number + assy_revision + 'PD'
    gerber_dir = starting_dir + '\\' + part_prefix + '-' + part_number + part_revision
    output_dir = starting_dir + '\\r' + part_revision + '_' + assy_revision
    output_pdf_dir = output_dir + '\\' + part_number + part_revision + 'PD'
    output_gerber_dir = output_dir + '\\' + part_number + part_revision    
    output_altium_dir = output_dir + '\\Altium Files'
    
    # remove previous deliverables folder for this revision
    while os.path.isdir(output_dir):
        try:
            shutil.rmtree(output_dir)
            
        except Exception as e:
            print(e)
            print('*** Error: Previous output could not be deleted ***')
            if sys.version_info[0] < 3:
                # python 2
                raw_input('Press ENTER to retry')
                
            else:
                input('Press ENTER to retry')
            # end if            
        # end try
        
        time.sleep(0.1)
    # end while
    
    # make the output directories
    os.mkdir(output_dir)
    os.mkdir(output_pdf_dir)
    os.mkdir(output_gerber_dir)
    os.mkdir(output_altium_dir)
    
    # warning tracker
    no_warnings = False
    
    # move master ASSY Config document
    Altium_Excel.copy_assy_config(starting_dir)
    
    # create list to load file modified dates into.
    modified_dates = []
    
    # check the design rule check document
    modified_dates.append(Altium_PDF.check_DRC(pdf_dir))
    
    # check the electrical rule check document
    modified_dates.append(Altium_PDF.check_ERC(pdf_dir))    
    
    # Move all of the Altium files into their folder
    modified_dates.extend(Altium_Files.move_Altium_files(starting_dir, 
                                                         output_altium_dir))
    
    # Move the gerber files and create a readme file for them
    [gerber_dates, layers] = Altium_Files.move_gerbers(gerber_dir, 
                                                       output_gerber_dir, 
                                                       (part_number + part_revision))
    
    # add the gerber modified dates to the list
    modified_dates.extend(gerber_dates)
    
    # move the xps file
    modified_dates.append(Altium_Files.move_xps(starting_dir, 
                                                output_dir, 
                                                (part_number + part_revision)))    
    
    # move all of the other documents
    modified_dates.extend(Altium_Files.move_documents(starting_dir, 
                                                      pdf_dir, 
                                                      output_pdf_dir, 
                                                      gerber_dir,
                                                      (part_number + part_revision),
                                                      layers))
    
    # zip the step file
    modified_dates.append(Altium_Files.zip_step_file(starting_dir, output_dir, part_number))
    
    # find the oldest and newest files used.
    no_warnings = Altium_helpers.check_modified_dates(modified_dates)
    
    # check for warnings
    if not (no_warnings and 
            Altium_Excel.log_warning(get=True) and 
            Altium_PDF.log_warning(get=True) and 
            Altium_Files.log_warning(get=True)):
        print('\n*** Warnings were raised so please reveiw ***')
        
        if sys.version_info[0] < 3:
            # python 2
            raw_input('When the warnings have been reviewed/recitified press ENTER to continue')
            
        else:
            input('When the warnings have been reviewed/recitified press ENTER to continue')
        # end if
    # end if
    
    print("\nUploading of Project information to the Google Drive is disabled")
    #Altium_GS.upload_files(output_dir, Altium_Excel.set_directory.path)    
    
    # construct the final zip file and remove un-needed directories
    zip_filename = Altium_helpers.construct_root_archive(output_dir, (part_number + part_revision))    
    
    # check for errors
    if not (Altium_Excel.log_error(get=True) and 
            Altium_PDF.log_error(get=True) and 
            Altium_Files.log_error(get=True)):
        print('\n*** Errors occurred so please reveiw ***')
    
    print("\nDeliverable generation is complete")
    
    # close the log file
    sys.stdout.close()
    
    # add log file to zip
    shutil.copy(log_filename, output_dir + '\\' + os.path.basename(log_filename))
    zip_file = zipfile.ZipFile(zip_filename, 'a')
    zip_file.write(log_filename, os.path.basename(log_filename))
    zip_file.close()
    
# end if

