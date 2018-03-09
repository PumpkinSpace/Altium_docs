# PCB document backup creater
# Pumpkin Inc.
# David Wright 2016
#
# In order to work the starting directory must contain:
#     A project Outputs Folder contiaining:
#           Full Gerber outputs
#           A BOM in .xls format
#     A pdf of the schematics with a name at least 12 characters long
#     A pdf of all the layers in the design (Default Prints) called layers.pdf
#     Circuit Board design of up to 12 layers (maybe less depending on plane/signal layer reakdown


# Required imports
import os
import shutil
import datetime
import Altium_OCR
import Altium_Excel
import Altium_helpers
import Altium_Files


#################### Change this for each implementation #######################
# directory where the Circuit board files are stored
starting_dir = 'C:\Users\Asteria\Dropbox\Satellite\Pumpkin PCBs\ADCS Interface Module 2 (01845B)'

exe_OCR = False

# go to desired working directory
os.chdir(starting_dir)

if not Altium_helpers.clear_output(starting_dir):
    print '*** Error: Previous output could not be deleted ***'
# end if

no_warnings = False

# create list to load file modified dates into
modified_dates = Altium_Files.move_Altium_files(starting_dir)

# Move the gerber files and create a readme file for them
[gerber_dates, layers] = Altium_Files.move_gerbers(starting_dir)

modified_dates.extend(gerber_dates)

modified_dates.extend(Altium_Files.move_documents(starting_dir, exe_OCR, layers))

modified_dates.append(Altium_Files.zip_step_file(starting_dir))

# find the oldest and newest files used.

min_time = modified_dates[0]
max_time = modified_dates[0]

for time in modified_dates:
    if time != None:
        if time < min_time:
            min_time = time
            
        elif time > max_time:
            max_time = time
        # end if
    # end if
# end for

# detect old files
if ((max_time - min_time) > 600):
    # there is more than 10 mins between the oldest and youngest file dates
    early_date = datetime.datetime.fromtimestamp(min_time).strftime('%Y-%m-%d %H:%M:%S')
    print '*** WARNING possibly delivering old files, date ' + early_date + ' ***'
    no_warnings = False
# end if

print 'Constructing Archive...'

# create archives and remove un-needed directorys

part_number = Altium_Files.get_part_number(starting_dir)

andrews_dir = Altium_helpers.get_Andrews_dir(starting_dir)

shutil.make_archive(starting_dir+'\\'+part_number+'_Folder', 'zip', andrews_dir)
shutil.rmtree(andrews_dir, ignore_errors=True)

# indicate completion
print '\n*** Directory ' + part_number + '_Folder.zip' + ' has been generated successfully ***'

if not (no_warnings and Altium_Excel.log_warning(get=True) and 
        Altium_OCR.log_warning(get=True) and Altium_Files.log_warning(get=True)):
    print '\n*** Warnings were raised so please reveiw ***'
# end if

if not (Altium_Excel.log_error(get=True) and 
        Altium_OCR.log_error(get=True) and Altium_Files.log_error(get=True)):
    print '\n*** Errors occurred so please reveiw ***'
# end if

    
