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
import re
import sys
import pyPdf
import pdfminer
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO
import subprocess
import time
# This program also requires the following installed packages:
# pypdfocr 
# imagemagik
# Pillow
# reportlab
# watchdog
# pypdf2

#import reportlab
#import watchdog
#import PyPDF2



#################### Change this for each implementation #######################
# directory where the Circuit board files are stored
starting_dir = 'C:\Users\pumpkinadmin\Dropbox\Satellite\Pumpkin\Power Supply Board (01666A)'

##################### Function to extract the text from a PDF ##################
# From: stackoverflow.com/questions/40031622/pdfminer-error-for-one-type-of-
#       pdfs-too-many-vluae-to-unpack
# Courtesy of Chianti5
def convert_pdf_to_txt(path):
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

    fp.close()
    device.close()
    retstr.close()
    return text
# end

########## Function to find the page number in a schematic document ############
# Altium often outputs the PDF with the apges in the wrong order so this
# function finds the page number on each page and returns a string of the number
def get_page_number(path):
    # extract the text from a pdf page
    pdf_text = convert_pdf_to_txt(path)
    
    # find the location of the of string that separates the page numers
    of_index = pdf_text.rfind(' of ')
    if of_index == -1:
        # if of is not found return an error
        return -1
    elif pdf_text[of_index+5].isdigit() and (pdf_text[of_index+6] != '5'):
        # there are two digits in the page number, return both
        return pdf_text[of_index-2:of_index-1]
    
    else:
        # return the single digit page number
        return pdf_text[of_index-1]
    # end
# end


#################### Function to 'Beautify' input text #########################
# The OCDed text is noisy and full of errors.
# This function performs sustitutions of common miss-readings to enhance
# performance. New errors should be added here as the are encountered
def beautify(text):
    text = text.lower()
    text = text.replace('u', 'w')
    text = text.replace('1', 'l')
    text = text.replace('i', 'l')
    text = text.replace('n', 'r')
    text = text.replace('0', 'o')
    text = text.replace('j', 'l')
    text = text.replace('g', 'y')
    text = text.replace('h', 'a')
    text = text.replace('u', 'w')
    text = text.replace('f', 'r')
    text = text.replace('v', 'r')
    text = text.replace('¦', 'a')
    return text
# end
    

############ Function to detemine the correct name of a PDF layer file #########
# extract the text from a OCRed pdf and see if certain substrings are present
# within that text to determine the correct file name
def get_filename(path):
    pdf_text = beautify(''.join(convert_pdf_to_txt(path).split()))
    if (beautify('number') in pdf_text)\
       and (beautify('Drill') in pdf_text):
        # This is a Mechanical Drawing file
        get_filename.MECHDWG = True
        return 'MECHDWG.pdf'
    
    elif (beautify('Bottom') in pdf_text)\
         and (beautify('silk') in pdf_text)\
         and (beautify('Solder') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Bottom Assembly Drawing file
        get_filename.ADB = True
        return 'ADB0230.pdf'
    
    elif (beautify('Top') in pdf_text)\
         and (beautify('silk') in pdf_text)\
         and (beautify('Solder') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Top Assembly Drawing file
        get_filename.ADT = True
        return 'ADT0127.pdf'
    
    elif (beautify('Top') in pdf_text)\
         and (beautify('silk') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Top Silkscreen File
        get_filename.SST = True
        return 'SST0126.pdf' 
    
    elif (beautify('Top') in pdf_text)\
         and (beautify('mask') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Top Soldermask file
        get_filename.SMT = True
        return 'SMT0125.pdf'     
    
    elif (beautify('Bottom') in pdf_text)\
         and (beautify('silk') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Bottom Silkscreen file
        get_filename.SSB = True
        return 'SSB0229.pdf'  
    
    elif (beautify('Bottom') in pdf_text)\
         and (beautify('mask') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Bottom Soldermask file
        get_filename.SMB = True
        return 'SMB0223.pdf'  
    
    elif (beautify('Drill') in pdf_text)\
         and (beautify('Drawing') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Drill Drawing File
        get_filename.DD = True
        return 'DD0124.pdf' 
    
    elif (beautify('Bottom') in pdf_text)\
         and (beautify('Solder') in pdf_text)\
         and (beautify('Paste') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Bottom Solder Paste file
        get_filename.SPB = True
        return 'SPB0223.pdf'  
    
    elif (beautify('Top') in pdf_text)\
         and (beautify('Solder') in pdf_text)\
         and (beautify('Paste') in pdf_text)\
         and (beautify('COMPO') not in pdf_text):
        # This is a Top Solder Paste file
        get_filename.SPT = True
        return 'SPT0123.pdf'   
    
    elif ((beautify('Layer') in pdf_text) or (beautify('Plane') in pdf_text)) \
         and (beautify('COMPO') not in pdf_text):
        # This is a Layer Artwork file
        name = 'ART' + format(get_filename.layer, '02') + '.pdf'
        get_filename.layer += 1
        return name
    
    else:
        return -1
    # end
# end

part_list = starting_dir.split('(')
part_number = part_list[1][:-1]

# go to desired working directory
os.chdir(starting_dir)

# if the folder already exists delete it and then make again from scratch
andrews_dir = starting_dir + '\\Andrews Format'
if os.path.exists(andrews_dir):
    shutil.rmtree(andrews_dir)
# end

# create temporary directory in which to place all of the files needed
os.makedirs(andrews_dir)

############################### Altium Files ###################################

# create directory for the altium files
altium_dir = andrews_dir + '\Altium Files'
os.makedirs(altium_dir)

# desired file extensions
altium_ext = ['Outjob', 'cam', 'PrjPcb', 'PrjPcbStructure', \
              'PcbDoc', 'SchDoc', 'Harness']

# get file list of root directory
root_file_list = os.listdir(starting_dir)

# copy desired files
for filename in root_file_list:
    # compare each filename with the desired set of extensions
    for ext in altium_ext:
        if filename.endswith(ext):
            # This file is desired to copy it across
            shutil.copyfile(starting_dir+'\\'+filename, altium_dir+'\\'+filename)
        # end
    # end
    
    # also remove files left over from previous usages of this program
    if filename.endswith('.zip') or filename.endswith('_ocr.pdf'):
        os.remove(starting_dir + '\\' + filename)
    # end
# end

##################### Move desired Gerber Files to directory ###################
        
# find project Outputs folder
for filename in root_file_list:
    if filename.startswith('Project Outputs'):
        os.chdir(starting_dir+'\\'+filename)
        break
    # end
# end

outputs_dir = os.getcwd()

if outputs_dir == starting_dir:
    # The required folder was not found -> EXIT
    print '***   No Project Outputs Directory Found   ***\n\n'
    sys.exit()
# end

# make gerber directory
gerbers_dir = andrews_dir+'\\'+'Gerbers'

os.makedirs(gerbers_dir)

# get list of gerber files
gerber_file_list = os.listdir(outputs_dir)

if len(gerber_file_list) < 10:
    # No gerbers exist
    print '***   No Gerbers have been generated   ***\n\n'
    sys.exit()
# end 

# rejected file extensions
bad_gerber_ext = ['zip', 'ods', 'xls', 'Report.Txt', '2.txt', \
                  '4.txt', '6.txt', '8.txt', 'drc', 'html']

# file extensions that represent layers
layer_gerber_list = ['.GTL', '.GBL', '.G1', '.G2', '.G3', '.G4', '.G5', \
                     '.G6', '.GP1', '.GP2', '.GP3', '.GP4', '.GP5', '.GP6']

# count the number of layers by comparing filename to valid extensions
layers = 0
for filename in gerber_file_list:
    for ext in layer_gerber_list:
        
        if filename.endswith(ext):
            # is a valid gerber layer
            layers+=1
            
        if filename.endswith('G7') or filename.endswith('GP7'):
            # Too many layers
            print '***   Script needs to be extended to handle this many layers   ***\n\n'
            sys.exit()
        # end
    # end
# end

############################# Create Readme file ###############################

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
                     'GPX': '.GPX			Artwork file for internal plane Layer X	Gerber\n',
                     'LDP': '.LDP			Altium .LDP file\n',
                     'REP': '.REP			Altium Report file				ASCII\n',
                     'RUL': '.RUL			Altium .RUL file\n',		
                     'TXT': '.TXT			Altium Drill file				ASCII\n',
                     'csv': '.CSV			Pick and Place File				ASCII\n',
                     'APR_LIB': '.APR_LIB	        Altium Aperture Library 		        Gerber\n'}

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

# list to store new lines in
new_readme_lines = []

# list of extensions already used to prevent repeats
extensions_used = []

# Iterate through every file
for filename in gerber_file_list:
    
    # reject unwanted file extensions
    good_filename = True
    for ext in bad_gerber_ext:
        if filename.endswith(ext):
            # is not wanted
            good_filename = False
        # end
    # end
            
    if good_filename:
        # add wanted extensions to list
        shutil.copyfile(outputs_dir+'\\'+filename, gerbers_dir+'\\'+filename)
        
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
            # end
            
            extensions_used.append(extension)
        # end
    # end
# end
            
# order gerber file lines
new_readme_lines.sort()

# add line for the readme file
new_readme_lines.append('\nREADME'+ str(layers) + '.TXT             This file                       		ASCII\n')

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
readme_file = open(gerbers_dir+'\\'+'README'+str(layers)+'.TXT', 'w')
readme_file.writelines(Readme_lines)
readme_file.close()

# move all gerbers into the desired zip archive
shutil.make_archive(andrews_dir+'\\'+part_number, 'zip', gerbers_dir)
attempt_count = 0
while attempt_count < 10:
    try:
        shutil.rmtree(gerbers_dir) 
        break
    except os.error:
        attempt_count += 1
        time.sleep(100)
    # end
# end
if attempt_count == 10:
    print '*** Random error, restart the program ***'
    sys.exit()
# end

###################### Manage the Schematic and BOM files ######################

# create the PDF Directory
pdf_dir = andrews_dir+'\\'+'pdfs'
os.makedirs(pdf_dir)

# find the Schematic and BOM documents
no_schematic = True
no_bom = True

# search for BOM in project outputs folder
for filename in gerber_file_list:
    if filename.endswith('xls'):
        # BOM found
        shutil.copyfile(outputs_dir + '\\' + filename, \
                        pdf_dir + '\\' + part_number + '_BOM.xls')
        no_bom = False
        break
    # end
# end   

if no_bom:
    # No BOM was found
    print('***   No BOM was found in project outputs   ***')
    sys.exit()
# end

# search for a schematic document in the root directory
for filename in root_file_list:
    if filename.endswith('pdf') and (len(filename) > 12) and (filename != 'PCB Prints.pdf') and ('layers' not in filename):
        # schematic found
        no_schematic = False
        break
    # end
# end

if no_schematic:
    # No schematic was found
    print('***   No Schematic Document was found   ***')
    sys.exit()
# end


# open pdf file
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
        # end
    # end
# end

# rename the pdfs with the correct filenames
for i in range(1,page+2):
    old_file_name = pdf_dir + '\\' + part_number + '--' + str(i) + '.pdf'
    
    # find page number
    page_number = get_page_number(old_file_name)
    if page_number == -1:
        # Could not find a page number
        print('***   No page number found in document '+ \
              old_file_name + '  ***')
        sys.exit()      
    # end
    
    # rename the file
    os.rename(old_file_name, pdf_dir + '\\' + \
              part_number + '-' + page_number + '.pdf')
# end

################## Create the PDF files for each Gerber layer ##################

# Find the layers.pdf file
if ('Layers.pdf' not in root_file_list) and ('layers.pdf' not in root_file_list) and ('PCB Prints.pdf' not in root_file_list):
    # Could not find layers.pdf
    print('***   No layers.pdf or PCB Prints.pdf file found  ***')
    sys.exit()      
# end

if ('Layers.pdf' in root_file_list):
    os.rename(starting_dir+'\\Layers.pdf', starting_dir+'\\layers.pdf')
# end

if ('PCB Prints.pdf' in root_file_list):
    os.rename(starting_dir+'\\PCB Prints.pdf', starting_dir+'\\layers.pdf')
# end

# copy the layers pdf into the ocr directory to allow ocr to be performed
ocr_dir = 'C:\\Python27\\Lib\\site-packages\\pypdfocr'
shutil.copy(starting_dir+'\\layers.pdf', ocr_dir +'\\layers.pdf')

# perform OCR on the layers pdf and wait for it to complete
cmd = subprocess.Popen(['python', 'pypdfocr.py', 'layers.pdf'], cwd=ocr_dir)
cmd.wait()

# line breaks following OCR output
print '\n\n'

# return OCR file from OCR directory and clean the OCR directory
os.remove(ocr_dir +'\\layers.pdf')
shutil.move(ocr_dir +'\\layers_ocr.pdf', starting_dir +'\\layers_ocr.pdf')

# read the ocred file
layers_pdf = pyPdf.PdfFileReader(open(starting_dir+'\\layers_ocr.pdf', "rb"))

# define static variables for the get_filename function
# to facilitate the generation of warnings
get_filename.layer = 1
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
    # end
    
    # find the desired filename for the new file
    new_filename = get_filename(file_name)
    if new_filename != -1:
        # This file is desired so rename with the correct name
        if os.path.isfile(pdf_dir + '\\' + new_filename):
            print '*** Error, ' + new_filename + ' already exists ***'
        else:
            os.rename(file_name, pdf_dir + '\\' + new_filename)
        # end
        
    else:
        # file is not wanted so remove it
        os.remove(file_name)
    # end
# end

# Generate warnings for pecuiliar outputs
no_warnings = True
if (get_filename.layer % 2 == 0) or (get_filename.layer != (layers + 1)):
    print '*** WARNING wrong number of layers printed ***'
    no_warnings = False
# end
if (get_filename.MECHDWG == False):
    print '*** WARNING No MECHDWG file output ***'
    no_warnings = False
# end
if (get_filename.ADB == False):
    print '*** WARNING No ADB file output ***'
    no_warnings = False
# end
if (get_filename.ADT == False):
    print '*** WARNING No ADT file output ***'
    no_warnings = False
# end
if (get_filename.SST == False):
    print '*** WARNING No SST file output ***'
    no_warnings = False
# end
if (get_filename.SMT == False):
    print '*** WARNING No SMT file output ***'
    no_warnings = False
# end
if (get_filename.SSB == False):
    print '*** WARNING No SSB file output ***'
    no_warnings = False
# end
if (get_filename.SMB == False):
    print '*** WARNING No SMB file output ***'
    no_warnings = False
# end
if (get_filename.DD == False):
    print '*** WARNING No DD file output ***'
    no_warnings = False
# end
if (get_filename.SPB == False): 
    print '*** WARNING No SPB file output ***'
    no_warnings = False
# end
if (get_filename.SPT == False):
    print '*** WARNING No SPT file output ***'
    no_warnings = False
# end

# create archives and remove un-needed directorys
shutil.make_archive(andrews_dir+'\\'+part_number+ 'PD', 'zip', pdf_dir)
shutil.rmtree(pdf_dir) 

shutil.make_archive(starting_dir+'\\'+part_number+'_Folder', 'zip', andrews_dir)
shutil.rmtree(andrews_dir, ignore_errors=True)

# indicate completion
print '\n*** Directory has been generated successfully ***'
if not no_warnings:
    print '\n*** Warnings were raised so please reveiw ***'
# end
    
