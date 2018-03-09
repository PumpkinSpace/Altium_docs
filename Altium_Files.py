import Altium_Excel
import Altium_helpers
import os
import shutil
import pyPdf
import re
import Altium_OCR

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

def log_error(get = False):
    if get:
        return log_error.no_errors
    
    else:
        log_error.no_errors = False
    # end if
# end def
log_error.no_errors = True


def log_warning(get = False):
    if get:
        return log_warning.no_warnings
    
    else:
        log_warning.no_warnings = False
    # end if
# end def
log_warning.no_warnings = True


def get_part_number(starting_dir):
    part_list = starting_dir.split('(')
    return part_list[1][:-1]    
# end def

def extract_assy_info(pdf_text, starting_dir):
    
    assy_blocks = pdf_text.split('case')
    
    if len(assy_blocks) != 3:
        print "*** Warning, too many ASSY_REV blocks. ***"
        log_warning()
    # end if
    
    list_1 = assy_blocks[2].split(';')[:-1]
    
    while ((assy_blocks[1][0].isalpha() == False) and (assy_blocks[1] != '')):
        assy_blocks[1] = assy_blocks[1][1:]
    # end while
    
    list_0 = assy_blocks[1].split(';')[:-1]
    
    if len(list_0) != len(list_1):
        print "*** Warning, missmatched ASSY_REV blocks ***"
        log_warning()
    # end if        
    
    if not Altium_Excel.set_assy_options(starting_dir, list_0, list_1):
        log_error()
    # end if
# end def


########## Function to find the page number in a schematic document ############
# Altium often outputs the PDF with the apges in the wrong order so this
# function finds the page number on each page and returns a string of the number
def get_page_number(path, pn, starting_dir):
    
    # extract the text from a pdf page
    pdf_text = Altium_OCR.convert_pdf_to_txt(path)
    
    if 'ASSY' in pdf_text:
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

def move_Altium_files(starting_dir):
    print 'Moving Altium Files...'
    
    modified_dates = []
    
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
                
                modified_dates.append(os.path.getmtime(starting_dir+'\\'+filename))
            # end if
        # end for
    # end for
    
    try:
        shutil.make_archive(andrews_dir+'\\Altium Files', 'zip', altium_dir)
        shutil.rmtree(altium_dir)     
        
    except:
        print '*** Error: could not create Altium files.zip ***'
        log_error()
    # end try
    
    print 'Complete! \n'    
    
    return modified_dates
# end def

def create_readme(starting_dir, layers):
    print '\tGenerating Readme File...'
    
    # list to store new lines in
    new_readme_lines = []
    
    # list of extensions already used to prevent repeats
    extensions_used = []
    
    outputs_dir = Altium_helpers.get_output_dir(starting_dir)
    
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
    new_readme_lines.append('README'+ str(layers) + '.TXT             This file                       		ASCII\n')
    
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

def check_gerber_folder(gerber_dir):
    
    file_list = os.listdir(gerber_dir)
    
    for ext in required_gerber_list:
        ext_found = False
        for filename in file_list:
            if filename.endswith(ext):
                ext_found = True
            # end if
        # end for
        
        if not ext_found:
            print '*** Error: no ' + ext + ' file output to gerbers ***'
            log_error()
        # end if
    # end for
    
    readme_found = False
    pick_found = False
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
                
def move_gerbers(starting_dir):
    print 'Moving Gerber Files...'
    
    modified_dates = []
    
    outputs_dir = Altium_helpers.get_output_dir(starting_dir)
    
    # make gerber directory
    gerbers_dir = Altium_helpers.get_gerbers_dir(starting_dir)
    
    andrews_dir = Altium_helpers.get_Andrews_dir(starting_dir)
    
    # get list of gerber files
    gerber_file_list = os.listdir(outputs_dir)
    
    if len(gerber_file_list) < 10:
        # No gerbers exist
        print '***   Error: No Gerbers have been generated   ***\n\n'
        log_error()
        return None, None
    # end 
    
    # count the number of layers by comparing filename to valid extensions
    # also move the gerber files into the 
    layers = 0
    
    for filename in gerber_file_list:
        good_filename = True
        for ext in layer_gerber_list:
    
            if filename.endswith(ext):
                # is a valid gerber layer
                layers+=1
    
            if filename.endswith('G7') or filename.endswith('GP7'):
                # Too many layers
                print '***  Error: Script needs to be extended to handle this many layers   ***\n\n'
                log_error()
            # end
        # end
        
        for ext in bad_gerber_ext:
            if (filename.endswith(ext) and ("Pick Place" not in filename)):
                good_filename = False
            # end if
        # end for
        
        if good_filename:
            if filename.endswith('.xls'):
                if ('DNP' not in filename):
                    # is the BOM file
                    
                    part_number = get_part_number(starting_dir)
                    shutil.copyfile(outputs_dir + '\\' + filename, \
                                            gerbers_dir + '\\' + part_number + '_BOM.xls')
                    modified_dates.append(os.path.getmtime(outputs_dir+'\\'+filename))    
                # end if
            
            else:
                try:
                    shutil.copyfile(outputs_dir+'\\'+filename, gerbers_dir+'\\'+filename)
                    
                except:
                    print '*** Error: could not move ' + filename + ' ***'
                    log_error()
                # end try
                
                modified_dates.append(os.path.getmtime(outputs_dir+'\\'+filename))     
            # end if
        # end if
    # end for
    
    create_readme(starting_dir, layers)
    
    check_gerber_folder(gerbers_dir)
    
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

def move_bom(starting_dir):
    print 'Moving BOM...'
    no_bom = True
    
    modified_dates = []
    
    part_number = get_part_number(starting_dir)
    
    outputs_dir = Altium_helpers.get_output_dir(starting_dir)
    
    # get list of gerber files
    gerber_file_list = os.listdir(outputs_dir)    
    
    # create the PDF Directory
    pdf_dir = Altium_helpers.get_pdf_dir(starting_dir)    
    
    # search for BOM in project outputs folder
    for filename in gerber_file_list:
        if filename.endswith('xls'):
            # BOM found
            # shutil.copyfile(outputs_dir + '\\' + filename, \
            #                pdf_dir + '\\' + part_number + '_BOM.xls')
            # modified_dates.append(os.path.getmtime(outputs_dir+'\\'+filename))
            
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

def manage_schematic(starting_dir):
    print 'Finding Schematic Document...'
    
    modified_date = None
    
    pdf_dir = Altium_helpers.get_pdf_dir(starting_dir)
    
    part_number = get_part_number(starting_dir)
    
    root_file_list = os.listdir(starting_dir)
    
    # search for a schematic document in the root directory
    for filename in root_file_list:
        if (filename.lower().endswith('pdf') and 
            (len(filename) > 12) and 
            (filename != 'PCB Prints.pdf') and 
            ('layers' not in filename)):
            # schematic found
            no_schematic = False
            modified_date = os.path.getmtime(starting_dir+'\\'+filename)
            break
        # end
    # end
    
    if no_schematic:
        # No schematic was found
        print('***   Error: No Schematic Document was found   ***')
        log_error()
        return None, None
    # end
    
    print '\tReading the Schematic file...'
    
    # open pdf file
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
        return None, None        
    # end try
    
    print '\tComplete!'
    
    print '\tRenaming the PDFs...'
    
    mod_doc_found = False
    
    # rename the pdfs with the correct filenames
    for i in range(1,page+2):
        old_file_name = pdf_dir + '\\' + part_number + '--' + str(i) + '.pdf'
        
        # find page number
        [page_number, is_assy] = get_page_number(old_file_name, part_number, starting_dir)
        
        if is_assy:
            mod_doc_found = True
        # end if
            
        # rename the file
        try:
            os.rename(old_file_name, pdf_dir + '\\' + \
                      part_number + '-' + page_number + '.pdf')
            
        except:
            print('***   Error: Could rename pdf document   ***')
            log_error()
        # end try
    # end for
    
    if not mod_doc_found:
        print('***   Warning: No Modification information found in schematic   ***')
        log_warning()
    # end if
    
    print '\tComplete!'    
    
    return modified_date
# end def

def move_xps(starting_dir):
    
    root_file_list = os.listdir(starting_dir)
    
    andrews_dir = Altium_helpers.get_Andrews_dir(starting_dir)
    
    xps_file = ''
    modified_date = None
    
    for filename in root_file_list:
        if filename.endswith('xps'):
            xps_file = filename
            modified_date = os.path.getmtime(starting_dir+'\\'+filename)
        # end if
    # end for
    
    if xps_file == '':
        print '*** Error: no .xps file found ***'
        log_error()
        return None
    # end if
    
    xps_ext = xps_file.split('.')[1]
    
    xps_dir = andrews_dir + '//xps_temp'
    
    if not os.path.exists(xps_dir):
        os.makedirs(xps_dir)
    # end if
    
    part_number = get_part_number(starting_dir)
    
    # copy file into temp directory
    shutil.copy(starting_dir+'\\'+xps_file, xps_dir +'\\'+part_number + '.' + xps_ext)

    # make archive
    shutil.make_archive(andrews_dir+'\\'+part_number+'_xps', 'zip', xps_dir)
    shutil.rmtree(xps_dir, ignore_errors=True)     
    
    return modified_date
# end def
    

def move_documents(starting_dir, exe_OCR, layers):
    
    # find the Schematic and BOM documents
    no_schematic = True
    
    andrews_dir = Altium_helpers.get_Andrews_dir(starting_dir)
    pdf_dir = Altium_helpers.get_pdf_dir(starting_dir)
    
    modified_dates = move_bom(starting_dir)
    
    modified_dates.append(manage_schematic(starting_dir))
    
    modified_dates.extend(Altium_Excel.construct_assembly_doc(starting_dir))
    
    modified_dates.append(move_xps(starting_dir))
    
    root_file_list = os.listdir(starting_dir)
    
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
    
    ################## Create the PDF files for each Gerber layer ##################
    
    modified_dates.extend(Altium_OCR.perform_Altium_OCR(exe_OCR, starting_dir, layers))
    
    if not (Altium_OCR.log_error(get=True) and Altium_Excel.log_error(get=True)):
        log_error()
    # end if
    
    if not (Altium_OCR.log_warning(get=True) and Altium_Excel.log_warning(get=True)):
        log_warning()
    # end if   
    
    shutil.make_archive(andrews_dir+'\\'+get_part_number(starting_dir)+ \
                        'PD', 'zip', pdf_dir)
    shutil.rmtree(pdf_dir)     
    
    return modified_dates
# end def


def zip_step_file(starting_dir):
    
    print 'Zipping Step File...'
    
    step_file_found = False
    modified_date = None
    
    andrews_dir = Altium_helpers.get_Andrews_dir(starting_dir)
    
    root_file_list = os.listdir(starting_dir)
    
    # search for step file
    for filename in root_file_list:
        if filename.endswith('.step'):
            # step file has been found
            
            modified_date = os.path.getmtime(starting_dir+'\\'+filename)
            
            # make a folder to put the step file in temporarily
            step_dir = starting_dir + '\\step_temp'
            if os.path.exists(step_dir):
                shutil.rmtree(step_dir)
            # end
            
            # create temporary directory in which to place all of the files needed
            os.makedirs(step_dir)    
            
            part_number = get_part_number(starting_dir)
            
            # copy file into temp directory
            shutil.copy(starting_dir+'\\'+filename, step_dir +'\\'+part_number+'.step')
            
            # make archive
            shutil.make_archive(andrews_dir+'\\'+part_number+'_step', 'zip', step_dir)
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