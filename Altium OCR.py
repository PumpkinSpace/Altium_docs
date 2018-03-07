import os
import sys
import shutil
import subprocess
import pyPdf


def get_OCR_dir(exe_OCR):
    if exe_OCR:
        return os.getcwd()
        
    else:
        return 'C:\\Python27\\Lib\\site-packages\\pypdfocr'
    # end if
# end def

def adjust_layer_filename(starting_dir):
    
    root_file_list = os.listdir(starting_dir)
    
    # Find the layers.pdf file
    if (('Layers.pdf' not in root_file_list) 
        and ('layers.pdf' not in root_file_list) 
        and ('PCB Prints.pdf' not in root_file_list)):
        # Could not find layers.pdf
        print('***   No layers.pdf or PCB Prints.pdf file found  ***')
        sys.exit()      
    # end if
    
    if ('Layers.pdf' in root_file_list):
        mod_date = os.path.getmtime(starting_dir+'\\Layers.pdf')
        os.rename(starting_dir+'\\Layers.pdf', starting_dir+'\\layers.pdf')
    
    elif ('PCB Prints.pdf' in root_file_list):
        mod_date = os.path.getmtime(starting_dir+'\\PCB Prints.pdf')
        os.rename(starting_dir+'\\PCB Prints.pdf', starting_dir+'\\layers.pdf')
    
    else:
        mod_date = os.path.getmtime(starting_dir+'\\layers.pdf')
    # end if  
    
    return mod_date
# end def

def get_filename_init():
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
# end def
    
############ Function to detemine the correct name of a PDF layer file #########
# extract the text from a OCRed pdf and see if certain substrings are present
# within that text to determine the correct file name
def get_filename(path):
    pdf_text = beautify(convert_pdf_to_txt(path))
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
        #name = 'ART' + format(get_filename.layer, '02') + '.pdf'
        get_filename.layer += 1
        #return name
        return get_layer_number(pdf_text)
    
    else:
        return -1
    # end if
# end def

#################### Function to 'Beautify' input text #########################
# The OCDed text is noisy and full of errors.
# This function performs sustitutions of common miss-readings to enhance
# performance. New errors should be added here as the are encountered
def beautify(text):
    text = ''.join(text.split())
    text = text.lower()
    text = ''.join([c for c in text if c.isalnum()])
    #text = text.replace('u', 'w')
    text = text.replace('1', 'l')
    text = text.replace('i', 'l')
    #text = text.replace('n', 'r')
    #text = text.replace('0', 'o')
    #text = text.replace('j', 'l')
    #text = text.replace('g', 'y')
    #text = text.replace('h', 'a')
    #text = text.replace('u', 'w')
    #text = text.replace('f', 'r')
    #text = text.replace('v', 'r')
    return text
# end def


def get_layer_number(page_text):
    if (beautify('Layer 1') in page_text):
        return 'ART01.pdf'
    
    elif (beautify('Layer 2') in page_text):
        return 'ART02.pdf'
    
    elif (beautify('Layer 3') in page_text):
        return 'ART03.pdf'
    
    elif (beautify('Layer 4') in page_text):
        return 'ART04.pdf'

    elif (beautify('Layer 5') in page_text):
        return 'ART05.pdf'
    
    elif (beautify('Layer 6') in page_text):
        return 'ART06.pdf'

    elif (beautify('Layer 7') in page_text):
        return 'ART07.pdf'
    
    elif (beautify('Layer 8') in page_text):
        return 'ART08.pdf'
    
    else:
        print "this layer did not get named"
        print beautify('Layer 1')
        return "unnamed layer"
    #end if
#end def

def split_OCR_pages(starting_dir, ocr_dir):
    # return OCR file from OCR directory and clean the OCR directory
    os.remove(ocr_dir +'\\layers.pdf')
    shutil.move(ocr_dir +'\\layers_ocr.pdf', starting_dir +'\\layers_ocr.pdf')
    
    # read the OCR'ed file
    layers_pdf = pyPdf.PdfFileReader(open(starting_dir+'\\layers_ocr.pdf', "rb"))
    
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
        # end with
        
        # find the desired filename for the new file
        new_filename = get_filename(file_name)
        if new_filename != -1:
            # This file is desired so rename with the correct name
            if os.path.isfile(pdf_dir + '\\' + new_filename):
                print '*** Error, ' + new_filename + ' already exists ***'
                
            else:
                os.rename(file_name, pdf_dir + '\\' + new_filename)
            # end if
            
        else:
            # file is not wanted so remove it
            os.remove(file_name)
        # end if
    # end for
# end def

def check_OCR_outputs(no_warnings):
    # Generate warnings for pecuiliar outputs
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
    
    
def perform_Altium_OCR(no_errors, no_warnings, exe_OCR, starting_dir):
    
    print 'Starting OCR on Layers file...'
    
    modified_dates = [adjust_layer_filename(starting_dir)]
    
    # start the OCR on the layers file
    orc_dir = get_OCR_dir(exe_OCR)
    if exe_OCR:
        # copy the layers pdf into the ocr directory to allow ocr to be performed
        shutil.copy(starting_dir+'\\layers.pdf', ocr_dir +'\\layers.pdf')
        
        # perform OCR on the layers pdf
        cmd = subprocess.Popen(['pypdfocr.exe', 'layers.pdf'], cwd=ocr_dir, shell=True)
        
    else:
        
        # copy the layers pdf into the ocr directory to allow ocr to be performed
        shutil.copy(starting_dir+'\\layers.pdf', ocr_dir +'\\layers.pdf')
        
        # perform OCR on the layers pdf
        cmd = subprocess.Popen(['python', 'pypdfocr.py', 'layers.pdf'], cwd=ocr_dir)
    # end if
    
    # wait for analysis to complete
    cmd.wait()
    
    print 'Complete! \n'
    
    print 'Renaming the layer PDFs...'
    
    split_OCR_pages(starting_dir, ocr_dir)
    
    check_OCR_outputs(no_warnings)
    
    print 'Complete! \n'
    
    return modified_dates
#end def