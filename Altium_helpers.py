import os
import shutil

def get_output_dir(starting_dir):
    
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
    andrews_dir = starting_dir + '\\Andrews Format'
    
    if not os.path.exists(andrews_dir):
        os.makedirs(andrews_dir)
    # end if
    
    return andrews_dir
#end def

def get_pdf_dir(starting_dir):
    andrews_dir = get_Andrews_dir(starting_dir)
    
    pdf_dir = andrews_dir+'\\'+'pdfs'
    
    if not os.path.exists(pdf_dir):
        os.makedirs(pdf_dir)
    # end if
    
    return pdf_dir
# end def

def get_gerbers_dir(starting_dir):
    
    andrews_dir = get_Andrews_dir(starting_dir)
    gerbers_dir = andrews_dir + '\\Gerbers'
    
    if not os.path.exists(gerbers_dir):
        os.makedirs(gerbers_dir)
    # end if
    
    return gerbers_dir
#end def
    
        
def clear_output(starting_dir):
    andrews_dir = starting_dir + '\\Andrews Format'
    
    if os.path.exists(andrews_dir):
        try:
            shutil.rmtree(andrews_dir)
            
        except:
            print '***  Error: Could not remove previous Andrews Format Folder  ***\n\n'
            return False
        # end try
    # end if
    
    if os.path.isfile(starting_dir + '\\test.xlsx'):
        try:
            os.remove(starting_dir + '\\test.xlsx')
            
        except:
            print '***  Error: Could not remove previous test.xlsx file  ***\n\n'
            return False 
        # end try  
    
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
    
    return True
# end def

