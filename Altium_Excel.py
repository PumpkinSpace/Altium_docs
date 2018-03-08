import xlrd
import os
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl import Workbook

thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))

# define BOM layaout constants
bom_d_col = 0
bom_pn_col = 5
bom_dnp_col = 1
bom_header_rows = 6


def get_output_dir(starting_dir):
    
    root_file_list = os.listdir(starting_dir)
    # find project Outputs folder
    for filename in root_file_list:
        if filename.startswith('Project Outputs'):
            return starting_dir + '\\' + filename
        # end
    # end
    
    # if this code was reached, then no folder was found
    print '***   No Project Outputs Directory Found   ***\n\n'
    sys.exit()    
# end def 


def extract_items(cell):
    cell_string = repr(cell).strip('text:u\'')
    cell_list = cell_string.split(', ')
    return cell_list
# end def


def get_bom_lists(starting_dir, d_list, pn_list, DNP = False):
    
    # find output directory
    output_dir = get_output_dir(starting_dir)
    
    filename = ''
    
    # find two BOM docs.
    for name in os.listdir(output_dir):
        if DNP == True and name.startswith('DNP'):
            filename = name
            
        elif DNP == False and name.endswith('.xls'):
            filename = name
        # end if
    # end for    
    
    # open the BOM sheet
    doc = xlrd.open_workbook(output_dir + '\\' + filename).sheet_by_index(0)
    
    for row in range(bom_header_rows, doc.nrows):
        # find the part number in the BOM Doc.
        pn_list.append(extract_items(doc.cell(row,bom_pn_col)))
        d_list.append(extract_items(doc.cell(row,bom_d_col)))
    # end for    
    
    return doc
# end def


def fill_assy_bom(starting_dir, dnp_d_list, comp_dnp_list, dnp_doc):
    
    # find assy_rev document
    assy_filename = ''
    
    for filename in os.listdir(starting_dir):
        if ('ASSY' in filename) and ('REV' in filename):
            assy_filename = filename
        # end if
    # end for
        
    # open the assy_rev document
    assy_doc = openpyxl.load_workbook(starting_dir + '\\' + assy_filename)
    
    # open the BOM sheet
    bom_sheet = assy_doc.get_sheet_by_name('BOM')
    
    # full replace all cells in the BOM sheet
    for i in range(0,dnp_doc.nrows):
        for j in range(0,dnp_doc.ncols):
            if (i >= bom_header_rows):
                # into BOM information
                if (j == 0) and (i < dnp_doc.nrows):
                    # add the designators to place
                    bom_sheet.cell(i+1,j+1).value = ', '.join(dnp_d_list[i-6])
                    
                elif ((j == 1) and (i < dnp_doc.nrows-1)):
                    # add the designators not to place
                    bom_sheet.cell(i+1,j+1).value = ', '.join(comp_dnp_list[i-6])   
                    
                else:
                    # copy the part info from the BOM
                    bom_sheet.cell(i+1,j+1).value = dnp_doc.cell_value(i,j)
                # end if
                
                # add borders to all cells
                bom_sheet.cell(i+1,j+1).border = thin_border
            
            else:
                # copy the BOM info from the BOM
                bom_sheet.cell(i+1,j+1).value = dnp_doc.cell_value(i,j)            
            # end if
        # end for
    # end for
    
    # save the file
    assy_doc.save(starting_dir + '\\test.xlsx')
    
    # replace the old file
    os.remove(starting_dir + '\\' + assy_filename)
    os.rename(starting_dir + '\\test.xlsx', starting_dir + '\\' + assy_filename)
# end def
    
def construct_assembly_doc(starting_dir):
    
    # find output directory
    output_dir = get_output_dir(starting_dir)

    # initialise BOM column arrays
    bom_d_list = []
    bom_pn_list = []
    dnp_d_list = []
    dnp_pn_list = []
    
    # fill the BOM column arrays
    bom_doc = get_bom_lists(starting_dir, bom_d_list, bom_pn_list)
    dnp_doc = get_bom_lists(starting_dir, dnp_d_list, dnp_pn_list, DNP=True)

    # initialise the DNP component list
    comp_dnp_list = []
    
    # iternate through the part numbers in the total list
    for index in range(0,len(dnp_pn_list)):
        if dnp_pn_list[index] not in bom_pn_list:
            # if the part number is not present in the list of only placed
            # then move all such parts to the DNP list
            comp_dnp_list.append(dnp_d_list[index])
            
            # empty the placed list
            dnp_d_list[index] = []
            
        else:
            # the part number is present in the placed list so check every designator
            designator_list = []
            
            # find the designator list in the other BOM
            bom_index = bom_pn_list.index(dnp_pn_list[index])
            
            # check each designator
            for designator in dnp_d_list[index]:
                # compare to the other list
                if designator not in bom_d_list[bom_index]:
                    # it is not found so add the componant to the dnp list
                    designator_list.append(designator)
                    
                    # remove the designator from this list
                    dnp_d_list[index].remove(designator)
                # end if
            # end for
            
            # add to the dnp list
            comp_dnp_list.append(designator_list)
            
        # end if 
    # end for
    
    fill_assy_bom(starting_dir, dnp_d_list, comp_dnp_list, dnp_doc)
#end def

def test():
    """
    Test code for this module.
    """
    construct_assembly_doc(os.getcwd() + '\\test folder')
#end def

if __name__ == '__main__':
    # if this code is not running as an imported module run test code
    test()
# end if



                


                
        
        
