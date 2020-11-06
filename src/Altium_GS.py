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
@package Altium_GS.py

Package that updates google sheets with BOM information
"""

__author__ = 'David Wright (david@asteriaec.com)'
__version__ = '0.2.0' #Versioning: http://www.python.org/dev/peps/pep-0386/


#
# -------
# Imports

import gspread
import os
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import httplib2
import time
import argparse
import datetime


max_assy_rev = 24

BOM_TEMPLATE_KEY = '1BZkbVMP741pOVWVwJmrh_eax_7kRtmDlvPcFPEm8y90'

#
# -------
# Classes
    
class assembly_info:
    """
    Object to house all of the bom information to be loaded into the google sheet

    @attribute   is_blocking          access control boolean, True if the object
                                      is being used by another task (bool).
    @attribute   list_0               list of '0' value assembly optinos (list).
    @attribute   list_1               list of '1' value assembly options (list).
    @attribute   designators          list of lists of designators (list).
    @attribute   dnp_designators      list of lists of dnp designators (list).
    @attribute   descriptions         list of component descriptions (list).
    @attribute   quantities           list of component quantities (list).
    @attribute   manufacturers        list of component manufacturers (list).
    @attribute   manufacturer_pns     list of manufacturer part numbers (list).
    @attribute   suppliers            list of component suppliers (list).
    @attribute   supplier_pns         list of supplier part numbers (list).
    @attribute   sub_manufacturer_pns list of substitute manufacturer part 
                                      numbers (list).
    @attribute   sub_supplier_pns     list of substitute supplier part numbers
                                      (list).
    @attribute   subtotals            list of subtotal prices (list).
    """    
    
    def __init__(self):
        """
        Initialise all of the attributes.
        """        
        self.is_blocking = False
        self.list_0 = []
        self.list_1 = []
        self.designators = []
        self.dnp_designators = []
        self.descriptions = []
        self.quantities = []
        self.manufacturers = []
        self.manufacturer_pns = []
        self.suppliers = []
        self.supplier_pns = []
        self.sub_manufacturer_pns = []
        self.sub_supplier_pns = []
        self.sub_suppliers = []
        self.sub_manufacturer = []
        self.subtotals = []
    # end def
    
    def is_free(self):
        """
        Determine if the object is available for access.
        
        @return    (bool)     True if the object is free for access
        """         
        return not self.is_blocking
    # end def
    
    def __enter__(self):
        """
        Enter the object and take ownership.
        """        
        if self.is_blocking:
            raise Exception('Race condition on access, access has already been claimed')
        # end if
        self.is_blocking = True
        
        return self
    # end def
    
    def __exit__(self, type, value, traceback):
        """
        Exit the object and relinquish ownership.
        """         
        self.is_blocking = False
    # end def
# end class


#
# -------
# Public Functions

def upload_files(output_dir, prog_dir):
    """
    Uploads the generated .zip file to the google dirve folder.

    @param:    output_dir     The full path of output package (string).
    @param:    prog_dir       The full path that the program is running from
                              (string).
    """    
    print('\nUploading to google drive...')
    file_list = os.listdir(output_dir)
    
    for filename in file_list:
        if filename.endswith('Folder.zip'):
            break
        # end if
    # end for
    
    # authorize google drive API
    drive = authorise_google_drive(prog_dir + '\\src')    
    
    # get the list of all the files in deliverables folder
    file_list = drive.ListFile({'q': "'1vDTz6N-1QbUlkbb7QrFj082YUpkIRZFL' in parents and trashed=false"}).GetList()
    
    # convert this list to a useable dictionary
    file_dict = {i.get('title').encode('ascii', 'ignore'): i for i in file_list}
    
    # check to see if the file we want to edit is already there, if so delete it
    if file_dict.has_key(filename):
        # it is so return it opened
        file_dict[filename].Delete()
    # end if
    
    # create the file and upload it
    zip_file = drive.CreateFile({'title': filename, 
                                 'parents': [{'kind':'drive#fileLink', 
                                              'id': '1vDTz6N-1QbUlkbb7QrFj082YUpkIRZFL'}]})
    zip_file.SetContentFile(starting_dir + '\\' + filename)
    zip_file.Upload()
    
    print('Complete!\n')
# end def    
    

def populate_online_bom(prog_dir, part_number, assy_number, revision, assy_info):
    """
    Populates a BOM in the Pumpkin google drive the the appropriate information 
    for this project.

    @param:    prog_dir       The full path that the program is running from
                              (string).
    @param:    part_number    The part number of the BOM to write to (string).
    @param:    assy_number    The assembly number of the BOM to write to (string).
    @param:    assy_info      The information to populate the BOM with 
                              (assembly_info).
    """    
    
    print('Updating the google sheet BOM...')
    # determine the scr directory path
    src_dir = prog_dir + '\\src'
    
    # attempt to authorize the google credentials
    try:
        # authorize google drive and google sheet APIs
        drive = authorise_google_drive(src_dir)
        gsheet = authorise_google_sheet(src_dir)
    
    except:
        print('*** Error: Failed to Authorize google credentials, no BOM uploaded ***\n')
        return None
    # end try
    
    # create the name of the BOM from the part number and open it.
    bom_name = part_number + '/' + assy_number + revision
    online_bom = open_bom(drive, gsheet, bom_name)
    
    # open the options worksheet
    options = online_bom.worksheet("Options")
    
    # find the headers for the columns
    header_row = options.find("0 value").row
    col_0 = options.find("0 value").col
    col_1 = options.find("1 value").col
    
    #read cell array
    cells = options.range(header_row+1, min(col_0, col_1), 
                          options.row_count, max(col_0, col_1))
    
    # edit cell array
    for cell in cells:
        if (cell.row <= len(assy_info.list_0)+header_row):
            # cell has an assy_rev associated with it
            if cell.col == col_0:
                cell.value = assy_info.list_0[cell.row-header_row-1]
                
            elif cell.col == col_1:
                cell.value = assy_info.list_1[cell.row-header_row-1]
            # end if   
        # end if
    # end for
    
    options.update_cells(cells)
    
    # open the bom sheet
    bom = online_bom.worksheet("PCBA Components")
    
    # find the header row
    header_row = bom.find("Item").row
    col_headers = bom.row_values(header_row)
    
    cells = bom.range(header_row+1, 1, max(header_row + len(assy_info.designators), bom.row_count), len(col_headers))
    
    for cell in cells:
        i = cell.col
        j = cell.row-header_row
        
        if j-1 >= len(assy_info.designators):
            cell.value = ''
        
        elif col_headers[i-1] == "Item":
            cell.value = j
            
        elif col_headers[i-1] == 'Qty':
            cell.value = assy_info.quantities[j-1]
            
        elif col_headers[i-1] == 'Reference Designator':
            cell.value = ', '.join(assy_info.designators[j-1])
            
        elif col_headers[i-1] == 'Description':
            cell.value = assy_info.descriptions[j-1]

        elif col_headers[i-1] == 'DNP':
            cell.value = ', '.join(assy_info.dnp_designators[j-1])
            
        elif col_headers[i-1] == 'Manufacturer':
            cell.value = assy_info.manufacturers[j-1]
                
        elif col_headers[i-1] == 'MPN':
            cell.value = assy_info.manufacturer_pns[j-1]

        elif col_headers[i-1] == 'Supplier':
            cell.value = assy_info.suppliers[j-1]
            
        elif col_headers[i-1] == 'Sub Supplier':
            cell.value = assy_info.sub_suppliers[j-1]            
            
        elif col_headers[i-1] == 'SPN':
            cell.value = assy_info.supplier_pns[j-1]

        elif col_headers[i-1] == 'SubSPN':
            cell.value = assy_info.sub_supplier_pns[j-1]
            
        elif col_headers[i-1] == 'SubMPN':
            cell.value = assy_info.sub_manufacturer_pns[j-1]     
            
        elif col_headers[i-1] == 'Sub Manufacturer':
            cell.value = assy_info.sub_manufacturer[j-1]            
            
        elif col_headers[i-1] == 'Ext. Cost (USD)':
            cell.value = assy_info.subtotals[j-1]
        # end if
    # end for
    
    bom.update_cells(cells)
    
    # open the ECO sheet
    bom = online_bom.worksheet("ECOs")    
    
    # get the cells to load data into
    cells = bom.range(1, 1, 1, 6)
    
    # load the relevant data into the correct place
    for cell in cells:
        i = cell.col
        
        if (i == 2):
            cell.value = '705-' + part_number
            
        elif (i == 4):
            cell.value = assy_number
            
        elif (i ==6):
            cell.value = revision
        # end if
    # end for
    
    # upload the data
    bom.update_cells(cells)
                            
    print('Complete!\n')
    return True
#end def
    
    
#
# -------
# Private Functions

def write_cell(sheet, row, col, value):
    """
    Updates a single cell in the google sheet and catches the write quota 
    exception and will re-attempt until it is successful.

    @param:    sheet          The gsheet object to write to (gspread.gsheet).
    @param:    row            The row index within the sheet (int).
    @param:    col            The column index within the sheet (int).
    @param:    value          What to write to the cell (string).
    """    
    no_write = True
    while no_write:
        # the write has not yet been successful
        try:
            # attempt to write to the cell
            sheet.update_cell(row, col, value)
            no_write = False
            
        except gspread.v4.exceptions.APIError as e:
            # if the error was not due to write group expiry the throw the error
            if 'WriteGroup' not in str(e):
                raise e
            # end if
            
            # wait for a second to see if the token has renewed.
            time.sleep(1)
        #end try
    #end while
#end def


def get_credentials(src_dir):
    """
    Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.
    
    @param:    src_dir        The path of the folder containing the credentials 
                              file (string).
    @return:   (credentials)  User account credentials for google drive.
    """
    
    # find the credentials file.
    for filename in os.listdir(src_dir):
        if filename.endswith('.json'):
            CLIENT_SECRET_FILE = src_dir + '\\' + filename
        # end if
    # end for
    
    # If modifying these scopes, delete your previously saved credentials
    # at ~/.credentials/sheets.googleapis.com-python-quickstart.json
    #SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
    SCOPES = ['https://spreadsheets.google.com/feeds',
              'https://www.googleapis.com/auth/drive']
    APPLICATION_NAME = 'Altium_GS'
    
    # find the User directory and create a credentials directory within it
    # if one does not already exist
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    # end if
    
    # create the file to store the credentials in
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')
    store = Storage(credential_path)
    credentials = store.get()
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
    
    # update credentials if needed
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    # end if
    
    # refresh credentials if they are expired
    if credentials.access_token_expired:
        credentials.refresh(httplib2.Http())
    # end if    
    
    return credentials
#end def


def authorise_google_drive(src_dir):
    """
    Authorizes access to google drive.

    @param:  src_dir                The full path of the folder containing the 
                                    credentials file (string).
    @return: (pydrive.GoogleDrive)  Object for accessing the google drive.
    """    
    # create the authorization object
    gauth = GoogleAuth()
    
    # load the credentials
    gauth.credentials = get_credentials(src_dir)
    
    # authorize
    gauth.Authorize()
    
    # use that authorization to open the drive
    drive = GoogleDrive(gauth)   
    
    return drive
# end def


def authorise_google_sheet(src_dir):
    """
    Authorizes access to google sheets.

    @param:    src_dir             The full path of the folder containing the 
                                   credentials file (string).
    @return:   (gspread.gspread)   Object for accessing google sheets.
    """    
    
    #authenticate the gspread object
    gc = gspread.authorize(get_credentials(src_dir))
    
    return gc
# end def


def open_bom(drive, gsheet, new_filename):
    """
    Opens a Bill of materials from the google drive.

    @param:    drive             Authenticated google drive object 
                                 (pydrive.GoogleDrive).
    @param:    gsheet            Authenticated gspread object (gspread.gspread).
    @param:    new_filename      The filename of the new BOM (string).
    @return:   (gspread.gsheet)  google sheet object for modification.
    """ 
    
    # get the list of all the files in the BOM folder
    file_list = drive.ListFile({'q': "'1sXLSZtFsRanD2RMn1Q1BLcLsUHGEH7tV' in parents and trashed=false"}).GetList()
    
    # convert this list to a useable dictionary
    file_dict = {i.get('title').encode('utf-8', 'ignore').decode(): i.get('id').encode('utf-8', 'ignore').decode() for i in file_list}
    
    # get the modified date of the BOM template
    temp_file = drive.CreateFile({'id': BOM_TEMPLATE_KEY})
    template_mod_date = temp_file['modifiedDate']   
    
    create_new_bom = True
    
    # check to see if the file we want to edit is already there
    if new_filename in file_dict:
        
        # get the modified date of the BOM and the last modifier
        temp_file = drive.CreateFile({'id': file_dict[new_filename]})
        bom_mod_date = temp_file['modifiedDate']
        bom_modifier = temp_file['lastModifyingUser']['displayName']
        
        # if the BOM was modified more recently than the template
        # or I was not the modifier then update the BOM rather than deleting it
        if ((template_mod_date < bom_mod_date) or (bom_modifier != 'David Wright')):
            # it is so return it opened
            print('\t Updating the ' + new_filename + ' google sheet')
            create_new_bom = False
            return gsheet.open_by_key(file_dict[new_filename])  
        
        else:
            # the file is out of date so delete it
            print('\t Deleting the out of date ' + new_filename + ' google sheet')
            temp_file.Trash()
        # end if
    # end if
    
    if create_new_bom:
        # it is not so copy the master file to create it
        new_sheet = drive.auth.service.files().copy(fileId=BOM_TEMPLATE_KEY,
                                                    body={"parents": [{"kind": "drive#fileLink",
                                                                       "id": '1sXLSZtFsRanD2RMn1Q1BLcLsUHGEH7tV'}], 
                                                          'title': new_filename}).execute()    
        print('\t Creating the ' + new_filename + ' google sheet')
        
        # return the new file opened
        return gsheet.open_by_key(new_sheet['id'])
    # end if
# end def


def test():
    """
    Test code for this module.
    """
    
    src_dir = os.getcwd()
    prog_dir = '\\'.join(src_dir.split('\\')[:-1])
    
    # authorize google drive and google sheet APIs
    drive = authorise_google_drive(src_dir)
    gsheet = authorise_google_sheet(src_dir)
    
    # open a 
    sheet = open_bom(drive, gsheet, 'Test_BOM')
    
    assy_info = assembly_info()
    
    if assy_info.is_free:
        with assy_info:
            assy_info.list_0 = ['item 1', 'item 2']
            assy_info.list_1 = ['item 3', 'item 4']
            populate_online_bom(prog_dir, 'Test_BOM', assy_info)
        #end with
    # end if
    
    
    
    
# end def


if __name__ == '__main__':
    # if this code is not running as an imported module run test code
    test()
# end if
