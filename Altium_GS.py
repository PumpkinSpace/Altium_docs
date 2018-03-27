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

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


#
# -------
# Public Functions






#
# -------
# Private Functions

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    
    # If modifying these scopes, delete your previously saved credentials
    # at ~/.credentials/sheets.googleapis.com-python-quickstart.json
    #SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
    SCOPES = ['https://spreadsheets.google.com/feeds',
              'https://www.googleapis.com/auth/drive']
    CLIENT_SECRET_FILE = 'client_secret.json'
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
    
    # update credentials if needed
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    # end if
    
    return credentials
#end def


def authorise_google_drive():
    # create the authorization object
    gauth = GoogleAuth()
    
    # load the credentials
    gauth.credentials = get_credentials()
    
    # authorize
    gauth.Authorize()
    
    # use that authorization to open the drive
    drive = GoogleDrive(gauth)   
    
    return drive
# end def

def authorise_google_sheet():
    gc = gspread.authorize(get_credentials())
    
    return gc
# end def

def open_bom(drive, gsheet, new_filename):
    
    # get the list of all the files in the BOM folder
    file_list = drive.ListFile({'q': "'1sXLSZtFsRanD2RMn1Q1BLcLsUHGEH7tV' in parents and trashed=false"}).GetList()
    
    # convert this list to a useable dictionary
    file_dict = {i.get('title').encode('ascii', 'ignore'): i.get('id').encode('ascii', 'ignore') for i in file_list}
    
    # check to see if the file we want to edit is already there
    if file_dict.has_key(new_filename):
        # it is so return it opened
        return gsheet.open_by_key(file_dict[new_filename])
    
    else:
        # it is not so copy the master file to create it
        new_sheet = drive.auth.service.files().copy(fileId='1ZCsUHbq6u5djKq659IaI-8rGAfTSbFEjbOaxvYuMSAE',
                                                    body={"parents": [{"kind": "drive#fileLink",
                                                                       "id": folder}], 
                                                          'title': new_filename}).execute()    
        # return the new file opened
        return gsheet.open_by_key(new_sheet['id'])
# end def


def test():
    """
    Test code for this module.
    """
    
    # authorize google drive and google sheet APIs
    drive = authorise_google_drive()
    gsheet = authorise_google_sheet()
    
    # open a 
    sheet = open_bom(drive, gsheet, 'Test_BOM')
# end def


if __name__ == '__main__':
    # if this code is not running as an imported module run test code
    test()
# end if