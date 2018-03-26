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
from pydrive.auth import ServiceAccountCredentials


#
# -------
# Public Functions

def get_GS_credentials(prog_dir):
    """
    Function to retreive credentials for accessing google sheets.

    @param[in]   prog_dir:            The folder in which this code was executed 
                                      (full path) (string).
    @return      (credentials)        Modification dates of the schematic.
    """    
    
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

    return ServiceAccountCredentials.from_json_keyfile_name('Pumpkin BOM manipulation-e03a7790b73b.json', scope)
# end def

gc = gspread.authorize(get_GS_credentials(os.getcwd()))

gauth = GoogleAuth()
gauth.credentials = (get_GS_credentials(os.getcwd()))
gauth.Authorize()
drive = GoogleDrive(gauth)

folder = '1sXLSZtFsRanD2RMn1Q1BLcLsUHGEH7tV'
filekey = '101OUsGfmhATnClCmyXC54kGcfUMoizxiulZEWxcqS0A'
filename = 'Test_BOM'

textfile = drive.CreateFile()
textfile.SetContentFile('test.txt')
textfile.Upload()

drive.CreateFile({'id':textfile['id']}).GetContentFile('test.txt')

drive.auth.service.files().copy(fileId=filekey,
                                body={"parents": [{"kind": "drive#fileLink",
                                              "id": folder}], 'title': filename}).execute()

#spread = gc.open_by_key('101OUsGfmhATnClCmyXC54kGcfUMoizxiulZEWxcqS0A')




