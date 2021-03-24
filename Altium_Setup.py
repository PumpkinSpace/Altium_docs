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
@package Altium_Setup.py

Package that sets up the module on a computer and makes sure all dependancies 
are satisfied
"""

__author__ = 'David Wright (david@asteriaec.com)'
__version__ = '0.3.0' #Versioning: http://www.python.org/dev/peps/pep-0386/

import subprocess
import os
import sys
import shutil


#
# ----------------
# Install required dependancies

print("Checking dependancies\n")

print("checking the pip installer framework")
try:
    subprocess.check_call([sys.executable, '-m', 'pip', '--version'])
    print("\tpip is already installed\n")
    
except:
    print("\tinstalling pip")
    inst = subprocess.Popen(['C:\\Python27\python.exe', os.getcwd() + '\\src\\get-pip.py'])
    inst.wait()
    try:
        import pip
        
        print("\tinstall successful\n")
        
    except:
        print("\tinstall unsuccessful! go to https://github.com/BurntSushi/nfldb/wiki/Python-&-pip-Windows-installation")
        input()
        sys.exit()
    # end try
# end try
   
print("checking the .xls reader")
try:
    import xlrd
    print("\txlrd is already installed\n")

except:
    print("\tinstalling xlrd\n")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'xlrd'])
    try:
        import xlrd
        print("\n\tinstall successful\n")
        
    except:
        print("\n\tinstall unsuccessful! go to https://github.com/python-excel/xlrd")
        input()
        sys.exit()
    # end try        
# end try
  
print("checking the .xlsx reader/writer")
try:
    import openpyxl  
    print("\topenpyxl is already installed\n")

except:
    print("\tinstalling openpyxl\n")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl'])
    try:
        import openpyxl
        print("\n\tinstallation successful\n")
        
    except:
        print("\n\tinstall unsuccessful! go to https://openpyxl.readthedocs.io/en/stable/")
        input()
        sys.exit()
    # end try          
# end try

print("checking the pdf reader")
try:
    from PyPDF2 import PdfFileReader
    print("\tpyPdf2 is already installed\n")
    
except:
    print("\tinstalling PyPDF2\n")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'PyPDF2'])
    try:
        from PyPDF2 import PdfFileReader
        print("\n\tinstallation successful\n")
        
    except:
        print("\n\tinstall unsuccessful! go to https://github.com/mstamy2/PyPDF2 and follow the instructions to install")
        input()
        sys.exit()
    # end try     
# end try

print("checking the google sheet modifier")
try:
    import gspread
    print("\tgspread is already installed\n")
    
except:
    print("\tinstalling gspread\n")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'gspread'])
    
    try:
        import gspread
        print("\n\tinstallation successful\n")
        
    except:
        print("\n\tinstall unsuccessful! go to https://github.com/burnash/gspread")
        input()
        sys.exit()
    # end try
# end try  
 
print("checking the google drive client" )
try:
    import pydrive
    print("\tpydrive is already installed\n")
    
except:
    print("\tinstalling pydrive\n")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pydrive'])
    try:
        import pydrive
        print("\n\tinstallation successful\n")
        
    except:
        print("\n\tinstall unsuccessful! go to http://pythonhosted.org/PyDrive/")
        input()
        sys.exit()
    # end try
# end try    
    
print("checking the google authentication tool")
try:
    import oauth2client
    print("\toauth2client is already installed\n")
    
except:
    print("\tinstalling oauth2client\n")
    subprocess.check_call([sys.executable, '-m', 'pip', '--upgrade', 'oauth2client'])
    try:
        import oauth2client
        print("\n\tinstallation successful\n")
        
    except:
        print("\n\tinstall unsuccessful! go to https://github.com/google/oauth2client/")
        input()
        sys.exit()
    # end try
# end try    

print("checking the argument parser")
try:
    import argparse
    print("\targparse is already installed\n")
    
except:
    print("\tinstalling argparse\n")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'argparse'])
    try:
        import argparse
        print("\n\tinstallation successful\n")
        
    except:
        print("\n\tinstall unsuccessful! go to https://pypi.python.org/pypi/argparse")
        input()
        sys.exit()
    # end try
# end try         

print("checking the pdf mining tool")
try:
    import pdfminer
    print("\tpdfminer.six is already installed\n")
    
except:
    print("\tinstalling argparse\n")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pdfminer.six'])
    try:
        import pdfminer
        print("\n\tinstallation successful\n")
        
    except:
        print("\n\tinstall unsuccessful! go to https://github.com/pdfminer/pdfminer.six")
        input()
        sys.exit()
    # end try
# end try  

print('Dependancy check successful')


#
# ----------------
# Create the batch file to run the code

print('\nGenerating Batch file\n')
# lines to write to batch file
batch_lines = ['@echo off',
               'C:\\Python27\python.exe \"' + os.getcwd() + '\\Altium Documentation.py\" \"%CD%\" False\n',
               'Pause']

# path of batch file
batch_file = os.getcwd() + '\\Deliverable.bat'

# remove it if it is already there
if os.path.isfile(batch_file):
    os.remove(batch_file)
# end if

with open(batch_file,'w') as batch:
    batch.writelines(batch_lines)
#end with

shutil.copy(batch_file, os.getcwd() + '\\test folder (02190A)\\Deliverable.bat')

print('Successful')