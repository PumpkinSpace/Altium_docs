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
__version__ = '0.1.0' #Versioning: http://www.python.org/dev/peps/pep-0386/

import subprocess
import os
import sys
import shutil


#
# ----------------
# Install required dependancies

print 'Checking dependancies\n'

# check the pip installer framework
try:
    import pip
    print "pip is already installed"
    
except:
    print "installing pip"
    inst = subprocess.Popen(['C:\\Python27\python.exe', os.getcwd() + '\\src\\get-pip.py'])
    inst.wait()
    try:
        import pip
        print "install successful"
        
    except:
        print "install unsuccessful! go to https://github.com/BurntSushi/nfldb/wiki/Python-&-pip-Windows-installation"
        sys.exit()
    # end try
# end try
   
# check the .xls reader
try:
    import xlrd
    print "xlrd is already installed"

except:
    print "installing xlrd"
    pip.main(['install', 'xlrd'])
    try:
        import xlrd
        print "install successful"
        
    except:
        print "install unsuccessful! go to https://github.com/python-excel/xlrd"
        sys.exit()
    # end try        
# end try
  
# check the .xlsx reader/writer 
try:
    import openpyxl  
    print "openpyxl is already installed"

except:
    print "installing openpyxl"
    pip.main(['install', 'openpyxl'])
    try:
        import openpyxl
        print "installation successful"
        
    except:
        print "install unsuccessful! go to https://openpyxl.readthedocs.io/en/stable/"
        sys.exit()
    # end try          
# end try

# check the pdf reader
try:
    import pyPdf
    print "pyPdf is already installed"
    
except:
    print "to install pypdf go to http://pybrary.net/pyPdf/pyPdf-1.13.win32.exe"
    sys.exit()
# end try

# check the google sheet modifier   
try:
    import gspread
    print "gspread is already installed"
    
except:
    print "installing gspread"
    pip.main(['install', 'gspread'])
    try:
        import gspread
        print "installation successful"
        
    except:
        print "install unsuccessful! go to https://github.com/burnash/gspread"
        sys.exit()
    # end try
# end try  
 
# check the google drive client 
try:
    import pydrive
    print "pydrive is already installed"
    
except:
    print "installing pydrive"
    pip.main(['install', 'pydrive'])
    try:
        import pydrive
        print "installation successful"
        
    except:
        print "install unsuccessful! go to http://pythonhosted.org/PyDrive/"
        sys.exit()
    # end try
# end try    
    
# check the google authentication tool
try:
    import oauth2client
    print "oauth2client is already installed"
    
except:
    print "installing oauth2client"
    pip.main(['install', '--upgrade', 'oauth2client'])
    try:
        import oauth2client
        print "installation successful"
        
    except:
        print "install unsuccessful! go to https://github.com/google/oauth2client/"
        sys.exit()
    # end try
# end try    

# check the argument parser
try:
    import argparse
    print "argparse is already installed"
    
except:
    print "installing argparse"
    pip.main(['install', 'argparse'])
    try:
        import argparse
        print "installation successful"
        
    except:
        print "install unsuccessful! go to https://pypi.python.org/pypi/argparse"
        sys.exit()
    # end try
# end try         

# check the pdf mining tool
try:
    import pdfminer
    print "pdfminer is already installed"
    
except:
    print 'pdfminer is not installed'
    print 'go to https://github.com/euske/pdfminer and follow the instructions to install'
    sys.exit()
# end try  

print 'Dependancy check successful'

#
# ----------------
# Test the OCR executable

print '\nTesting OCR'
ocr_dir = os.getcwd() + '\\OCR'
filename = ocr_dir + '\\Instructions.pdf'

# copy instructions file to OCR directory
shutil.copy(os.getcwd() + '\\Instructions.pdf', filename)

# perform OCR on it
cmd = subprocess.Popen(['pypdfocr.exe', filename], cwd=ocr_dir, shell=True)
cmd.wait()

# cleanup
try:
    os.remove(filename)
    os.remove(ocr_dir + '\\Instructions_ocr.pdf')
    print "OCR test was successful"
    
except:
    print 'OCR test was unsuccessful'
    print 'please copy the output from this script and email to david@asteriaec.com'
    print 'please also ensure that there are no .pdf files in the OCR directory'
    sys.exit()
# end try

#
# ----------------
# Create the batch file to run the code

print '\nGenerating Batch file'
# lines to write to batch file
batch_lines = ['C:\\Python27\python.exe \"' + os.getcwd() + '\\Altium Documentation.py\" \"%CD%\" True\n',
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

shutil.copy(batch_file, os.getcwd() + '\\test folder\\Deliverable.bat')

print 'Successful'