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

print "Checking dependancies\n"

print "checking the pip installer framework"
try:
    import pip
    pip.main(['install', 'pyopenssl', 'ndg-httpsclient', 'pyasn1'])
    print "\tpip is already installed\n"
    
except:
    print "\tinstalling pip"
    inst = subprocess.Popen(['C:\\Python27\python.exe', os.getcwd() + '\\src\\get-pip.py'])
    inst.wait()
    try:
        import pip
        
        print "\tinstall successful\n"
        
    except:
        print "\tinstall unsuccessful! go to https://github.com/BurntSushi/nfldb/wiki/Python-&-pip-Windows-installation"
        raw_input()
        sys.exit()
    # end try
# end try
   
print "checking the .xls reader"
try:
    import xlrd
    print "\txlrd is already installed\n"

except:
    print "\tinstalling xlrd\n"
    pip.main(['install', 'xlrd'])
    try:
        import xlrd
        print "\n\tinstall successful\n"
        
    except:
        print "\n\tinstall unsuccessful! go to https://github.com/python-excel/xlrd"
        raw_input()
        sys.exit()
    # end try        
# end try
  
print "checking the .xlsx reader/writer"
try:
    import openpyxl  
    print "\topenpyxl is already installed\n"

except:
    print "\tinstalling openpyxl\n"
    pip.main(['install', 'openpyxl'])
    try:
        import openpyxl
        print "\n\tinstallation successful\n"
        
    except:
        print "\n\tinstall unsuccessful! go to https://openpyxl.readthedocs.io/en/stable/"
        raw_input()
        sys.exit()
    # end try          
# end try

print "checking the pdf reader"
try:
    import pyPdf
    print "\tpyPdf is already installed\n"
    
except:
    print "\tto install pypdf go to http://pybrary.net/pyPdf/pyPdf-1.13.win32.exe"
    raw_input()
    sys.exit()
# end try

print "checking the google sheet modifier"
try:
    import gspread
    print "\tgspread is already installed\n"
    
except:
    print "\tinstalling gspread\n"
    pip.main(['install', 'gspread'])
    try:
        import gspread
        print "\n\tinstallation successful\n"
        
    except:
        print "\n\tinstall unsuccessful! go to https://github.com/burnash/gspread"
        raw_input()
        sys.exit()
    # end try
# end try  
 
print "checking the google drive client" 
try:
    import pydrive
    print "\tpydrive is already installed\n"
    
except:
    print "\tinstalling pydrive\n"
    pip.main(['install', 'pydrive'])
    try:
        import pydrive
        print "\n\tinstallation successful\n"
        
    except:
        print "\n\tinstall unsuccessful! go to http://pythonhosted.org/PyDrive/"
        raw_input()
        sys.exit()
    # end try
# end try    
    
print "checking the google authentication tool"
try:
    import oauth2client
    print "\toauth2client is already installed\n"
    
except:
    print "\tinstalling oauth2client\n"
    pip.main(['install', '--upgrade', 'oauth2client'])
    try:
        import oauth2client
        print "\n\tinstallation successful\n"
        
    except:
        print "\n\tinstall unsuccessful! go to https://github.com/google/oauth2client/"
        raw_input()
        sys.exit()
    # end try
# end try    

print "checking the argument parser"
try:
    import argparse
    print "\targparse is already installed\n"
    
except:
    print "\tinstalling argparse\n"
    pip.main(['install', 'argparse'])
    try:
        import argparse
        print "\n\tinstallation successful\n"
        
    except:
        print "\n\tinstall unsuccessful! go to https://pypi.python.org/pypi/argparse"
        raw_input()
        sys.exit()
    # end try
# end try         

print "checking the pdf mining tool"
try:
    import pdfminer
    print "\tpdfminer is already installed\n"
    
except:
    print '\tpdfminer is not installed'
    print '\tgo to https://github.com/euske/pdfminer and follow the instructions to install'
    raw_input()
    sys.exit()
# end try  

print "checking for Tesseract"

try:
    cmd = subprocess.check_output(['tesseract', '-v'])
    print "\tTesseract is already installed\n"
    
except:
    print '\tTesseract is not installed or is not on the system PATH'
    print '\tgo to https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-setup-3.05.01.exe'
    raw_input()
# end try

print "checking for ImageMagick"

try:
    cmd = subprocess.check_output(['magick'])
    print "\tImageMagick is already installed\n"
    
except:
    print '\tImageMagick is not installed or is not on the system PATH'
    print '\tgo to https://www.imagemagick.org/download/binaries/ImageMagick-7.0.8-0-Q16-x64-dll.exe'
    raw_input()
# end try

print "checking for GhostScript"

try:
    cmd = subprocess.check_output(['gswin32c', '-h'])
    print "\tGhostScript is already installed\n"
    
except:
    print '\tGhostScript is not installed or is not on the system PATH'
    print '\tgo to https://www.ghostscript.com/download/gsdnld.html'
    raw_input()
# end try

print 'Dependancy check successful'

#
# ----------------
# Test the OCR executable

print '\nTesting OCR\n'
sys.stdout.flush()
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
    print "\n\tOCR test was successful"
    
except:
    print '\n\tOCR test was unsuccessful'
    print '\tplease copy the output from this script and email to david@asteriaec.com'
    print '\tplease also ensure that there are no .pdf files in the OCR directory'
    raw_input()
    sys.exit()
# end try

#
# ----------------
# Create the batch file to run the code

print '\nGenerating Batch file\n'
# lines to write to batch file
batch_lines = ['@echo off',
               'C:\\Python27\python.exe \"' + os.getcwd() + '\\Altium Documentation.py\" \"%CD%\" False\n',
               'timeout 10']

# path of batch file
batch_file = os.getcwd() + '\\Deliverable.bat'

# remove it if it is already there
if os.path.isfile(batch_file):
    os.remove(batch_file)
# end if

with open(batch_file,'w') as batch:
    batch.writelines(batch_lines)
#end with

shutil.copy(batch_file, os.getcwd() + '\\test folder (01234A)\\Deliverable.bat')

print 'Successful'