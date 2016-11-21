Altium Documentaion.py Readme
**********************************

1. Required dependancies:
Pypdf:
pdfminer: https://github.com/euske/pdfminer/
imagemagik: https://www.imagemagick.org/download/binaries/ImageMagick-7.0.3-7-Q16-x64-dll.exe
pypdf_ocr: https://github.com/virantha/pypdfocr/blob/master/dist/pypdfocr.exe?raw=true
Pillow: in installer above
reportlab: in installer above
watchdog: in installer above
pypdf2: https://github.com/mstamy2/PyPDF2
tesseract: https://github.com/tesseract-ocr/tesseract
Ghostscript: https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs920/gs920w32.exe


2. Requirements of directory:
	root directory input into the script must contain:
		Project Outputs Folder containing:
			Full Gerber outputs
			A BOM in .xls format
		A (multi-sheet if required) PDF of the schematics with a name atleast 12 characters long,
			This pdf must have the text for page number in the bottom right corner formatted as: x of y
			Use Altium's Smart PDF to generate this file, do not include a BOM or PCB prints in this file.
		A pdf of all layers in the board, use default prints to generate the following prints, printing a 1.0 scaled print in monochrome to pdf. 
		This PDF should either be called layers.pdf or PCB_Prints.pdf
			Layer set:
				Layer for all in layers
				Keepout
			Mech Dwg:
				Mechanical 2 with all manufacturing specs
				Keepout layer
				Drill drawing with drill table
			Top assembly drawing:
				Top Soldermask
				Top Silkscreen
				keepout
			Bottom Assembly drawing:
				Bottom Soldermask
				Bottom Silkscreen
				Keepout
		All altium files associated with the design. The .pcb file needs to have all the layers laelled as per the default document provided.
		

3. Using the script:
IDE version: 	Open Altium Documentation.py in you IDE
				Change starting_dir to the directory that contains your altium files (called root above)
				Run the script
				If successful the script will place partnumber_Folder.zip in the directory provided
				If unsuccessful see below or talk to David

Other Version:	Open Altium Documentation.py in a text editor
				Change starting_dir to the directory that contains your altium files (called root above) 
				In command line navigate to the directory of Altium Documentation.py
				run 'python Altium Documentation.py'
				If successful the script will place partnumber_Folder.zip in the directory provided
				If unsuccessful see below or talk to David
		

3. Common Errors and solutions:

Cannot delete directory, it isn't empty: re-run the script

Cannot delete something, it is being used by another program: Close the other program

Cannot create filename, file already exists: This is the most common error - 
	This occurs when the code tries to create a pdf file that already exists, 
	because it either erroniously created a file earlier or is trying to create this file in error now.
	the easiest solution to this is to go into the newly created Andrews Format/pdfs folder and click on the remaining layers-# files and see what they should be called.
	Then manually rename that file and delete conflicting ones as needed.
	When this has been done, create a partnumberPD.zip (eg. 01254BPD) archive of all the the files in this directory and place it in the Andrews Format directory
	Then in Andrews Format create a zip of the partnumber.zip, partnumberPD.zip and Altium files folder, call this partnumber_Folder.zip (eg. 01254B_Folder) and place it in the root directory
	Then delete the Andrews Format Directory
	
	