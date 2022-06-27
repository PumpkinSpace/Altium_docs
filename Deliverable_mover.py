# Deliverable.bat 

import os
import shutil

#directory to search:
search_dir = "Y:\Shared drives\Asteria - Engineering\Pumpkin\Pumpkin Circuit Boards"

bat_paths = []

for root, dirs, files in os.walk(search_dir):
    for filename in files:
        if filename.endswith("Deliverable.bat"):
            bat_paths.append(os.path.join(root,filename))
        #end if
    #end for
# end for

for bat_file in bat_paths:
    os.remove(bat_file)
    shutil.copy("C:\Pumpkin\Altium_docs\Deliverable.bat", bat_file)
#end for

print("Complete")
