import glob
import os
#deletes files if file is 1KB
list_of_files = glob.glob('C:\\Users\\DDiaz\\Documents\\UNcomtrade\\*.csv')

for file in list_of_files:
    if round(os.stat(file).st_size/1024) == 1:

        #deletes file
        #os.remove(file)
