import os

pathBase = "C:\\Users\\DDiaz\\Documents\\data.txt"




if(not os.path.exists(pathBase)):
    file = open(pathBase, "w+")
    file.close()
