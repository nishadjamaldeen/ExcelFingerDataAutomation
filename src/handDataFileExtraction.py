import os, time
from os import listdir


def main():
    initialPath = 'C:/Users/eTreatMD/Downloads/RRH_Image_Data';
    destinationPath = 'C:/Users/eTreatMD/Documents/hand/Inputs';

    if not os.path.exists(destinationPath):
        os.makedirs(destinationPath)

    lookupFiles(initialPath, destinationPath)
    print("Finished")

def lookupFiles(initialPath, destinationPath):

    directory = listdir(initialPath)
    for files in directory:
        path = os.path.join(initialPath, files)
        newFilePath = os.path.join(destinationPath, files)
        print(path)

        if files.endswith('.jpg') and not files.endswith('Output.jpg'):
            os.rename(path, newFilePath)
        elif os.path.isdir(path):
            lookupFiles(path, destinationPath)

main()


            
