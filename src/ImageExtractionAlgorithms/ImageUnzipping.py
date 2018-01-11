import os, tarfile
from tkinter import Tk, filedialog
from os import listdir
import bz2, zipfile

imageLoc = 'C:/Users/eTreatMD/Documents/ImageDatabase'

def getSD(path):
    return filter(os.path.isdir, [os.path.join(path, f) for f in os.listdir(path)])

imageOutputLoc = 'C:/Users/eTreatMD/Documents/Images'
if not os.path.exists(imageOutputLoc):
    os.makedirs(imageOutputLoc)

def main():

    subdir = getSD(imageLoc)
    for folderA in subdir:
        subsubdir = getSD(folderA)
        for folderB in subsubdir:
            subsubsubdir = getSD(folderB)
            for folderC in subsubsubdir:
                subsubsubsubdir = getSD(folderC)
                for folderD in subsubsubsubdir:
                    subsubsubsubsubdir = getSD(folderD)
                    for folderE in subsubsubsubsubdir:
                        files = os.listdir(folderE)
                        for file in files:
                            path = folderE + '\\' + file
                            print(path)
                            dc = bz2.decompress(open(path, 'rb').read())
                            print(dc)
                            
main()
