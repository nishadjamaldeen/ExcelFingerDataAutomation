import os, tarfile, bz2, zipfile, shutil
from os import listdir
from PIL import Image

#inputPath = filedialog.askopenfilename(initialdir='C:/Users/eTreatMD/Downloads')



def main():
    
    inputPath = 'C:/Users/eTreatMD/Downloads/colorferet.tar'
    tarOutputPath = 'C:/Users/eTreatMD/Documents/ImageDatabase'
    imageCollectionPath = 'C:/Users/eTreatMD/Documents/ImageDatabaseCollection'
    jpgPath = 'C:/Users/eTreatMD/Documents/MATLAB/FacialRecognition/Images'

    if not os.path.exists(tarOutputPath):
        os.makedirs(tarOutputPath)

    if not os.path.exists(imageCollectionPath):
        os.makedirs(imageCollectionPath)

    if not os.path.exists(jpgPath):
        os.makedirs(jpgPath)

##    findArchives(inputPath, tarOutputPath)
##    lookupExtract(tarOutputPath, imageCollectionPath)
    PPMtoJPGMassConvert(imageCollectionPath, jpgPath)

##    shutil.rmtree(imageCollectionPath)
##    shutil.rmtree(tarOutputPath)
    
    

def findArchives(inputPath, destinationPath):

        tar = tarfile.open(inputPath, 'r')
        counter = 1000
        
        for name in tar.getnames():
            try:
                if "images" in name and name.endswith('fa.ppm.bz2') and counter > 0:
                    tar.extract(tar.getmember(name), destinationPath)
                    counter -= 1
                elif counter == 0:
                    break
            except PermissionError:
                print("Permission Error for ", name)
                pass


        
    


def getSubdirectories(path):
    return filter(os.path.isdir, [os.path.join(f) for f in os.listdir(path)])



def lookupExtract(initialPath, destinationPath):
    
    directory = listdir(initialPath)
    
    for files in directory:
        path = os.path.join(initialPath, files)
        
        if files.endswith('fa.ppm.bz2'):
            newFile = files[:-3]
            outputPath = os.path.join(destinationPath, newFile)
            
                
            with open(outputPath, 'wb') as outFile, open(path, 'rb') as file:
                dc = bz2.BZ2Decompressor()
                for data in iter(lambda : file.read(100*1024), b''):
                    outFile.write(dc.decompress(data))

        elif os.path.isdir(path):
            lookupExtract(path, destinationPath)


def PPMtoJPGMassConvert(initialPath, destinationPath):
    files = listdir(initialPath)

    for file in files:
        path = os.path.join(initialPath, file)
        image = Image.open(path)
        JPG = file[:-3] + 'jpg'
        outPath = os.path.join(destinationPath, JPG)
        image.save(outPath)



main()


    


        
