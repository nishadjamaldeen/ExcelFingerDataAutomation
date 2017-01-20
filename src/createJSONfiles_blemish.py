import json
import os
from os import listdir

directory = listdir()
for file in directory:
    if file.endswith('.tif'):
        name = ''
        for c in file:
            if (c != '.'):
                name = name + c
            else:
                break
        data = {
            "days" : "10",
            "image" : str(file),
            "name" : str(name)
            }

        output_file = open(str(name) + '.json', "w")
        json.dump(data, output_file, indent=4)
        output_file.close()
            
            
