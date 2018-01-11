import json
import os
from os import listdir

directory = listdir()
for file in directory:
    if file.endswith('.jpg'):


        name = ''
        for c in file:
            if (c != '.'):
                name = name + c
            else:
                break

        if name[5] == '1':
            orientation = "up"
            side = "left"
        elif name[5] == '2':
            orientation = "down"
            side = "left"
        elif name[5] == '3':
            orientation = "up"
            side = "right"
        elif name[5] == '4':
            orientation = "down"
            side = "right"
            
        data = {
            "name" : str(name),
            "image" : str(file),
            "name" : str(name),
            "country":"Sri Lanka",
            "paper": "A4",
            "orientation" : orientation,
            "side" : side
            }

        output_file = open(str(name) + '.json', "w")
        json.dump(data, output_file, indent=4)
        output_file.close()
            
            
