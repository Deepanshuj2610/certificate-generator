import sys
import os
import pandas as pd
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont

def generate(name,cwd):
    try:  
        img  = Image.open("template.jpeg")   #  template
    except IOError: 
        pass
    draw = ImageDraw.Draw(img)
    selectFont = ImageFont.truetype(cwd + '/text_design/'+"Pacifico.ttf", size=25)   #  font selection
    draw.text((455,200), name , font=selectFont, fill=(110, 110, 110))      #  430,360 is the x,y co-ordinates 105,105,105 is the code for font colour
    img.save(cwd + '/final_certificates/' + name + '.jpeg')        #  certificate will be saved with the name of attendee
    
cwd = os.getcwd()
cars = pd.read_csv('StateNames.csv')  #  txt file containing the names of the attendee
cases = cars.Name.tolist()
for i in cases:
    i=i.replace("\n","")
    generate(i,cwd)

