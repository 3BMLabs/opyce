#if you have not installed python:
#	go to the run section (click on the triangle bug icon on the left)
#	click 'Run and Debug'
#	 follow instructions.

from time import sleep
from backend import main

#Launch Excel and Open Workbook
opyce = main.Opyce()

#Run Macro
Presentation = opyce.app.Activepresentation

slidenr = Presentation.Slides.Count    
slide = Presentation.slides(slidenr)
    
shape1 = slide.Shapes.AddTextbox(Orientation=0x1,Left=100,Top=100,Width=100,Height=100)
shape1.TextFrame.TextRange.Text='Hello, world'    

#Manipulate font size, name and boldness
shape1.TextFrame.TextRange.Font.Size=20
shape1.TextFrame.TextRange.Characters(1, 4).Font.Name = "Times New Roman"
shape1.TextFrame.TextRange.Font.Bold=True

def RGB(red, green, blue):
    assert 0 <= red <=255    
    assert 0 <= green <=255
    assert 0 <= blue <=255
    return red + (green << 8) + (blue << 16)

names = []
for r in range(4):
    for g in range(4):
        for b in range(4):
            #1 = rectangle
            cube = slide.Shapes.AddShape(1, r * 100, g * 100, 100, 100)
            cube.Fill.ForeColor.RGB = RGB(r * 64,g * 64,b * 64)  # Red fill color
            names.append(cube.name)
            cube.ThreeD.z = b * 100
            cube.ThreeD.depth = 100
            #remove line
            cube.Line.Visible = 0
            #cube.Line.ForeColor.RGB = 0x0000FF  # Blue border color

shape_range = slide.Shapes.Range(names)
grouped_cube = shape_range.Group()
grouped_cube.ThreeD.RotationX = 50

for i in range(-45,45):

    # Rotate the shape to the degree specified.
    grouped_cube.ThreeD.RotationY = i
    #grouped_cube.ThreeD.RotationX = i
    sleep(0.1)

    # Refresh the slide. This step is needed to redraw the screen
    # after the rotation step; Otherwise, the animation effect is
    # invisible.

#Cleanup the com reference. 
del opyce
