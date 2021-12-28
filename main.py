import os

from pptx import Presentation
from pptx.util import Inches

#to get the name of all files in an array
path = 'C:/Users/harsh/PycharmProjects/pythonProject/pdftoppt/img'
dir_list: list[str] = os.listdir(path)

# to know the number of files
dir_name = 'C:/Users/harsh/PycharmProjects/pythonProject/pdftoppt/img'
a = len(os.listdir(dir_name))

# coverting png to slides
i = 0
for i in range(a):
    img_path: str = 'C:/Users/harsh/PycharmProjects/pythonProject/pdftoppt/img/' + dir_list[i]
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    # to print  bigger ppt
    left = top = Inches(1)
    height = Inches(5.5)
    pic = slide.shapes.add_picture(img_path, left, top, height=height)
    #saving files in the name of the dir_name[i]
    prs.save(dir_list[i] + '.pptx')
    i = i + 1

"""    
#to print smaller ppt
left = top = Inches(1)
pic = slide.shapes.add_picture(img_path, left, top)
"""
