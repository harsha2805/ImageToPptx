import os
# To avoid attribute errors if you are running a Python version above 3.9.6
import collections.abc
c = collections
c.abc = collections.abc

from pptx import Presentation
from pptx.util import Inches

# Get the name of all files in an array
path = '<file path which contains images>'
dir_list = os.listdir(path)

# Get the number of files
a = len(dir_list)

# Convert png to slides
for filename in dir_list:
    img_path = os.path.join(path, filename)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    left = top = Inches(1)
    height = Inches(3.5)
    pic = slide.shapes.add_picture(img_path, left, top, height=height)

    prs.save(f'{filename}.pptx')
