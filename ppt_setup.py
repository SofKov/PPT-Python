#Import Presentation from pptx module (pip install module)
from pptx import Presentation
from pptx.util import Inches

#Create PPT Presentation 
prs = Presentation()

#Assign layout type to slide 1 - Main Title
slide1_layout = prs.slide_layouts[0]

#Add slide 1 to presentation
slide1 = prs.slides.add_slide(slide1_layout)

#Main title placeholder
title1 = slide1.shapes.title

#Subtitle placeholder
subtitle1 = slide1.placeholders[1]

#Adding text to main title placeholder
title1.text = "print('Python')"
subtitle1.text = "For those seeking a simpler way of life..."

###############################################################################

#Assigning layout to slide 2 - Title and Content
slide2_layout = prs.slide_layouts[1]

#Adding slide 2 to presentation
slide2 = prs.slides.add_slide(slide2_layout)
shapes = slide2.shapes

#Slide 2 title placeholder
title2 = slide2.shapes.title
b_shape = shapes.placeholders[1]

#Adding text to slide 2 title placeholder
title2.text = "What is Python?"

#Adding bullet point 1
tf = b_shape.text_frame
tf.text = "Released in 1991"

#Adding bullet point 2
tf2 = tf.add_paragraph()
tf2.text = "Python is an object-oriented programming language (OOP)"
tf2.level = 0

#Adding bullet point 3
tf3 = tf.add_paragraph()
tf3.text = "Beginner-friendly due to its simpler syntax - write programs with fewer lines of code"
tf3.level = 0

#Adding bullet point 4
tf4 = tf.add_paragraph()
tf4.text = "Not only used by software engineers - also by mathmatecians, data analysts, scientists, accountants and network engineers"
tf4.level = 0

#Adding bullet point 5
tf5 = tf.add_paragraph()
tf5.text = "One of the most popular programming languages in the world!"
tf5.level = 0

###############################################################################

#Assigning layout to slide 3 - Title and Empty space
slide3_layout = prs.slide_layouts[5]

#Adding slide 3 to presentation
slide3 = prs.slides.add_slide(slide3_layout)

#Slide 3 title placeholder
title3 = slide3.shapes.title

#Adding text to slide 3 title
title3.text = "Are you ready to RUUUMBLE?! - Python vs JavaScript"













#Saves above to .pptx file called CN_PPT_Sofia
prs.save('CN_PPT_Sofia.pptx')