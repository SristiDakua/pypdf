The PdfWriter Class
-------------------

.. autoclass:: PyPDF2.PdfWriter
    :members:
    :undoc-members:
    :show-inheritance:
    # Import the required libraries for creating a PowerPoint presentation
from pptx import Presentation
from pptx.util import Inches

# Create a new PowerPoint presentation
prs = Presentation()

# Add slide 1 - Introduction
slide1 = prs.slides.add_slide(prs.slide_layouts[0])
title1 = slide1.shapes.title
title1.text = "Farming Methods Used in Agriculture"
subtitle1 = slide1.placeholders[1]
subtitle1.text = "An Overview"
picture1 = slide1.shapes.add_picture('farm.jpg', Inches(1), Inches(1.5), width=Inches(8), height=Inches(4.5))

# Add slide 2 - Traditional Farming Methods
slide2 = prs.slides.add_slide(prs.slide_layouts[1])
title2 = slide2.shapes.title
title2.text = "Traditional Farming Methods"
bullet_slide2 = slide2.shapes.placeholders[1].text_frame
bullet_slide2.text = "Advantages:\n- low cost\n- use of organic methods\nDisadvantages:\n- time-consuming\n- low yield"
picture2 = slide2.shapes.add_picture('traditional_farming.jpg', Inches(0.5), Inches(2), width=Inches(5), height=Inches(3.5))

# Add slide 3 - Modern Farming Methods
slide3 = prs.slides.add_slide(prs.slide_layouts[1])
title3 = slide3.shapes.title
title3.text = "Modern Farming Methods"
bullet_slide3 = slide3.shapes.placeholders[1].text_frame
bullet_slide3.text = "Advantages:\n- high yield\n- efficient\nDisadvantages:\n- expensive\n- use of pesticides"
picture3 = slide3.shapes.add_picture('modern_farming.jpg', Inches(0.5), Inches(2), width=Inches(5), height=Inches(3.5))

# Add slide 4 - Sustainable Farming Methods
slide4 = prs.slides.add_slide(prs.slide_layouts[1])
title4 = slide4.shapes.title
title4.text = "Sustainable Farming Methods"
bullet_slide4 = slide4.shapes.placeholders[1].text_frame
bullet_slide4.text = "Advantages:\n- environmentally friendly\n- long-term viability\nDisadvantages:\n- initial investment\n- may require more labor"
picture4 = slide4.shapes.add_picture('sustainable_farming.jpg', Inches(0.5), Inches(2), width=Inches(5), height=Inches(3.5))

# Add slide 5 - Conclusion
slide5 = prs.slides.add_slide(prs.slide_layouts[1])
title5 = slide5.shapes.title
title5.text = "Conclusion"
bullet_slide5 = slide5.shapes.placeholders[1].text_frame
bullet_slide5.text = "Choosing the right farming method depends on the goals of the farm.\n\nSupporting sustainable farming practices is important for a healthy environment and long-term viability of the farm."
picture5 = slide5.shapes.add_picture('happy_family.jpg', Inches(0.5), Inches(1.5), width=Inches(8), height=Inches(4))

# Save the PowerPoint presentation
prs.save('farming_methods.pptx')
