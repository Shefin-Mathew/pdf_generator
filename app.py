from reportlab.platypus import SimpleDocTemplate,Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import pandas as pd

def header_footer(canvas,doc):
    canvas.saveState()
    
    canvas.setFont("Times-Roman","14")
    canvas.drawString(50,100,"My Notes")
    
    canvas.setFont("Times-Roman","10")
    page_number=canvas.getPageNumber()
    canvas.drawString(500,20,f"Page {page_number}")
    canvas.restoreState()
    
text_list =[]

while True:
    line = input("enter a text: ")
    if line.lower() == "end":
        break
    else:
        text_list.append(line)

styles = getSampleStyleSheet()

story = []

for text in text_list:
    story.append(Paragraph(text,styles["Normal"]))

pdf=SimpleDocTemplate("output.pdf")
pdf.build(story,onFirstPage=header_footer,onLaterPages=header_footer)
