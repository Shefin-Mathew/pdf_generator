from reportlab.platypus import SimpleDocTemplate,Paragraph,PageBreak
from reportlab.lib.styles import getSampleStyleSheet
import pandas as pd
from docx import Document
import subprocess

def footer(canvas,doc):
    canvas.setFont("Times-Roman",12)
    canvas.line(50,40,550,40)
    page_number=canvas.getPageNumber()
    canvas.drawString(500,20,f"Page {page_number}")
    

def plain_page(canvas,doc):
    footer(canvas,doc)
    
def rules_page(canvas,doc):
    y = 700
    while y>=60:
        canvas.line(50,y,550,y)
        y=y-25
    footer(canvas,doc)
    
styles=getSampleStyleSheet()   

styles["Title"].alignment=1

story= []

print("\n Choose Option: ")
print("1 → Convert DOCX to PDF")
print("2 → Generate PDF from CSV")
print("3 → Combine CSV + DOCX") 
print("4 → Manual text to PDF")

choice =input("Enter Option 1/2/3/4: ")
design = input("Choose Page Design (1- Plain,2- Ruled ): ")

if choice == "1":
    doc_path= input("enter doc file path: ")
    subprocess.run(
        [
            "Libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            doc_path
        ]
    )
    print("Docx is converted to pdf.")
    exit()

elif choice == "2":
    csv_path = input("Enter csv fiile path: ")
    df=pd.read_csv(csv_path)
    for index,row in df.iterrows():
        topics = row["Topic"]
        pages=row["Pages"]
        
        story.append(Paragraph(topics,styles["Title"]))
        story.append(PageBreak())
        for i in range(pages -1):
            story.append(PageBreak)

elif choice == "3":
    csv_path = input("Enter csv fiile path: ")
    doc_path= input("enter doc file path: ")
    df =pd.read(csv_path)
    document = Document(doc_path)
    
    doc_text =[]
    
    for p in document.paragraphs:
        doc_text.append(p.text)
    
    for index,row in df.iterrows():
        topic = row["Topic"]
        story.append(Paragraph(topic,styles["Title"]))
    
    for text in doc_text:
        story.append(Paragraph(text,styles["Normal"]))
        styles["Normal"]
        story.append(PageBreak())
        
elif choice == "4":

    text_list = []

    while True:
        line = input("Enter text: ")

        if line.lower() == "end":
            break
        else:
            text_list.append(line)

    for text in text_list:
        story.append(Paragraph(text, styles["Normal"]))
        
pdf = SimpleDocTemplate("output.pdf")
pdf.topMargin = 35


if design == "2":
    pdf.build(story,onFirstPage=rules_page,onLaterPages=rules_page)
    print("pdf generated successfully")
else:
    pdf.build(story,onFirstPage=plain_page,onLaterPages=plain_page)
    print("pdf generated successfully")    