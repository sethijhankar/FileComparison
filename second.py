import pandas as pd
from fpdf import FPDF
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.acroform import AcroForm
from reportlab.lib import colors
import os

def add_password(input_pdf_path, output_pdf_path, password):
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    dob_date = password.strftime('%Y%m%d')
    writer.encrypt(user_password=dob_date, owner_password=dob_date, use_128bit=True)
    with open(output_pdf_path, "wb") as f:
        writer.write(f)

def create_text_field_pdf(output_pdf_path):
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    form = AcroForm(c)

    # Label for name field
    c.setFont("Helvetica", 12)
    c.drawString(10, 46, "Name:")

    # Text field for name
    form.textfield(
        name='name_field',
        tooltip='Enter your name here',
        x=60, y=40,
        borderColor=colors.black,
        fillColor=colors.white,
        width=200,
        height=20,
        textColor=colors.black,
        forceBorder=True
    )

    # Label for date field
    c.drawString(300, 46, "Date:")

    # Text field for date
    form.textfield(
        name='date_field',
        tooltip='Enter the date here',
        x=350, y=40,
        borderColor=colors.black,
        fillColor=colors.white,
        width=200,
        height=20,
        textColor=colors.black,
        forceBorder=True
    )

    c.save()

def merge_pdfs(input_pdf_path, text_field_pdf_path, output_pdf_path):
    original_pdf = PdfReader(input_pdf_path)
    text_field_pdf = PdfReader(text_field_pdf_path)
    writer = PdfWriter()

    for page_num in range(len(original_pdf.pages)):
        page = original_pdf.pages[page_num]
        if page_num == 0:
            text_field_page = text_field_pdf.pages[0]
            page.merge_page(text_field_page)
        writer.add_page(page)

    with open(output_pdf_path, "wb") as output_file:
        writer.write(output_file)

def excel_to_pdf(excel_file_path, image_path):
    df = pd.read_excel(excel_file_path, engine='openpyxl', usecols=['Name', 'Department', 'Email', 'Contact', 'Date of birth', 'Course', 'College'])

    today_date = datetime.now().strftime('%Y-%m-%d')
    if not os.path.exists(today_date):
        os.makedirs(today_date)
        
    password_protected_folder = os.path.join(today_date, 'password_protected')
    if not os.path.exists(password_protected_folder):
        os.makedirs(password_protected_folder)

    for index, row in df.iterrows():
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()

        pdf.set_font("Times", "B", size=24)
        pdf.set_xy(0, 10)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(210, 10, text="Résumé", align='C')

        pdf.set_font("Times", size=20)
        pdf.set_xy(0, 20)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(210, 10, text="uequdb eoihwf ioqhf8", align='C')

        pdf.set_font("Times", size=18)
        pdf.set_xy(0, 30)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(210, 10, text="udicue89 8ewqh9 jqedhoqwg ouehq8", align='C')

        pdf.set_font("Times", size=16)
        pdf.set_xy(10, 45)
        pdf.set_text_color(0, 0, 150) 
        pdf.cell(200, 10, text=f"Name: {row['Name']}", align='L')

        image_x = 110  
        image_y = 45   
        try:
            pdf.image(image_path, x=image_x, y=image_y, w=90) 
        except RuntimeError as e:
            print(f"Error loading image: {e}")

        pdf.set_font("Times", size=10)
        pdf.set_xy(10, 60)
        pdf.set_text_color(0, 0, 150) 
        pdf.cell(200, 10, text=f"Course: {row['Course']}", align='L')
        pdf.set_xy(10, 70)
        pdf.cell(200, 10, text=f"College: {row['College']}", align='L')

        pdf.set_font("Times", size=12)
        pdf.set_xy(15, 110)
        pdf.set_text_color(0, 0, 0) 
        pre_written_text = (
            "{name} is a student in the {department} department. {name} is currently enrolled in the {course} course at "
            "{college}. {name} has shown exceptional skills and dedication in their field of study. They have participated "
            "in various projects and extracurricular activities, showcasing their ability to apply theoretical knowledge "
            "to practical situations. {name} is also known for their strong communication skills and teamwork, making them "
            "a valuable member of any group or organization. {name} continues to strive for excellence in their academic "
            "and personal endeavors."
        )
        filled_text = pre_written_text.format(name=row['Name'], department=row['Department'], course=row['Course'], college=row['College'])
        pdf.multi_cell(0, 10, text=filled_text, align='L')

        y_after_text = pdf.get_y()

        new_image_path = "img2.jpeg"  
        new_image_x = 100  
        new_image_y = y_after_text + 10  
        try:
            pdf.image(new_image_path, x=new_image_x, y=new_image_y, w=100)  # Adjust width as needed
        except RuntimeError as e:
            print(f"Error loading new image: {e}")

        pdf.set_font("Times", size=10)
        email_y = new_image_y + 40  
        pdf.set_xy(10, email_y)  
        pdf.set_text_color(0, 0, 150)  
        pdf.cell(0, 10, text=f"Email: {row['Email']}", align='L')
        pdf.set_xy(10, email_y + 10)
        pdf.cell(0, 10, text=f"Phone: {row['Contact']}", align='L')

        pdf.set_fill_color(51, 171, 249)  # blue
        pdf.rect(0, 292, 80, 5, 'F')
        pdf.set_fill_color(255, 165, 0)   # yellow
        pdf.rect(71.6, 292, 66.6, 5, 'F')
        pdf.set_fill_color(255, 0, 0)     # red
        pdf.rect(138.2, 292, 80, 5, 'F')
        
        output_pdf_path = os.path.join(today_date, f"{row['Name']}.pdf")
        pdf.output(output_pdf_path)

        text_field_pdf_path = os.path.join(today_date, f"text_field_{index}.pdf")
        create_text_field_pdf(text_field_pdf_path)

        final_output_pdf_path = os.path.join(today_date, f"{row['Name']}_final.pdf")
        merge_pdfs(output_pdf_path, text_field_pdf_path, final_output_pdf_path)

        print(f"Final PDF generated: {final_output_pdf_path}")

        if not pd.isnull(row['Date of birth']):
            dob = pd.to_datetime(row['Date of birth'])
            password_output_path = os.path.join(password_protected_folder, f"{row['Name']}_final_protected.pdf")
            add_password(final_output_pdf_path, password_output_path, dob)
            print(f"Password-protected PDF generated: {password_output_path}")

excel_file_path = 'f1.xlsx'
image_path = 'img.jpeg'

excel_to_pdf(excel_file_path, image_path)
