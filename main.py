from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
import os
from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import fitz
import openpyxl



app = Flask(__name__)


wb = openpyxl.open("C:/Users/79858/Documents/Flask-App-F/static/files/data.xlsx")
sheet = wb.active
# Имя файла в котором ищу номера ярлыков
filename = 'C:/Users/79858/Documents/Flask-App-F/static/files/original.pdf'
pdf = fitz.open(filename)
# read your existing PDF
existing_pdf = PdfFileReader(open("C:/Users/79858/Documents/Flask-App-F/static/files/original.pdf", "rb"))
output = PdfFileWriter()

UPLOAD_FOLDER = 'C:/Users/79858/Documents/Flask-App-F/static/files/'
ALLOWED_EXTENSIONS = set(['pdf', 'jpg', 'jpeg', 'gif'])
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# app.config['C:\Users\79858\Documents\Flask-App-F\UPLOAD_FOLDER']

menu = [
    {"name": "Установка", "url": "install-flask"},
    {"name": "Первое приложение", "url": "first-app"},
    {"name": "Обратная связь", "url": "contact"},
]

@app.route("/")
def index():
    return render_template('index.html', menu=menu)

@app.route("/about")
def about():
    return render_template('about.html', title="О сайте", menu=menu)

@app.route("/contact", methods=["POST", "GET"])
def contact():
    if request.method == 'POST':
        print(request.form)
    return render_template('contact.html', title="Обратная связь", menu=menu)

@app.route("/base")
def base():
    return render_template('base.html', title="Base", menu=menu)

@app.route('/upload')
def upload_file():
   return render_template('upload.html')
	
@app.route('/uploader', methods = ['GET', 'POST'])
def uploader_file():
   if request.method == 'POST':
        f = request.files['file']
        f.save(os.path.join(app.config['UPLOAD_FOLDER'], f.filename))
        # Массив в котором ищу на какой странице этот ярлык
        for i in range(3, sheet.max_row+1):
            # print(i[0])
            search_term = sheet[i][1].value
            for current_page in range(len(pdf)):
                page = pdf.load_page(current_page)
                
                if page.search_for(search_term):
                    print('%s найдено на %i странице' % (search_term, current_page + 1))

                    packet1 = io.BytesIO()
                    can = canvas.Canvas(packet1, pagesize=letter)
                    can.setFont('Helvetica', 6)
                    can.rotate(90)
                    can.drawString(120, -338, sheet[i][3].value + ' (' + str(int(sheet[i][5].value)) + 'pcs)')
                    can.save()

                    #move to the beginning of the StringIO buffer
                    packet1.seek(0)
                    new_pdf1 = PdfFileReader(packet1)
                    
                    # add the "watermark" (which is the new pdf) on the existing page
                    page = existing_pdf.getPage(current_page)
                    page.mergePage(new_pdf1.getPage(0))
                    output.addPage(page)

                
        # finally, write "output" to a real file
        outputStream = open("addedindexes.pdf", "wb")
        output.write(outputStream)
        outputStream.close()
        return 'file uploaded successfully'

if __name__ == "__main__":
    app.run(debug=True)