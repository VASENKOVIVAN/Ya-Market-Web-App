from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
import os
from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import fitz
import openpyxl
from flask_cors import CORS
from flask import *

app = Flask(__name__)
cors = CORS(app)





UPLOAD_FOLDER = 'C:/Users/79858/Documents/Flask-App-F/static/files/'
ALLOWED_EXTENSIONS = set(['pdf', 'jpg', 'jpeg', 'gif'])
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# app.config['C:\Users\79858\Documents\Flask-App-F\UPLOAD_FOLDER']

menu = [
    {"name": "Главная", "url": "install-flask"},
    # {"name": "Не главная", "url": "first-app"},
    {"name": "Конвертация", "url": "contact"},
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
    return render_template('contact.html', title="Конвертация", menu=menu)

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
        f1 = request.files['file1']
        f1.save(os.path.join(app.config['UPLOAD_FOLDER'], f1.filename))
        sheet = openpyxl.open("C:/Users/79858/Documents/Flask-App-F/static/files/data.xlsx").active
        # Имя файла в котором ищу номера ярлыков
        pdf = fitz.open('C:/Users/79858/Documents/Flask-App-F/static/files/original.pdf')
        # read your existing PDF
        existing_pdf = PdfFileReader(open("C:/Users/79858/Documents/Flask-App-F/static/files/original.pdf", "rb"))
        output = PdfFileWriter()
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
        return render_template('download.html')

@app.route('/download')
def download():
    filename = 'addedindexes.pdf'
    return send_file(filename,as_attachment=True)

if __name__ == "__main__":
    app.run(port=80)