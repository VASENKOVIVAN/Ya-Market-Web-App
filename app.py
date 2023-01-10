from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
import os
from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import registerFontFamily
import fitz
import openpyxl
from flask_cors import CORS
from flask import *
from collections import Counter
import pandas as pd
import numpy as np

app = Flask(__name__)
cors = CORS(app)

UPLOAD_FOLDER = 'C:/Users/79858/Documents/Flask-App-F/static/files/'
DOWL_FOLDER = 'C:/Users/79858/Documents/Flask-App-F/templates/'

ALLOWED_EXTENSIONS = set(['pdf', 'jpg', 'jpeg', 'gif'])
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# df = pd.read_excel('C:/Users/79858/Documents/Flask-App-F/static/files/tttfirst_mile_shipment_orders_41681422_2022-12-21-2.xlsx')
# result = df.sort_values('testtw')
# print(df.head)
# app.config['C:\Users\79858\Documents\Flask-App-F\UPLOAD_FOLDER']

menu = [
    {"name": "Главная", "url": "index"},
    # {"name": "Не главная", "url": "first-app"},
    {"name": "Конвертация", "url": "contact"},
]

@app.route("/index")
@app.route("/")
def index():
    return render_template('contact.html', menu=menu)

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

        folder_files_saving = "C:/Users/79858/Documents/Flask-App-F/static/files/"
        folder_files_data = "C:/Users/79858/Documents/Flask-App-F/static/data/"

        f = request.files['file']
        f.save(os.path.join(app.config['UPLOAD_FOLDER'], f.filename))
        filename_pdf = f.filename

        f1 = request.files['file1']
        f1.save(os.path.join(app.config['UPLOAD_FOLDER'], f1.filename))
        filename_excel= f1.filename

        if filename_pdf[len(filename_pdf)-3:len(filename_pdf)] == "pdf":
            sheet_export_ya = openpyxl.open(folder_files_saving + filename_excel).active
            # Имя файла в котором ищу номера ярлыков
            pdf = fitz.open(folder_files_saving + filename_pdf)
            # read your existing PDF
            existing_pdf = PdfFileReader(open(folder_files_saving + filename_pdf, "rb"))
        else:
            sheet_export_ya = openpyxl.open(folder_files_saving + filename_pdf).active
            # Имя файла в котором ищу номера ярлыков
            pdf = fitz.open(folder_files_saving + filename_excel)
            # read your existing PDF
            existing_pdf = PdfFileReader(open(folder_files_saving + filename_excel, "rb"))

        sheet_sku_data_base = openpyxl.open(folder_files_data + "sku-data-base.xlsx").active
        output = PdfFileWriter()

        print("ТАК ТАК ТК: " , (sheet_export_ya.max_row))

        musical_notes = []
        for i in range(3, sheet_export_ya.max_row+1):
            musical_notes.append(int(sheet_export_ya[i][1].value))
        
        c = Counter(musical_notes)
        print("И ЧО??????????? - ", c)

        searched_pages_on_pdf = []

        # Цикл в котором пробегаю по столбцу с номерами заказов из таблицы экспорта
        for i in range(3, sheet_export_ya.max_row+1):
            # Переменная, хранящая номер заказа из таблицы экспорта
            search_sku_in_sheet_ya = sheet_export_ya[i][1].value
            # Цикл в котором пробегаю по каждой странице в pdf и ищу на какой странице этот номер заказа
            for current_page in range(len(pdf)):
                page = pdf.load_page(current_page)
                # Если я нашел страницу на которой этот номер заказа
                if page.search_for(search_sku_in_sheet_ya):
                    print('\n' + '%s найдено на %i странице' % (search_sku_in_sheet_ya, current_page + 1))
                    searched_pages_on_pdf.append(current_page)
                    print("Найдены страницы: ", searched_pages_on_pdf)
                    # Беру sku этого заказа и ищу его в базе транслейта
                    for j in range(2, sheet_sku_data_base.max_row+1):
                        # Если я нашел этот sku в базе, то пишу его на странице pdf
                        if (sheet_export_ya[i][3].value) == (sheet_sku_data_base[j][1].value):

                            # print(
                            #     "Это 1: ", (sheet_export_ya[i][3].value),"\n",
                            #     "Это 2: ", (sheet_sku_data_base[j][1].value),"\n",
                            #     "Это 3: ", str((sheet_sku_data_base[j][2].value)),"\n",
                            # )

                            packet1 = io.BytesIO()
                            can = canvas.Canvas(packet1, pagesize=letter)
                            pdfmetrics.registerFont(TTFont('Roboto', 'Roboto-Medium.ttf'))
                            can.setFont('Roboto', 6)
                            can.rotate(90)
                            print(
                                        "Цикл: " + 
                                        str(int(i)) 
                                        # ", " + 
                                        # (sheet_export_ya[i][1].value) + 
                                        # ", " + 
                                        # (sheet_export_ya[i+1][1].value)
                                    )

                            can.drawString(
                                120, 
                                -338, 
                                str((sheet_sku_data_base[j][2].value)) + 
                                ' (' + 
                                str(int(sheet_export_ya[i][5].value)) + 
                                'pcs)'
                            )
                            print(
                                str((sheet_sku_data_base[j][2].value)) + 
                                ' (' + 
                                str(int(sheet_export_ya[i][5].value)) + 
                                'pcs)'
                            )

                            can.save()

                            #move to the beginning of the StringIO buffer
                            packet1.seek(0)
                            new_pdf1 = PdfFileReader(packet1)
                            
                            # add the "watermark" (which is the new pdf) on the existing page
                            page = existing_pdf.getPage(current_page)
                            page.mergePage(new_pdf1.getPage(0))
                            output.addPage(page)
                # else:
                #     page = existing_pdf.getPage(current_page)
                #     page.mergePage(new_pdf1.getPage(0))
                #     output.addPage(page)


        for current_page in range(len(pdf)):
            # for searched_pages in searched_pages_on_pdf:
            if current_page not in searched_pages_on_pdf:
                print("Этого нет:",  current_page)

                packet3 = io.BytesIO()
                can = canvas.Canvas(packet3, pagesize=letter)
                pdfmetrics.registerFont(TTFont('Roboto', 'Roboto-Medium.ttf'))
                can.setFont('Roboto', 6)
                can.rotate(90)

                can.drawString(
                    120, 
                    -338,
                    'Нет данных'
                )

                can.save()

                #move to the beginning of the StringIO buffer
                packet3.seek(1)
                new_pdf3 = PdfFileReader(packet3)

                page = existing_pdf.getPage(current_page)
                page.mergePage(new_pdf3.getPage(0))
                output.addPage(page)


        # finally, write "output" to a real file
        outputStream = open("addedindexes.pdf", "wb")
        output.write(outputStream)
        outputStream.close()
        return render_template('download.html')

@app.route('/download')
def download():
    
    # dist_dir = 'C:/Users/79858/Documents/Flask-App-F/templates/'
    dist_dir = 'C:/Users/79858/Documents/Flask-App-F/'

    entry = os.path.join(dist_dir, 'addedindexes.pdf')
    return send_file(entry)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
    # app.run(host = 'localhost', debug = True)