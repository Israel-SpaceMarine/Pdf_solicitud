from io import BytesIO
from flask import Flask, request, send_file
import io
from os import remove
from flask_cors import CORS
import json
import os
import glob
from pathlib import Path
#from reportlab.pdfgen import canvas
#from reportlab.lib.units import inch
from openpyxl import load_workbook
import tkinter as tk

app = Flask(__name__)
CORS(app)

app.config['TEMPLATES_AUTO_RELOAD'] = True


@app.route('/pdf',  methods=['POST'])


def get_data():
    data = request.data
    data = json.loads(data)
    #for f in glob.iglob('*.xlsx', recursive=True):
     #   os.remove(f)
    return pdf(data)


def pdf(data):

    #fileName = 'sample.pdf'
    Fecha =	data["Fecha"]
    Tipo_de_movimiento = data["TipoDeMovimiento"]
    Nombre=	data["Nombre"]
    Apellido_paterno =	data["ApellidoPaterno"]
    Apellido_materno =	data["ApellidoMaterno"]
    Nombre_de_usuario =	data["NombreUsuario"]
    Correo = data["Correo"]
    Curp = data["CUPR"]
    Rfc = data["RFC"]
    Telefono =	data["Telefono"]
    Extensión =	data["Extensión"]
    Celular = data["Celular"]
    Tipo_de_usuario = data["AccesoApp"]
    Nota = data["Nota"] 
    
    workbook = load_workbook(filename="template_sol_usuarios.xlsx")
    workbook.sheetnames
    sheet = workbook.active
    print(workbook) 
    # img = open("NL-login-pdf.png")
    # #img2 = Path("NL-login-pdf.png")
    # img3 = "NL-login-pdf.png"
    # print(img3)
    # #img = img.read()
    # import fitz
    # pdf_firma = canvas.Canvas(fileName)
    # pdf_firma.setTitle("Prueba")
    # pdf_firma.drawCentredString(200, 785, "Prueba")
    # #pdf_firma.drawInlineImage(img3, 450, 695, width=1.4*inch, height=1.4*inch)
    # #pdf_firma.drawImage(img3, 450, 695, width=1.4*2, height=1.4*2, preserveAspectRatio=True, mask='auto')
    # pdf_firma.save()
    # pdfop =  Path("sample.pdf")
    sheet["C9"] = Fecha
    sheet["C10"] = Tipo_de_movimiento
    sheet["C11"] = Nombre
    sheet["C12"] = Apellido_paterno
    sheet["C13"] = Apellido_materno
    sheet["C14"] = Nombre_de_usuario
    sheet["C15"] = Correo
    sheet["C16"] = Curp
    sheet["C17"] = Rfc
    sheet["C18"] = Telefono
    sheet["C19"] = Extensión
    sheet["C20"] = Celular
    sheet["C21"] = Tipo_de_usuario 

    filename = Fecha + " " + Tipo_de_usuario + " solicitud"+".pdf"
    file_stream = BytesIO()
    workbook.save(filename= filename)
    file_stream.seek(0)
    
    return send_file(filename, as_attachment=True, mimetype="application/pdf")
    #, download_name="solicitud.pdf"



if __name__ == "__main__":

    app.run(host='0.0.0.0', port=90, debug=True)