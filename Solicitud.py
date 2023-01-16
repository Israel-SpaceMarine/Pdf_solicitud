from io import BytesIO
from flask import Flask, request, send_file
import io
from os import remove
from flask_cors import CORS
import json
import os
import glob
from pathlib import Path

import pythoncom
from win32com import client

from openpyxl import load_workbook

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
    print(Nota)
    workbook = load_workbook(filename="template_sol_usuarios.xlsx")
    workbook.sheetnames

    sheet = workbook.active
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
    

    filename = Tipo_de_usuario + ".xlsx"
    file_stream = BytesIO()
    workbook.save(filename)
    pdfop =  Path(filename)
    pdfop2 = open(filename, "w")
    print(type(pdfop2))
    file_stream.seek(0)

    
    # excel = client.Dispatch("Excel.Application" ,pythoncom.CoInitialize())
    # sheets = excel.workbook.Open(pdfop)
    # work_sheets = sheets.workbook[0]
    # pdfnew = work_sheets.ExportAsFixedFormat(0, pdfop)
    # import pandas as pd
    # import pdfkit
    # df = pd.read_excel(filename, index_col=0)
    # df.head()
    # f = open('exp.html','b')
    # a = df.to_html()
    # f.write(a)
    # f.close()
    # pdfkit.from_file('exp.html', 'example.pdf')

    return send_file(pdfop2, as_attachment=True, mimetype="application/pdf", download_name="solicitud.pdf")
    #,



if __name__ == "__main__":

    app.run(host='0.0.0.0', port=90, debug=True)