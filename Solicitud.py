from io import BytesIO
from flask import Flask, request, send_file
import io
from os import remove
from flask_cors import CORS
import json
import os
import glob
from auth.jwt import check_token


from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

app.config['TEMPLATES_AUTO_RELOAD'] = True


@app.route('/solicitud',  methods=['POST'])
@check_token

def get_data():
    data = request.data
    data = json.loads(data)
    for f in glob.iglob('.xlsx', recursive=True):
        os.remove(f)
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
    workbook = load_workbook(filename="./templates/template_sol_usuarios.xlsx")
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
    workbook.save(Tipo_de_usuario + ".xlsx")
    
    
    file_stream.seek(0)

    
    
    # excel = client.Dispatch("Excel.Application",pythoncom.CoInitialize())
    # #app.Interactive = False
    # #app.Visible = False
    # #pdfop2 = app.open(filename, "rb")
    # sheets = excel.Workbooks.Open('C:\\Users\\hp\\OneDrive\\Escritorio\\PDF_solicitud\\'+filename)
    # work_sheets = sheets.Worksheets[0]
    # pedf_T = work_sheets.ExportAsFixedFormat(0, 'C:\\Users\\hp\\OneDrive\\Escritorio\\PDF_solicitud\\'+filename)
    # sheets.close()
    return send_file(filename , as_attachment=True, )
   



if __name__ == "__main__":

    app.run(host='0.0.0.0', port=90, debug=True)