

from io import DEFAULT_BUFFER_SIZE
import docx
import os
from docx2pdf import convert

import time
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
import unicodedata
import time
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
import unicodedata




def importExcel(fileroute):
    #TODO la ruta de la base no puede ser estatia
    if not fileroute:  # base
        fileroute = "documents/alumnosSampleFilex.xlsx"

    base = load_workbook(fileroute)
    baseHoja = base.active

    print(fileroute)
    return baseHoja

def getNombreBase(hojaExcel, pos):
    try:
        nombre = unicodedata.normalize('NFD', ((hojaExcel['C' + str(pos)].value)))
    except:
        nombre = "-"
    try:
        apellidoPaterno = unicodedata.normalize('NFD', ((hojaExcel['D' + str(pos)].value)))
    except:
        apellidoPaterno = "-"
    try:
        apellidoMaterno = unicodedata.normalize('NFD', ((hojaExcel['E' + str(pos)].value)))
    except:
        apellidoMaterno = "-"

    fullName = [nombre, apellidoPaterno, apellidoMaterno]
    return fullName;

def getNombreFilex(hojaExcel,pos):
    try:
        fullNombre = unicodedata.normalize('NFD', ((hojaExcel['G' + str(pos)].value)))
    except:
        fullNombre =""
    print(fullNombre)
    return fullNombre

def getNivelacion(hojaExcel, pos):
    try:
        nivel =  str(hojaExcel['D' + str(pos)].value)
    except:
        nivel = ""
    return nivel

def getHoras(hojaExcel, pos):
    try:
        horas = unicodedata.normalize('NFD', ((hojaExcel['B' + str(pos)].value)))
    except:
        horas = ""
    return horas

def getDias(hojaExcel,pos):
    try:
        dias = unicodedata.normalize('NFD', ((hojaExcel['A' + str(pos)].value)))
    except:
        dias = ""
    return dias

def getCiclo(hojaExcel,pos):
    try:
        ciclo = unicodedata.normalize('NFD', ((hojaExcel['F' + str(pos)].value)))
    except:
        ciclo = ""
    return ciclo


def scanDocument(i):
    # RECORRIENDO EL ARCHIVO EN BUSCA DE SUSTITUCION
    print(i)
    try:
        docRoute = docx.Document("DOCUMENTS/machotefilex.docx")
        #print(doc)
    except:
        print("ERROOOOOOOOOOOOOOOOS")


    excelRoute = importExcel("")
    print(excelRoute)
    for paragraph in docRoute.paragraphs:
        #print(paragraph.text)
        font = paragraph.style.font
        font.name = 'Arial'
        if ("$DIAS$") in paragraph.text:
            dias = getDias(excelRoute, i)
            paragraph.text = paragraph.text.replace("$DIAS$", dias)
        if ("$GENERO$") in paragraph.text:
            # PENDIENTE
            gender = "el"
            paragraph.text = paragraph.text.replace("$GENERO$", gender)
        if ("$NOMBRE$") in paragraph.text:
            name = getNombreFilex(excelRoute, i)
            paragraph.text = paragraph.text.replace("$NOMBRE$", name)
        if ("$NIVELACION$") in paragraph.text:
            nivel = getNivelacion(excelRoute, i)
            nivelDocu = nivel
            paragraph.text = paragraph.text.replace("$NIVELACION$", nivel)
        if ("$HORAS$") in paragraph.text:
            horas = getHoras(excelRoute, i)
            paragraph.text = paragraph.text.replace("$HORAS$", horas)
        if ("$CICLO$") in paragraph.text:
            ciclo = getHoras(excelRoute, i)
            paragraph.text = paragraph.text.replace("$CICLO$", ciclo)

        # print(paragraph.text)
        # print("---")

    docRoute.save("documents/generadas/temp" + str(i) + ".docx")
    time.sleep(0.3)
    convert("documents/generadas/temp" + str(i) + ".docx", "documents/generadas/" + str(i) + "doc.pdf")
    time.sleep(0.3)
    os.remove("documents/generadas/temp"+str(i)+".docx")
    time.sleep(0.3)
    print("finalizado")
    time.sleep(3)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':


    #print(getNombreFilex(hojaExcel,2))


    scanDocument(2)
    scanDocument(3)
    scanDocument(4)


