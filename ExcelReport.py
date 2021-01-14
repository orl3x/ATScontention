import openpyxl as xl
from openpyxl.styles import Font, Fill, Alignment
import datetime
import os
import getpass

path = "C:/Users/Pruebas/Desktop/REPORTES/"

name = datetime.datetime.today().strftime("%d-%m-%Y")+".xlsx"

def createNewWorkbook():
    os.system("taskkill /im EXCEL.EXE -f")
    global name
    wb = xl.Workbook()
    ws = wb.active
    ws.title = "ATS1 REPORT"
    ws.column_dimensions["A"].width=30
    ws.column_dimensions["G"].width = 20
    ws.column_dimensions["C"].width = 15

    #MAIN TITLE
    ws.merge_cells("A1:D1")
    ws.merge_cells("G1:H1")
    ws.merge_cells("K1:L1")
    ws["A1"] = "REPORTE DIARIO - CONTENCIÓN DE ATS1"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A1"].alignment = Alignment(horizontal="center")

    #SERIAL COLUMN
    ws["A4"] = "SERIAL"
    ws["A4"].font = Font(bold=True)

    #STATUS COLUMN
    ws["B4"] = "STATUS"
    ws["B4"].font = Font(bold=True)

    #INFO COLUMN
    ws["C4"] = "INFO"
    ws["C4"].font = Font(bold=True)

    # INFO COLUMN
    ws["D4"] = "2ND TIME"
    ws["D4"].font = Font(bold=True)

    #SUB RESULT
    ws["G1"] = "SUB RESULT"
    ws["G1"].font = Font(bold=True)
    ws["G1"].alignment = Alignment(horizontal="center")

    #RESULT
    ws["k1"] = "RESULT"
    ws["k1"].font = Font(bold=True)
    ws["k1"].alignment = Alignment(horizontal="center")

    # TOTAL PASSED ROW
    ws["G2"] = "PASSED:"
    ws["G2"].font = Font(bold=True,color='259e15')

    #TOTAL PASSED RESULT
    ws["H2"] = '=COUNTIF(B:B,"PASSED")'

    # TOTAL FAILED ROW
    ws["G3"] = "FAILED:"
    ws["G3"].font = Font(bold=True,color='cf1729')

    #TOTAL FAILED RESULT
    ws["H3"] = '=COUNTIF(B:B,"FAILED")'

    #2ND TRY PASSED ROW
    ws["G4"] = "PASSED in 2nd Try:"
    ws.column_dimensions["K"].width=20
    ws["G4"].font = Font(bold=True,color='259e15')

    #TOTAL FAILED 2ND TRY PASSED
    ws["H4"] = '=COUNTIF(D:D,"PASSED")'

    # TOTAL PASSED UNITS
    ws["K2"] = "TOTAL PASSED UNITS:"
    ws.column_dimensions["K"].width = 25
    ws["K2"].font = Font(bold=True, color='259e15')

    # TOTAL PASSED RESULT
    ws["L2"] = '=H2+H4'

    #TOTAL FAILED UNITS
    ws["K3"] = "TOTAL FAILED UNITS:"
    ws["K3"].font = Font(bold=True,color='cf1729')

    # TOTAL FAILED RESULT
    ws["L3"] = '=H3-H4'

    #TOTAL UNITS OF THE DAY
    ws["K4"] = "TOTAL UNITS PROCESSED:"
    ws["K4"].font = Font(bold=True,color='1410fe')

    # TOTAL UNITS OF THE DAY RESULT
    ws["L4"] = '=H2+H3'



    #SAVE WORKBOOK
    saveWB(wb)


def saveWB(wb):
    wb.save("C:/Users/Pruebas/Desktop/REPORTES/"+name)
    print("Saved succesfully")

def isFileExisting():
    global name
    global path
    return os.path.isfile(path+name)

def writeToWorkbook(serial, status):
    os.system("taskkill /im EXCEL.EXE -f")
    if status == 1:
        statusLabel = "PASSED"
        font = "Font(color='259e15')"

    else:
        statusLabel = "FAILED"
        font = "Font(color='cf1729')"

    if status == 1:
        infoLabel = " "
    elif status == 2:
        infoLabel = "FAILED IN ATS1"
    else:
        infoLabel = "NO ATS1"

    global name
    if isFileExisting():
        wb = xl.load_workbook(path+name)
    else:
        createNewWorkbook()

    wb = xl.load_workbook(path+name)
    ws = wb.active
    searchResult = isSerialFound(serial, ws)
    if searchResult == -1:
        print("Serial Guardado")
        serialRow = "A" + str(ws.max_row + 1)
        print(serialRow)
        statusRow = "B" + str(ws.max_row + 1)
        print(statusRow)
        infoRow = "C" + str(ws.max_row + 1)
        print(infoRow)
        ws[serialRow] = str(serial)

        ws[statusRow] = statusLabel
        ws[statusRow].font = eval(font)

        ws[infoRow] = infoLabel


        saveWB(wb)
        print("Files saved succesfully")

    elif searchResult > -1:
        foundRow = "B"+str(searchResult)
        print("Found row is "+foundRow)
        if ws[foundRow].value == "FAILED" and status == 1:
            print("Dentro de if")
            secondTimeCell = "D" + str(searchResult)
            ws[secondTimeCell] = "PASSED"
            ws[secondTimeCell].font = Font(color='259e15')
            print("Second test passed")
            saveWB(wb)
        else:
            print("Serial Repeated")


def isSerialFound(serial, ws):
    for row in range(3,ws.max_row+1):
        currentRow = "A"+str(row)
        if ws[currentRow].value == str(serial):
            print("Encontrado")
            return row
    print("No encontrado")
    return -1

# def getFailed():
#     if isFileExisting():
#         wb = xl.load_workbook(str(path+name), data_only=True)
#         ws = wb.active
#         return ws["H2"].value
#     else: return ""
#
# def getPassed():
#     if isFileExisting():
#         wb = xl.load_workbook(str(path+name), read_only=True)
#         ws = wb.active
#         return ws["H1"].value
#     else: return ""





