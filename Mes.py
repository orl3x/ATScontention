from PAG import *
import os
import pyperclip
import tkinter as tk
import tkinter.ttk as ttk
import ExcelReport

#serialNumber = "2045mm001234"
firstTimeFlagGUI = True
firstTimeFlagMes = True
firstTimeFlagMes2 = True
driverStatus = 1
previousScann = ""
driverModel = ""

def mainWindow():
    root = tk.Tk()
    root.title("Contención ATS1")
    root.attributes('-fullscreen', True)
    root.iconbitmap("shield.ico")


    # #TOTAL PASSED LABEL
    # totalPassedQty = ExcelReport.getPassed()
    # totalPassedLabel = tk.Label(root, text="PASSED: "+str(totalPassedQty))
    # totalPassedLabel.config(font=("ARIAL", 15), fg="green")
    #
    # #TOTAL FAILED LABEL
    # totalFailedQty = ExcelReport.getFailed()
    # totalFailedLabel = tk.Label(root, text="FAILED:   "+str(totalFailedQty))
    # totalFailedLabel.config(font=("ARIAL", 15), fg="red")

    #ENTRY LABEL
    Label = tk.Label(root, text="Escaneé el número de serie:")
    Label.config(font=("ARIAL",35))
    global scannedSerialNumber
    scannedSerialNumber = tk.StringVar()


    #ENTRY TEXTBOX
    textBox = tk.Entry(root, textvariable=scannedSerialNumber, width= 20 )
    textBox.config(font=("ARIAL",39), bd=4)

    #MAIN WINDOW TITLE LABEL
    mainWindowATScheckTitle = tk.Label(root, text="ATS Check ✓")
    mainWindowATScheckTitle.config(font=("TAHOMA", 100, "bold"), fg="blue")
    mainWindowTitleLabel = tk.Label(root, text="VALIDACIÓN\nDE ATS 1")
    mainWindowTitleLabel.config(font=("TAHOMA",70), fg="GREEN")
    serialLabel = tk.Label(root, text="S/N: {}".format(previousScann.upper()))

    #VALIDATION FOR PASSED OR FAILED FLAG
    if firstTimeFlagGUI is False:
        if driverStatus == 1:
            statusLabel = tk.Label(root, text="PASA")
            statusLabel.config(font=("ARIAL",250), fg="green")
            serialLabel.config(font=("ARIAL", 40), fg="green")
            ExcelReport.writeToWorkbook(driverModel,previousScann, True)

        else:
            statusLabel = tk.Label(root, text="NO\nPASA")
            statusLabel.config(font=("ARIAL", 150), fg="red")
            serialLabel.config(font=("ARIAL", 40), fg="red")
            ExcelReport.writeToWorkbook(driverModel, previousScann, False)

    def closeProgram():
        alert = pag.confirm( text="¿Seguro que desea salir?", title="Confirmar salida", buttons=["Ok","Cancelar"], icon="shield.ico")
        if(alert == "Ok"):
            root.destroy()
        else:
            return

    exitButton = tk.Button(root, text="SALIR", command=closeProgram)
    exitButton.config(font=("ARIAL", 15), fg="red", relief="groove")

    #DEFINE ENTER EVENT
    def enterEvent(event):
        global firstTimeFlagGUI
        if firstTimeFlagGUI is not True:
            pag.hotkey("alt","tab")
        root.destroy()
        firstTimeFlagGUI = False
        Mes()
    #CHECK IF FIRST TIME RUNNING PROGRAM
    if firstTimeFlagGUI is False:
        statusLabel.pack()
        serialLabel.pack()
    else:
        tk.Label(root, text=" ").pack()
        tk.Label(root, text=" ").pack()
        mainWindowATScheckTitle.pack()
        #mainWindowTitleLabel.pack()
        tk.Label(root, text=" ").pack()
        tk.Label(root, text=" ").pack()
        tk.Label(root, text=" ").pack()
    Label.pack()
    textBox.pack()
    root.bind('<Return>', enterEvent)
    root.focus_force()
    textBox.focus()
    tk.Label(root,text=" ").pack()
    #totalPassedLabel.pack()
    #totalFailedLabel.pack()
    exitButton.pack()
    root.mainloop()




def Mes():
    global driverModel
    global driverStatus

    #KILL TASK
    def killMes():
        os.system("taskkill /im MES(MEXICO).exe -f")
    #OPEN MES
    def launchMes():
        os.startfile("C:/Users/Pruebas/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/MES(MEXICO)/MES(MEXICO).appref-ms")
    global firstTimeFlagMes

    #VERIFY IF FIRST TIME OPENING PROGRAM
    if firstTimeFlagMes is True:
        killMes()
        launchMes()

        firstTimeFlagMes = False

        #Go to textBox
        findAndClick(mesLoginBtn, 5, 0.95, False)
        findAndClick(mesSideBtn, 5, 0.95, False)
        findAndClick(mesProcessScanSearch, 5, 0.95, False)

    # CLICK TEXTBOX
    findAndClick(mesTextBox, 5, 0.95, True)

    #ENTER SERIAL AND HIT ENTER
    pag.write(scannedSerialNumber.get().upper())
    global previousScann
    previousScann = scannedSerialNumber.get().upper()
    pag.press("tab")
    pag.press("enter")

    #NOT FOUND WINDOW SHOWS UP THEN DRIVER FAILS
    if imageFound(mesNotFoundAlert, 0.7, 0.93) or imageFound(mesNotFoundAlert2, 0.7, 0.93):
        driverStatus = 3
        driverModel = "Unknown"
        pyperclip.copy("3")
        print("Serial no encontrado")
        pag.press("enter")
    # elif imageFound(mesNotFoundAlert2, 0.8, 0.93):
    #         driverStatus = 3
    #         driverModel = "Unknown"
    #         print("Driver Model:" +driverModel)
    #         pyperclip.copy("3")
    #         print("Serial no encontrado")
    #         pag.press("enter")

    else:
    #GO TO RESULT OF SCANNED S/N
        global firstTimeFlagMes2
        if firstTimeFlagMes2 is True:
            pag.press("tab",5)
        else:
            pag.press("tab",4)

        pag.keyDown("ctrlleft")
        pag.press("c")
        pag.keyUp("ctrlleft")
        driverModel = pyperclip.paste()
        pag.hotkey("ctrlleft","tab")
        pag.hotkey("ctrlleft", "tab")
        if firstTimeFlagMes2:
            pag.hotkey("ctrlleft", "right")
            pag.hotkey("left")
        firstTimeFlagMes2 = False

        # COPY MES RESULT
        pag.keyDown("ctrlleft")
        pag.press("c")
        pag.keyUp("ctrlleft")
        # RETURN TO FIRST ROW
        pag.hotkey("ctrlleft", "tab")

     ## COPY CLIPBOARD TO A VARIABLE
    result = pyperclip.paste()
    if result == "1":
     driverStatus = 1
    elif result == "0":
     driverStatus = 2
    else:
        driverStatus = 3
    print(driverStatus)
    pyperclip.copy("")
    print("Model is "+driverModel+" serial status is "+str(driverStatus))
    mainWindow()


mainWindow()







