import pyautogui as pag
import time

def img(name):
    return 'ss\\'+name

mesLoginBtn = img("mesLoginBtn.PNG")
mesNotFoundAlert = img("mesNotFoundAlert.PNG")
mesSideBtn = img("mesSideBtn.PNG")
mesTextBox = img("mesTextBox.PNG")
mesProcessScanSearch = img("mesProcessScanSearch.PNG")

def findAndClick(img, timeLimit, conf, doubleClick):
    cords = None
    timeLimit = (timeLimit/0.4)
    i = timeLimit
    while cords is None:
     if i > 0:
        i=i-1
        cords = pag.locateCenterOnScreen(img, confidence=conf)
        time.sleep(0.2)
     else:
         print('Out')
         pag.alert("Ocurrió un error, solicite apoyo al técnico de pruebas")
         exit()

    if doubleClick:
         pag.doubleClick(cords)
    else:
         pag.click(cords)


def findAndBool(img, timeLimit, conf):
    cords = None
    timeLimit = (timeLimit/0.4)
    i = timeLimit
    while cords is None:
     if i > 0:
        i=i-1
        cords = pag.locateCenterOnScreen(img, confidence=conf)
        time.sleep(0.2)
     else:
         return False

     return True


def imageFound(img, timeLimit, conf):
    cords = None
    timeLimit = (timeLimit/0.4)
    i = timeLimit
    while cords is None:
     if i > 0:
        i=i-1
        cords = pag.locateCenterOnScreen(img, confidence=conf)
        time.sleep(0.2)
     else:
         return False

    return True
