from datetime import datetime
from os import system, path, mkdir
import win32com.client
from PIL import ImageGrab

# params
appName = "Excel.Application"
fileName = path.abspath('sample.xlsm')

sheetName = "Receipt"
imageFolder = "images"
imageFormat = "jpeg"

printRange = "A1:K24"
totalCell = "P6"
calcCell = "O5"
roomNameCell = "N5"

increaseStep = 2

# define


def KillProcess(appName):
    print("Trying  to close app ...")
    try:
        system('taskkill /F /IM EXCEL.exe')
    except:
        print("Can't kill running instance, please manual close it.")


def SaveToImage(workSheet, imageExtension, calcValue):
    workSheet.Range(calcCell).Value = calcValue
    workSheet.Range(printRange).Copy()

    roomName = workSheet.Range(roomNameCell).Value

    # create dir
    imageFolderName = imageFolder + "_" + datetime.today().strftime('%d-%m-%Y')
    if not path.exists(imageFolderName):
        mkdir(imageFolderName)

    # saving
    img = ImageGrab.grabclipboard()
    img.save(imageFolderName + "/" + sheetName +
             "_" + roomName + "." + imageExtension)


def Export(appRef, imageExtension):
    workBook = appRef.Workbooks.Open(Filename=fileName)
    try:
        # appRef.DisplayAlerts = False
        workSheet = workBook.Worksheets(sheetName)

        maxValue = int(workSheet.Range(totalCell).Value) + increaseStep
        for i in range(1, maxValue, increaseStep):
            SaveToImage(workSheet, imageExtension, i)

        workBook.Close(SaveChanges=False, Filename=fileName)
        appRef.Application.Quit()
        print("DONE.")

    except Exception as ex:
        print("Has error when export!.")
        print(ex)
        # KillProcess(appRef.Name)
        workBook.Close(False)
        appRef.Application.Quit()


# execute
xlsApp = win32com.client.gencache.EnsureDispatch(appName)
Export(xlsApp, imageFormat)
