from datetime import datetime
from os import mkdir, path
import string
import sys
from typing import List
from PIL import ImageGrab
import win32com.client
from PySide6 import QtCore
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QFormLayout,
    QPushButton,
    QLineEdit,
    QFileDialog,
    QMessageBox
)


class ExcellOptions:
    sheetName: string
    imageFolder: string = "images"
    imageFormat: string
    roomNameCell: string
    printRange: string
    totalCell: string
    increaseCell: string
    increaseStep: string
    fileName: string


class FormControl:
    def __init__(self, label, widget: QLineEdit):
        self.label = label
        self.widget = widget


class MainWidget(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Converter Excel")

        self.formControls = [
            FormControl(label="Sheet Name", widget=QLineEdit("Receipt")),
            FormControl(label="Print Range", widget=QLineEdit("A1:K24")),
            FormControl(label="Total Room Cell", widget=QLineEdit("P6")),
            FormControl(label="Room Name Cell", widget=QLineEdit("N5")),
            FormControl(label="Increase Cell", widget=QLineEdit("O5")),
            FormControl(label="Increase Step", widget=QLineEdit('2')),
            FormControl(label="Image Format", widget=QLineEdit('jpeg')),
        ]

        container = QVBoxLayout(self)

        formLayout = QFormLayout()
        for control in self.formControls:
            control.widget.setFixedWidth(300)
            formLayout.addRow(control.label, control.widget)

        self.filePathControl = QLineEdit()

        btnSubmit = QPushButton("Execute!")
        btnFileBrowser = QPushButton("File Browser")

        formLayout.setHorizontalSpacing(30)
        formLayout.addRow("File Path", self.filePathControl)
        formLayout.addWidget(btnFileBrowser)

        container.addLayout(formLayout)
        container.addWidget(btnSubmit)
        self.setLayout(container)

        btnSubmit.clicked.connect(self.submit)
        btnFileBrowser.clicked.connect(self.browserFile)

    @QtCore.Slot()
    def submit(self):
        options = self.getFormValue(self.formControls)
        options.fileName = self.filePathControl.text()
        self.Export(options)

    @QtCore.Slot()
    def browserFile(self):
        filepath = QFileDialog.getOpenFileName(
            self, "Open Excel file", "~", "Excel Files (*.xlsm *.xlsx)")[0]
        self.filePathControl.setText(filepath)

    def getFormValue(self, controls: List[FormControl]):
        options = ExcellOptions()
        options.sheetName = controls[0].widget.text()
        options.printRange = controls[1].widget.text()
        options.totalCell = controls[2].widget.text()
        options.roomNameCell = controls[3].widget.text()
        options.increaseCell = controls[4].widget.text()
        options.increaseStep = controls[5].widget.text()
        options.imageFormat = controls[6].widget.text()
        return options

    def SaveToImage(self, workSheet, increaseValue, imageFolderName,  options: ExcellOptions):
        workSheet.Range(options.increaseCell).Value = increaseValue
        workSheet.Range(options.printRange).Copy()

        roomName = workSheet.Range(options.roomNameCell).Value

        # saving
        img = ImageGrab.grabclipboard()
        img.save(imageFolderName + "/" + options.sheetName +
                 "_" + roomName + "." + options.imageFormat)

    def Export(self, options: ExcellOptions):
        appName = "Excel.Application"
        xlsApp = win32com.client.gencache.EnsureDispatch(appName)
        workBook = xlsApp.Workbooks.Open(Filename=options.fileName)
        try:
            xlsApp.DisplayAlerts = False
            workSheet = workBook.Worksheets(options.sheetName)
            increaseStep = int(options.increaseStep)

            maxValue = int(workSheet.Range(
                options.totalCell).Value) + increaseStep

            # create dir
            imageFolderName = options.imageFolder + "_" + \
                datetime.today().strftime('%d-%m-%Y')
            if not path.exists(imageFolderName):
                mkdir(imageFolderName)

            # save images
            for i in range(1, maxValue, increaseStep):
                self.SaveToImage(workSheet, i, imageFolderName, options)

            workBook.Close(False)
            xlsApp.Application.Quit()

            # show message
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Success")
            msg.setText("Export successfully!.")
            msg.exec()

            print("DONE.")

        except Exception as ex:
            workBook.Close(False)
            xlsApp.Application.Quit()

            # show message
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Error")
            msg.setText(str(ex))
            msg.exec()

            print("Has error when export!.")
            print(ex)


if __name__ == "__main__":
    app = QApplication([])

    widget = MainWidget()
    widget.resize(400, 400)
    widget.show()

    sys.exit(app.exec())
