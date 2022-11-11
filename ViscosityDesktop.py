from PyQt5 import uic
import xlsxwriter
from PyQt5.QtWidgets import QApplication, QMainWindow
#from PyQt5.QtCore import *
#from PyQt5.QtCore import QSize, Qt
#from PyQt5.QtWidgets import *

import os


class MainWindow(QMainWindow):
    

    listmeasurements = []
    listTime = []
    def __init__(self):
        super().__init__()

        uic.loadUi('pages.ui', self)
        self.CalculationButton.clicked.connect(self.Calculation)
        self.CalculationKButton.clicked.connect(self.CalculationK)
        self.ResetButton.clicked.connect(self.resetButton)
        self.CreateExcel.clicked.connect(self.createExcel)
        
    def Calculation (self):
        if self.inTime.text():
            inTime = float(self.inTime.text())
        else:
             inTime = 0

        if self.inDensityOil.text():
            inDensityOil = float(self.inDensityOil.text())
        else:
            inDensityOil = 0

        if self.inDensityBall.text():
            inDensityBall = float(self.inDensityBall.text())
        else:
            inDensityBall = 0

        if self.inK15.text():
            inK15 = float(self.inK15.text())
        else:
            inK15 = 0

        if self.inK30.text():
            inK30 = float(self.inK30.text())
        else:
            inK30 = 0

        if self.inK45.text():
            inK45 = float(self.inK45.text())
        else:
            inK45 = 0

        if  self.inK60.text():     
            inK60 = float(self.inK60.text())
        else:
            inK60 = 0

        radioButton_15 = self.radioButton_15.isChecked()
        radioButton_30 = self.radioButton_30.isChecked()
        radioButton_45 = self.radioButton_45.isChecked()
        radioButton_60 = self.radioButton_60.isChecked()
        Value = 0
        if radioButton_15 and inDensityBall>0 and inDensityOil >0 and inTime>0 and inK15>0:
            Value = ((inDensityBall - inDensityOil) * inTime * inK15)
        if radioButton_30 and inDensityBall>0 and inDensityOil >0 and inTime>0 and inK30>0:
            Value = ((inDensityBall - inDensityOil) * inTime * inK30)
        if radioButton_45 and inDensityBall>0 and inDensityOil >0 and inTime>0 and inK45>0:
            Value = ((inDensityBall - inDensityOil) * inTime * inK45)
        if radioButton_60 and inDensityBall>0 and inDensityOil >0 and inTime>0 and inK60>0:
            Value = ((inDensityBall - inDensityOil) * inTime * inK60)

        self.resViskosity.setText(str(round(Value,2)))

        self.listmeasurements.append(Value)
        self.listTime.append(inTime)
        summMeasurements  = 0

        for measurements in self.listmeasurements:
            summMeasurements = summMeasurements + measurements
        count = len(self.listmeasurements)
        if count  > 0:
            meanValue =  summMeasurements/ count 
            self.resMeanViskosity.setText(str(round(meanValue,2)))
            self.Count.setText(str(count))

    def CalculationK (self):
        inTime = float(self.inTime.text())
        inDensityOil = float(self.inDensityOil.text())
        inDensityBall = float(self.inDensityBall.text())
        inViscosityOilStandart = float(self.inViscosityOilStandart.text())
        self.resK.setText(str(round(inViscosityOilStandart/((inDensityBall - inDensityOil) * inTime),10) ))

    def resetButton (self):
        self.listmeasurements.clear()
        self.resMeanViskosity.setText(str(0))
        self.Count.setText(str(0))
        
    def createExcel (self):
        workbook = xlsxwriter.Workbook('Example.xlsx')
        worksheet = workbook.add_worksheet() 

        row = 1
        column = 1
        column2 = 2
        count = 0 
        id = 1
        content = ["ankit", "rahul", "priya", "harshita",
                    "sumit", "neeraj", "shivam"]

        worksheet.write(0, 0, "№")
        worksheet.write(0, 1, "время")
        worksheet.write(0, 2, "вязкость")

        for item in self.listTime :
            worksheet.write(row, 0, id)
            worksheet.write(row, column, item)
            worksheet.write(row, column2, self.listmeasurements[count])
            count += 1
            row += 1
            id += 1 
     
        workbook.close()
        
        os.system('start excel.exe Example.xlsx')


if __name__ == '__main__':
    
    app = QApplication([])
    window = MainWindow()
    window.show()
    #window.showMaximized()
    app.exec_()