#Эти библиотеки позволяют работать с графикой.
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem
from PyQt5 import  uic
from PyQt5.QtWidgets import QFileDialog


#Определяем имя и путь до файлас формой окна.
ui=uic.loadUiType("interface/Annual_volume_planning.ui")[0]

#Этот класс определяет параметры окна и взаимодействие с ним.
class Annual_volume_planning(QtWidgets.QMainWindow, ui):
    def __init__(self):
        #Инициализируем окно
        super().__init__()
        self.setupUi(self)
        self.tableWidget.insertRow(0)
        print(self.tableWidget.columnCount())
        print(self.tableWidget.rowCount())
        print(self.tableWidget.item(0, 1))
        self.tableWidget.setItem(0, 0, QTableWidgetItem("Text in column 1"))
