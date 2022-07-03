#Эти библиотеки позволяют работать с графикой.
from PyQt5 import QtWidgets
from PyQt5 import  uic
from PyQt5.QtWidgets import QFileDialog
#import time
#from openpyxl import load_workbook

#Определяем имя и путь до файлас формой окна.
ui=uic.loadUiType("Annual_volume_planning.ui")[0]

#Этот класс определяет параметры окна и взаимодействие с ним.
class MaimWindow(QtWidgets.QMainWindow, ui):
    def __init__(self):
        #Инициализируем окно
        super().__init__()
        self.setupUi(self)
