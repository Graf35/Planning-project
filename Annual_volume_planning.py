#Эти библиотеки позволяют работать с графикой.
from PyQt5 import QtWidgets
from PyQt5 import  uic
from PyQt5.QtWidgets import QFileDialog
import Script
import shutil
#Этот модуль позволяет использовать многопоточность
import threading
import logging

#Определяем имя и путь до файлас формой окна.
ui=uic.loadUiType("interface/Annual_volume_planning.ui")[0]

#Этот класс определяет параметры окна и взаимодействие с ним.
class Annual_volume_planning(QtWidgets.QMainWindow, ui):
    def __init__(self):
        #Инициализируем окно
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.btnClicked1)
        self.pushButton_2.clicked.connect(self.btnClicked2)
        self.pushButton_3.clicked.connect(self.btnClicked3)
        self.pushButton_4.clicked.connect(self.btnClicked4)
        logging.info("Окно годового планирования объёмов создано")

    def btnClicked1(self):
        # Определяем путь до файла
        fname = QFileDialog.getSaveFileName(self, 'Save File', '/Шаблон загрузки активов.xlsx')[0]
        shutil.copy('Database/Asset_loading_template.xlsx', str(fname))
        logging.info("Выгружен шаблон загрузки активов в "+str(fname))


    def btnClicked2(self):
        # Определяем путь до файла
        fname = QFileDialog.getOpenFileName(self, 'Open file', '/Шаблон загрузки активов')[0]
        # Объявляем новый поток
        self.deman = threading.Thread(target=Script.loading_assets(self, fname))
        # Запускаем новый поток
        self.deman.start()

    def btnClicked3(self):
        # Объявляем новый поток
        self.deman11 = threading.Thread(target=Script.creature_zvr(self))
        # Запускаем новый поток
        self.deman11.start()

    def btnClicked4(self):
        cfo = self.comboBox_5.currentText()
        project = self.comboBox_4.currentText()
        task = self.comboBox_3.currentText()
        department = self.comboBox_2.currentText()
        basis_operation = self.comboBox.currentText()
        Script.date_for_zvr(self, cfo, project, task, department, basis_operation)