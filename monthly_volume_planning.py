#Эти библиотеки позволяют работать с графикой.
from PyQt5 import QtWidgets
from PyQt5 import  uic
import Script
#Этот модуль позволяет использовать многопоточность
import threading
import logging
import pandas as pd
from datetime import date
from PyQt5.QtWidgets import QTableWidgetItem, QTableWidget, QTableWidgetItem




#Определяем имя и путь до файлас формой окна.
ui=uic.loadUiType("interface/Monthly_volume_planning.ui")[0]

#Этот класс определяет параметры окна и взаимодействие с ним.
class Monthly_volume_planning(QtWidgets.QMainWindow, ui):
    def __init__(self):
        #Инициализируем окно
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.btnClicked1)
        self.pushButton_2.clicked.connect(self.btnClicked2)
        logging.info("Окно месячного планирования объёмов создано")

    def btnClicked1(self):
        # Объявляем новый поток
        self.deman31 = threading.Thread(target=self.monthly_plan())
        # Запускаем новый поток
        self.deman31.start()

    def btnClicked2(self):
        self.deman32 = threading.Thread(target=self.release_zvr())
        # Запускаем новый поток
        self.deman32.start()

    def monthly_plan(self):
        month = self.comboBox.currentText()
        str_index = 0
        month_daes = {"Январь": 31, "Февраль": 28, "Март": 31, "Апрель": 30,
                      "Май": 31, "Июнь": 30, "Июль": 31, "Август": 31,
                      "Сентябрь": 30, "Октябрь": 31, "Ноябрь": 30, "Декабрь": 31}
        month_nom = {"Январь": "01", "Февраль": "02", "Март": "03", "Апрель": "04",
                     "Май": "05", "Июнь": "06", "Июль": "07", "Август": "08",
                     "Сентябрь": "09", "Октябрь": "10", "Ноябрь": "11", "Декабрь": "12"}
        logging.info("Синхронизируюсь с базой")
        assets = pd.read_excel('Database/assets.xlsx')
        for i in range((assets.shape[0])):
            if assets["Месяц ремонта"][i] == month:
                self.tableWidget.insertRow(str_index)
                self.tableWidget.setItem(str_index, 0, QTableWidgetItem(assets["Номер ЗВР"][i]))
                self.tableWidget.setItem(str_index, 1, QTableWidgetItem(assets["Номер родительского актива"][i]))
                self.tableWidget.setItem(str_index, 2, QTableWidgetItem(assets["Номер актива"][i]))
                self.tableWidget.setItem(str_index, 3, QTableWidgetItem(assets["Операция с активом"][i]))
                self.tableWidget.setItem(str_index, 4, QTableWidgetItem(assets["Статус ЗВР"][i]))
                self.tableWidget.setItem(str_index, 5, QTableWidgetItem(
                    "01-" + str(month_nom[assets["Месяц ремонта"][i]]) + "-2023"))
                self.tableWidget.setItem(str_index, 6, QTableWidgetItem(
                    str(month_daes[assets["Месяц ремонта"][i]]) + "-" + str(
                        month_nom[assets["Месяц ремонта"][i]]) + "-2023"))
                self.tableWidget.setItem(str_index, 7, QTableWidgetItem(str(assets["Дата начала"][i])))
                self.tableWidget.setItem(str_index, 8, QTableWidgetItem(str(assets["Дата начала"][i])))
                self.tableWidget.setItem(str_index, 9, QTableWidgetItem(str(assets["Дата окончания"][i])))
                self.tableWidget.setItem(str_index, 10, QTableWidgetItem(str(assets["Дата принятия"][i])))

    def release_zvr(self):
        month = self.comboBox.currentText()
        str_index = 0
        logging.info("Синхронизируюсь с базой")
        assets = pd.read_excel('Database/assets.xlsx')
        for i in range((assets.shape[0])):
            if assets["Месяц ремонта"][i] == month:
                assets["Дата начала"][i] = date.today()
                assets["Статус ЗВР"][i] = "Выпущено"
                self.tableWidget.setItem(str_index, 4, QTableWidgetItem(assets["Статус ЗВР"][i]))
                self.tableWidget.setItem(str_index, 7, QTableWidgetItem(str(assets["Дата начала"][i])))
                self.tableWidget.setItem(str_index, 8, QTableWidgetItem(str(assets["Дата начала"][i])))
                str_index += 1
                logging.info(
                    "ЗВР " + str(assets["Номер ЗВР"][i]) + " изменил статус на " + str(assets["Статус ЗВР"][i]))
        Script.base_update(assets)