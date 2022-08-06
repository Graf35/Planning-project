#Эти библиотеки позволяют работать с графикой.
from PyQt5 import QtWidgets
from PyQt5 import  uic
from PyQt5.QtWidgets import QFileDialog
import Script
import shutil
import pandas as pd
from PyQt5.QtWidgets import QTableWidgetItem
#Этот модуль позволяет использовать многопоточность
import threading
import logging
import numpy as np
from datetime import date

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
        self.deman = threading.Thread(target=self.loading_assets(fname))
        # Запускаем новый поток
        self.deman.start()

    def btnClicked3(self):
        # Объявляем новый поток
        self.deman11 = threading.Thread(target=self.creature_zvr())
        # Запускаем новый поток
        self.deman11.start()

    def btnClicked4(self):
        cfo = self.comboBox_5.currentText()
        project = self.comboBox_4.currentText()
        task = self.comboBox_3.currentText()
        department = self.comboBox_2.currentText()
        basis_operation = self.comboBox.currentText()
        self.date_for_zvr(cfo, project, task, department, basis_operation)

    def loading_assets(self, filename):
        logging.info("Запущена проверка загрузки активов")
        str_index = 0
        assets = pd.read_excel(filename)
        assets["Номер родительского актива"] = np.nan
        assets["Описание"] = np.nan
        EAM = pd.read_csv('Database\EAM607.csv', ";", encoding='windows-1251')
        logging.info("Открыт ЕАМ607")
        for i in range((EAM.shape[0])):
            if EAM["Номер актива"][i] in assets["Номер актива"].tolist():
                assets["Номер родительского актива"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])] = \
                EAM["Номер родительского актива"][i]
                assets["Описание"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])] = EAM["Описание актива"][i]
                # Выводим пользователю сообщение
                self.tableWidget.insertRow(str_index)
                self.tableWidget.setItem(str_index, 0, QTableWidgetItem(str(assets["Организация "][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])])))
                self.tableWidget.setItem(str_index, 1, QTableWidgetItem(str(assets["Номер актива"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])])))
                self.tableWidget.setItem(str_index, 2, QTableWidgetItem(str(assets["Номер родительского актива"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])])))
                self.tableWidget.setItem(str_index, 3, QTableWidgetItem(str(assets["Описание"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])])))
                str_index += 1
        logging.info("Сверка с базой закончена")
        Script.base_update(assets)

    def date_for_zvr(self, cfo, project, task, department, basis_operation):
        tasks = {'1': "Услуги ТМЦ сторонних", '3': "Услуги механик", '4': "Сервис ЧЭК"}
        departments = {"Внешний": "Услуги подрядной организации", "Механик": "ООО Механик"}
        logging.info("Заполняю ЗВР")
        logging.info("Синхронизируюсь с базой")
        assets = pd.read_excel('Database/assets.xlsx')
        if cfo == ' ':
            logging.error("Ошибка заполнения раздела ЦФО")
            assets["ЦФО"] = np.nan
        else:
            assets["ЦФО"] = cfo
        if project == ' ':
            logging.error("Ошибка заполнения раздела Проект")
            assets["Проект"] = np.nan
        else:
            assets["Проект"] = project
        if department == ' ':
            logging.error("Ошибка заполнения раздела отдел")
            assets["Отдел"] = np.nan
        else:
            assets["Отдел"] = department
        try:
            assets["Описание отдела"] = departments[str(department)]
        except:
            logging.error("Неизвестный отдел")
            assets["Описание отдела"] = np.nan
        if task == ' ':
            logging.error("Ошибка заполнения раздела Задание")
            assets["Задача"] = np.nan
        else:
            assets["Задача"] = task
        try:
            assets["Описание задачи"] = tasks[str(task)]
        except:
            logging.error("Неизвестная задача")
            assets["Описание задачи"] = np.nan
        if basis_operation == ' ':
            logging.error("Ошибка заполнения раздела основание операции")
            assets["Основание операции"] = np.nan
        else:
            assets["Основание операции"] = basis_operation
        for i in range(self.tableWidget.rowCount()):
            self.tableWidget.setItem(i, 4, QTableWidgetItem(str(assets["ЦФО"][i])))
            self.tableWidget.setItem(i, 5, QTableWidgetItem(str(assets["Проект"][i])))
            self.tableWidget.setItem(i, 6, QTableWidgetItem(str(assets["Отдел"][i])))
            self.tableWidget.setItem(i, 7, QTableWidgetItem(str(assets["Описание отдела"][i])))
            self.tableWidget.setItem(i, 8, QTableWidgetItem(str(assets["Задача"][i])))
            self.tableWidget.setItem(i, 9, QTableWidgetItem(str(assets["Описание задачи"][i])))
            self.tableWidget.setItem(i, 10, QTableWidgetItem(str(assets["Основание операции"][i])))
        Script.base_update(assets)

    def creature_zvr(self):
        month = {"01": "Январь", "02": "Февраль", "03": "Март", "04": "Апрель",
                 "05": "Май", "06": "Июнь", "07": "Июль", "08": "Август",
                 "09": "Сентябрь", "10": "Октябрь", "11": "Ноябрь", "12": "Декабрь"}
        operation = ["ТО", "ТР", "КР"]
        logging.info("Создаю ЗВР")
        logging.info("Синхронизируюсь с базой")
        assets = pd.read_excel('Database/assets.xlsx')
        assets["Номер ЗВР"] = np.nan
        assets["Статус ЗВР"] = np.nan
        assets["Трудоёмкость"] = np.nan
        assets["Описание операции"] = np.nan
        assets["Месяц ремонта"] = np.nan
        assets["Операция с активом"] = np.nan
        assets["Дата создания"] = date.today()
        assets["Дата начала"] = np.nan
        assets["Дата окончания"] = np.nan
        assets["Дата принятия"] = np.nan
        for i in range(self.tableWidget.rowCount()):
            try:
                if str(self.tableWidget.item(i, 13).text()) in month.keys():
                    assets["Месяц ремонта"][i] = month[str(self.tableWidget.item(i, 13).text())]
                else:
                    logging.error("У актива " + str(assets["Номер актива"][i]) + " неверный формат месяца")
                    continue
            except:
                logging.error("У актива " + str(assets["Номер актива"][i]) + " не указан месяц ремонта")
                continue
            try:
                assets["Операция с активом"][i] = str(self.tableWidget.item(i, 11).text())
            except:
                logging.error("У актива " + str(assets["Номер актива"][i]) + " не указана операция с активом")
                assets["Операция с активом"][i] = " "
                continue
            if assets["Операция с активом"][i] in operation:
                if self.tableWidget.item(i, 11).text() == "TO" or self.tableWidget.item(i,
                                                                                            11).text() == "TО" or self.tableWidget.item(
                        i, 11).text() == "ТO" or self.tableWidget.item(i, 11).text() == "ТО":
                    Technical_card_ТО = pd.read_csv('Database/ТО.csv', ";", encoding='windows-1251')
                    assets["Трудоёмкость"][i] = round(Technical_card_ТО["Трудоемкость"].drop(0).sum(), 2)
                    assets["Описание операции"][i] = "Техническое обслуживание"
                elif self.tableWidget.item(i, 11).text() == "ТР" or self.tableWidget.item(i,
                                                                                              11).text() == "TР" or self.tableWidget.item(
                        i, 11).text() == "ТP" or self.tableWidget.item(i, 11).text() == "ТР":
                    Technical_card_ТP = pd.read_csv('Database/ТР.csv', ";", encoding='windows-1251')
                    assets["Трудоёмкость"][i] = round(Technical_card_ТP["Трудоемкость"].drop(0).sum(), 2)
                    assets["Описание операции"][i] = "Текущий ремонт"
                elif self.tableWidget.item(i, 11).text() == "КР" or self.tableWidget.item(i,
                                                                                              11).text() == "KР" or self.tableWidget.item(
                        i, 11).text() == "КP" or self.tableWidget.item(i, 11).text() == "КР":
                    Technical_card_KP = pd.read_csv('Database/КР.csv', ";", encoding='windows-1251')
                    assets["Трудоёмкость"][i] = round(Technical_card_KP["Трудоемкость"].drop(0).sum(), 2)
                    assets["Описание операции"][i] = "капитальный ремонт"
            else:
                logging.error("У актива " + str(assets["Номер актива"][i]) + " неизвестный тип операции")
                continue
            config = Script.filereader("config.config")
            zvrnomber = int(config['zvrnomber'])
            assets["Номер ЗВР"][i] = 'АВ-' + str(int(zvrnomber) + 1)
            Script.configupdate(int(zvrnomber) + 1)
            self.tableWidget.setItem(i, 12, QTableWidgetItem(str(assets["Описание операции"][i])))
            self.tableWidget.setItem(i, 14, QTableWidgetItem(str(assets["Трудоёмкость"][i])))
            self.tableWidget.setItem(i, 15, QTableWidgetItem(str(assets["Номер ЗВР"][i])))
            assets["Статус ЗВР"][i] = "проект"
            self.tableWidget.setItem(i, 16, QTableWidgetItem(str(assets["Статус ЗВР"][i])))
            logging.info("Создан ЗВР " + str(assets["Номер ЗВР"][i]))
        logging.info("Создание ЗВР завершено")
        Script.base_update(assets)

