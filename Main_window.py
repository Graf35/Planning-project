#Эти библиотеки позволяют работать с графикой.
from PyQt5 import QtWidgets
from PyQt5 import  uic
from PyQt5.QtWidgets import QFileDialog
from time import ctime
from os import path
import Script
from Annual_volume_planning import Annual_volume_planning
from monthly_volume_planning import Monthly_volume_planning
from closed_form_of_gold_reserves import Closed_form_of_gold_reserves
import pandas as pd
import os
import shutil
from Office import World
from datetime import date
#Этот модуль позволяет использовать многопоточность
import threading
#Импортируем модуль logging. Он необходим для возможности логирования.
import logging
import log

#Определяем имя и путь до файлас формой окна.
ui=uic.loadUiType("interface/Main_window.ui")[0]

#Этот класс определяет параметры окна и взаимодействие с ним.
class MaimWindow(QtWidgets.QMainWindow, ui):
    def __init__(self):
        #Инициализируем окно
        super().__init__()
        self.setupUi(self)
        # Прописываем действие на нажатие кнопки
        self.pushButton.clicked.connect(self.btnClicked1)
        self.pushButton_2.clicked.connect(self.btnClicked2)
        self.pushButton_3.clicked.connect(self.btnClicked3)
        # Применяем настройки логирования
        logger = log.Deman_log()
        config = Script.filereader("config.config")
        if config['update_time']!="1":
            if ctime(path.getmtime("Database/assets.xlsx"))!= config['update_time']:
                self.closed_form_of_gold_reserves=Closed_form_of_gold_reserves()
                self.deman55 = threading.Thread(target=self.closed_form_of_gold_reserves.show())
                # Запускаем новый поток
                self.deman55.start()

    def btnClicked1(self):
        self.Annual_volume_planning_window = Annual_volume_planning()  # Создаём объект класса`
        # Объявляем новый поток
        self.deman1 = threading.Thread(target=self.Annual_volume_planning_window.show())
        # Запускаем новый поток
        self.deman1.start()

    def btnClicked2(self):
        self.monthly_volume_planning_window = Monthly_volume_planning()  # Создаём объект класса`
        # Объявляем новый поток
        self.deman1 = threading.Thread(target=self.monthly_volume_planning_window.show())
        # Запускаем новый поток
        self.deman1.start()

    def btnClicked3(self):
        try:
            report = self.comboBox.currentText()
        except:
            logging.error("Не выбран номер отчёта")
        try:
            month = self.comboBox_2.currentText()
        except:
            logging.error("Не выбран период отчёта")
        if report=="ЕАМ121":
            fname = QFileDialog.getExistingDirectory(self)
            self.report_EAM121(fname, month)
            logging.info("Выгружен EAM 121 в " + str(fname)+" за период "+str(month))
        elif report=="ЕАМ607":
            fname = QFileDialog.getSaveFileName(self, 'Save File', 'report607.csv')[0]
            self.report_EAM607(fname)
            logging.info("Выгружен EAM 607 в " + str(fname))

    def report_EAM121(self, fname, month):
        # Активируем возможность работы с файлами .doxc
        Word = World()
        temp_dir = os.getcwd()
        logging.info("Синхронизируюсь с базой")
        assets = pd.read_excel('Database/assets.xlsx')
        month_daes = {"Январь": 31, "Февраль": 28, "Март": 31, "Апрель": 30,
                      "Май": 31, "Июнь": 30, "Июль": 31, "Август": 31,
                      "Сентябрь": 30, "Октябрь": 31, "Ноябрь": 30, "Декабрь": 31}
        month_nom = {"Январь": "01", "Февраль": "02", "Март": "03", "Апрель": "04",
                     "Май": "05", "Июнь": "06", "Июль": "07", "Август": "08",
                     "Сентябрь": "09", "Октябрь": "10", "Ноябрь": "11", "Декабрь": "12"}
        os.chdir(str(fname))
        try:
            os.mkdir(str(fname) + "/EAM121")
        except:
            pass
        if month == "Год":
            logging.info("Создаю ЕАМ121 на год в директорию " + str(fname))
            os.chdir(str(fname) + "/EAM121")
            for i in range(len(list(month_daes.keys()))):
                try:
                    os.mkdir(str(fname) + "/EAM121/" + str(list(month_daes.keys())[i]))
                except:
                    continue
            for i in range((assets.shape[0])):
                os.chdir(temp_dir)
                if assets["Операция с активом"][i] == "TO" or assets["Операция с активом"][i] == "TО" or \
                        assets["Операция с активом"][i] == "ТO" or assets["Операция с активом"][i] == "ТО":
                    shutil.copy('template/EAM121/EAM121ТО.docx',
                                str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(
                                    assets["Номер ЗВР"][i]) + ".docx")
                elif assets["Операция с активом"][i] == "TP" or assets["Операция с активом"][i] == "TР" or \
                        assets["Операция с активом"][i] == "ТP" or assets["Операция с активом"][i] == "ТР":
                    shutil.copy('template/EAM121/EAM121ТР.docx',
                                str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(
                                    assets["Номер ЗВР"][i]) + ".docx")
                elif assets["Операция с активом"][i] == "КР" or assets["Операция с активом"][i] == "КP" or \
                        assets["Операция с активом"][i] == "KР" or assets["Операция с активом"][i] == "KP":
                    shutil.copy('template/EAM121/EAM121КР.docx',
                                str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(
                                    assets["Номер ЗВР"][i]) + ".docx")
                data_report = {"date": str(assets["Дата создания"][i]), "zvr": str(assets["Номер ЗВР"][i]),
                               "division": str(assets["Организация "][i]),
                               "asset": str(assets["Номер актива"][i]),
                               "parent_asset": str(assets["Номер родительского актива"][i]),
                               "operation": str(assets["Операция с активом"][i]),
                               "laboriousness": str(assets["Трудоёмкость"][i]),
                               "basis_of_the_operation": str(assets["Основание операции"][i]),
                               "project": str(assets["Проект"][i]),
                               "planned_start_date": "01-" + str(month_nom[assets["Месяц ремонта"][i]]) + "-2023",
                               "planned_end_date": str(month_daes[assets["Месяц ремонта"][i]]) + "-" + str(
                                   month_nom[assets["Месяц ремонта"][i]]) + "-2023",
                               "description": str(assets["Описание"][i]),
                               "description_of_the_operation": str(assets["Описание операции"][i]),
                               "department_description": str(assets["Описание отдела"][i]),
                               "department": str(assets["Отдел"][i]),
                               "task": str(assets["Задача"][i]), "task_description": str(assets["Описание задачи"][i]),
                               "statust": str(assets["Статус ЗВР"][i]), "tsfo": str(assets["ЦФО"][i]),
                               "actual_start_date": str(assets["Дата начала"][i]),
                               "actual_end_date": str(assets["Дата окончания"][i]),
                               "print_date": str(date.today()), "creator": str(os.environ.get("USERNAME"))}
                # Активируем файл
                os.chdir(str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]))
                Word.record(data_report)
                Word.save(str(assets["Номер ЗВР"][i]) + ".docx", temp_dir)
                if str(assets["Статус ЗВР"][i])=="Завершено":
                    Script.signature(str(assets["Номер ЗВР"][i]) + ".docx")
                os.chdir(temp_dir)
                logging.info("Отчёт " + str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(
                    assets["Номер ЗВР"][i]) + ".docx создан")
        else:
            logging.info("Создаю ЕАМ121 на " + str(month) + " в директорию" + str(fname))
            try:
                os.mkdir(str(fname) + "/EAM121/" + str(month))
            except:
                pass
            for i in range((assets.shape[0])):
                if assets["Месяц ремонта"][i] == month:
                    os.chdir(temp_dir)
                    if assets["Операция с активом"][i] == "TO" or assets["Операция с активом"][i] == "TО" or \
                            assets["Операция с активом"][i] == "ТO" or assets["Операция с активом"][i] == "ТО":
                        shutil.copy('template/EAM121/EAM121ТО.docx',
                                    str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(
                                        assets["Номер ЗВР"][i]) + ".docx")
                    elif assets["Операция с активом"][i] == "TP" or assets["Операция с активом"][i] == "TР" or \
                            assets["Операция с активом"][i] == "ТP" or assets["Операция с активом"][i] == "ТР":
                        shutil.copy('template/EAM121/EAM121ТР.docx',
                                    str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(
                                        assets["Номер ЗВР"][i]) + ".docx")
                    elif assets["Операция с активом"][i] == "КР" or assets["Операция с активом"][i] == "КP" or \
                            assets["Операция с активом"][i] == "KР" or assets["Операция с активом"][i] == "KP":
                        shutil.copy('template/EAM121/EAM121КР.docx',
                                    str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(
                                        assets["Номер ЗВР"][i]) + ".docx")
                    data_report = {"date": str(assets["Дата создания"][i]), "zvr": str(assets["Номер ЗВР"][i]),
                                   "division": str(assets["Организация "][i]),
                                   "asset": str(assets["Номер актива"][i]),
                                   "parent_asset": str(assets["Номер родительского актива"][i]),
                                   "operation": str(assets["Операция с активом"][i]),
                                   "laboriousness": str(assets["Трудоёмкость"][i]),
                                   "basis_of_the_operation": str(assets["Основание операции"][i]),
                                   "project": str(assets["Проект"][i]),
                                   "planned_start_date": "01-" + str(month_nom[month]) + "-2023",
                                   "planned_end_date": str(month_daes[month]) + "-" + month_nom[month] + "-2023",
                                   "description": str(assets["Описание"][i]),
                                   "description_of_the_operation": str(assets["Описание операции"][i]),
                                   "department_description": str(assets["Описание отдела"][i]),
                                   "department": str(assets["Отдел"][i]),
                                   "task": str(assets["Задача"][i]),
                                   "task_description": str(assets["Описание задачи"][i]),
                                   "statust": str(assets["Статус ЗВР"][i]), "tsfo": str(assets["ЦФО"][i]),
                                   "actual_start_date": str(assets["Дата начала"][i]),
                                   "actual_end_date": str(assets["Дата окончания"][i]),
                                   "print_date": str(date.today()), "creator": str(os.environ.get("USERNAME"))}
                    # Активируем файл
                    os.chdir(str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]))
                    Word.record(data_report)
                    Word.save(str(assets["Номер ЗВР"][i]) + ".docx")
                    if str(assets["Статус ЗВР"][i]) == "Завершено":
                        Script.signature(str(assets["Номер ЗВР"][i]) + ".docx", temp_dir)
                    os.chdir(temp_dir)
                    logging.info("Отчёт " + str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(
                        assets["Номер ЗВР"][i]) + ".docx создан")

    def report_EAM607(self, fname):
        shutil.copy('Database/EAM607.csv', str(fname))