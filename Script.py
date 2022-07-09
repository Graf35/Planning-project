#Импортируем модуль math. Он необходим для сложных вычислений.
import time
from datetime import date
#Импортируем модуль logging. Он необходим для возможности логирования.
import logging
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QTableWidgetItem, QTableWidget, QTableWidgetItem
import threading
import shutil

def loading_assets(parent, filename):
    logging.info("Запущена проверка загрузки активов")
    str_index=0
    assets = pd.read_excel(filename)
    assets["Номер родительского актива"]= np.nan
    assets["Описание"]= np.nan
    EAM = pd.read_csv('Database\EAM607.csv', ";", encoding='windows-1251')
    logging.info("Открыт ЕАМ607")
    for i in range((EAM.shape[0])):
        if EAM["Номер актива"][i] in assets["Номер актива"].tolist():
            assets["Номер родительского актива"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])] = EAM["Номер родительского актива"][i]
            assets["Описание"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])] = EAM["Описание актива"][i]
            # Выводим пользователю сообщение
            parent.tableWidget.insertRow(str_index)
            parent.tableWidget.setItem(str_index, 0, QTableWidgetItem(str(assets["Организация "][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])])))
            parent.tableWidget.setItem(str_index, 1, QTableWidgetItem(str(assets["Номер актива"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])])))
            parent.tableWidget.setItem(str_index, 2, QTableWidgetItem(str(assets["Номер родительского актива"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])])))
            parent.tableWidget.setItem(str_index, 3,QTableWidgetItem(str(assets["Описание"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])])))
            str_index +=1
    logging.info("Сверка с базой закончена")
    assets.to_excel("Database/assets.xlsx")
    logging.info("Создана база годового объёма ремонта")

def date_for_zvr(parent,cfo, project, task, department, basis_operation):
    tasks={'1':"Услуги ТМЦ сторонних", '3':"Услуги механик", '4':"Сервис ЧЭК"}
    departments={"Внешний":"Услуги подрядной организации", "Механик":"ООО Механик"}
    logging.info("Заполняю ЗВР")
    logging.info("Синхронизируюсь с базой")
    assets = pd.read_excel('Database/assets.xlsx')
    if cfo == '':
        logging.error("Ошибка заполнения раздела ЦФО")
        assets["ЦФО"] = np.nan
    else:
        assets["ЦФО"] = cfo
    if project == '':
        logging.error("Ошибка заполнения раздела Проект")
        assets["Проект"] = np.nan
    else:
        assets["Проект"] = project
    if department == '':
        logging.error("Ошибка заполнения раздела отдел")
        assets["Отдел"] = np.nan
    else:
        assets["Отдел"] = department
    try:
        assets["Описание отдела"] = departments[str(department)]
    except:
        logging.error("Неизвестный отдел")
        assets["Описание отдела"]=np.nan
    if task == '':
        logging.error("Ошибка заполнения раздела Задание")
        assets["Задача"] = np.nan
    else:
        assets["Задача"] = task
    try:
        assets["Описание задачи"] = tasks[str(task)]
    except:
        logging.error("Неизвестная задача")
        assets["Описание задачи"] = np.nan
    if basis_operation == '':
        logging.error("Ошибка заполнения раздела основание операции")
        assets["Основание операции"] = np.nan
    else:
        assets["Основание операции"] =basis_operation
    for i in range(parent.tableWidget.rowCount()):
        parent.tableWidget.setItem(i, 4, QTableWidgetItem(str(assets["ЦФО"][i])))
        parent.tableWidget.setItem(i, 5, QTableWidgetItem(str(assets["Проект"][i])))
        parent.tableWidget.setItem(i, 6, QTableWidgetItem(str(assets["Отдел"][i])))
        parent.tableWidget.setItem(i, 7, QTableWidgetItem(str(assets["Описание отдела"][i])))
        parent.tableWidget.setItem(i, 8, QTableWidgetItem(str(assets["Задача"][i])))
        parent.tableWidget.setItem(i, 9, QTableWidgetItem(str(assets["Описание задачи"][i])))
        parent.tableWidget.setItem(i, 10, QTableWidgetItem(str(assets["Основание операции"][i])))
    assets.to_excel("Database/assets.xlsx")
    logging.info("База годового объёма ремонта обновлена")


def creature_zvr(parent):
    month={"01":"Январь", "02":"Февраль", "03":"Март","04":"Апрель",
           "05":"Май", "06":"Июнь", "07":"Июль", "08":"Август",
           "09":"Сентябрь", "10":"Октябрь", "11":"Ноябрь", "12":"Декабрь"}
    operation=["ТО", "ТР", "КР"]
    logging.info("Создаю ЗВР")
    logging.info("Синхронизируюсь с базой")
    assets = pd.read_excel('Database/assets.xlsx')
    assets["Номер ЗВР"] = np.nan
    assets["Статус ЗВР"] = np.nan
    assets["Трудоёмкость"] = np.nan
    assets["Описание операции"] = np.nan
    assets["Месяц ремонта"] = np.nan
    assets["Операция с активом"]= np.nan
    assets["Дата создания"] = date.today()
    for i in range(parent.tableWidget.rowCount()):
        try:
            if str(parent.tableWidget.item(i, 13).text()) in month.keys():
                assets["Месяц ремонта"][i] = month[str(parent.tableWidget.item(i, 13).text())]
            else:
                logging.error("У актива " + str(assets["Номер актива"][i]) + " неверный формат месяца")
                continue
        except:
            logging.error("У актива " + str(assets["Номер актива"][i]) + " не указан месяц ремонта")
            continue
        try:
            assets["Операция с активом"][i] = str(parent.tableWidget.item(i, 11).text())
        except:
            logging.error("У актива " + str(assets["Номер актива"][i]) + " не указана операция с активом")
            assets["Операция с активом"][i]=" "
            continue
        if assets["Операция с активом"][i] in operation:
            if parent.tableWidget.item(i, 11).text()=="TO" or parent.tableWidget.item(i, 11).text()=="TО" or parent.tableWidget.item(i, 11).text()=="ТO" or parent.tableWidget.item(i, 11).text()=="ТО":
                Technical_card_ТО = pd.read_csv('Database/ТО.csv', ";", encoding='windows-1251')
                assets["Трудоёмкость"][i]=round(Technical_card_ТО["Трудоемкость"].drop(0).sum(), 2)
                assets["Описание операции"][i]="Техническое обслуживание"
            elif parent.tableWidget.item(i, 11).text()=="ТР" or parent.tableWidget.item(i, 11).text()=="TР" or parent.tableWidget.item(i, 11).text()=="ТP" or parent.tableWidget.item(i, 11).text()=="ТР":
                Technical_card_ТP = pd.read_csv('Database/ТР.csv', ";", encoding='windows-1251')
                assets["Трудоёмкость"][i] = round(Technical_card_ТP["Трудоемкость"].drop(0).sum(), 2)
                assets["Описание операции"][i] = "Текущий ремонт"
            elif parent.tableWidget.item(i, 11).text()=="КР" or parent.tableWidget.item(i, 11).text()=="KР" or parent.tableWidget.item(i, 11).text()=="КP" or parent.tableWidget.item(i, 11).text()=="КР":
                Technical_card_KP = pd.read_csv('Database/КР.csv', ";", encoding='windows-1251')
                assets["Трудоёмкость"][i] = round(Technical_card_KP["Трудоемкость"].drop(0).sum(), 2)
                assets["Описание операции"][i] = "капитальный ремонт"
        else:
            logging.error("У актива " + str(assets["Номер актива"][i]) + " неизвестный тип операции")
            continue
        config=filereader("config.config")
        zvrnomber=int(config['zvrnomber'])
        assets["Номер ЗВР"][i] = 'АВ-' + str(int(zvrnomber) + 1)
        configupdate(int(zvrnomber) + 1)
        parent.tableWidget.setItem(i, 12, QTableWidgetItem(str(assets["Описание операции"][i])))
        parent.tableWidget.setItem(i, 14, QTableWidgetItem(str(assets["Трудоёмкость"][i])))
        parent.tableWidget.setItem(i, 15, QTableWidgetItem(str(assets["Номер ЗВР"][i])))
        assets["Статус ЗВР"][i] = "проект"
        parent.tableWidget.setItem(i, 16, QTableWidgetItem(str(assets["Статус ЗВР"][i])))
        logging.info("Создан ЗВР " + str(assets["Номер ЗВР"][i]))
    logging.info("Создание ЗВР завершено")
    assets.to_excel("Database/assets.xlsx")
    logging.info("База годового объёма ремонта обновлена")


def monthly_plan(parent):
    month=parent.comboBox.currentText()

def report_EAM121(fname, month):
    pass

def report_EAM607(fname):
    shutil.copy('Database/EAM607.csv', str(fname))

def filereader(fail_name):
    sting=[]
    file=open(fail_name, 'r')
    for line in file:
        sting.append(line.strip().split(";"))
    file.close()
    filetext={}
    for i in range(len(sting)):
        filetext[sting[i][0]]=(sting[i][1])
    return (filetext)

def configupdate(zvr):
    lines = ["zvrnomber;" + str(zvr) + "\n"]
    with open("config.config", "w") as file:
        for line in lines:
            if line != "":
                file.write(line)


