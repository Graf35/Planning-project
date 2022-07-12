#Импортируем модуль math. Он необходим для сложных вычислений.
import time
from datetime import date
#Импортируем модуль logging. Он необходим для возможности логирования.
import logging
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QTableWidgetItem, QTableWidget, QTableWidgetItem
import os
import shutil
from Office import World



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
    assets["Дата начала"] = np.nan
    assets["Дата окончания"] = np.nan
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
    str_index=0
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
            parent.tableWidget.insertRow(str_index)
            parent.tableWidget.setItem(str_index, 0, QTableWidgetItem(assets["Номер ЗВР"][i]))
            parent.tableWidget.setItem(str_index, 1, QTableWidgetItem(assets["Номер родительского актива"][i]))
            parent.tableWidget.setItem(str_index, 2, QTableWidgetItem(assets["Номер актива"][i]))
            parent.tableWidget.setItem(str_index, 3, QTableWidgetItem(assets["Операция с активом"][i]))
            parent.tableWidget.setItem(str_index, 4, QTableWidgetItem(assets["Статус ЗВР"][i]))
            parent.tableWidget.setItem(str_index, 5, QTableWidgetItem("01-" + str(month_nom[assets["Месяц ремонта"][i]]) + "-2023"))
            parent.tableWidget.setItem(str_index, 6, QTableWidgetItem(str(month_daes[assets["Месяц ремонта"][i]]) + "-" + str(month_nom[assets["Месяц ремонта"][i]]) + "-2023"))


def release_zvr(parent):
    month = parent.comboBox.currentText()
    str_index = 0
    logging.info("Синхронизируюсь с базой")
    assets = pd.read_excel('Database/assets.xlsx')
    for i in range((assets.shape[0])):
        if assets["Месяц ремонта"][i] == month:
            assets["Дата начала"][i] = date.today()
            assets["Статус ЗВР"][i] = "Выпущено"
            parent.tableWidget.setItem(str_index, 4, QTableWidgetItem(assets["Статус ЗВР"][i]))
            parent.tableWidget.setItem(str_index, 7, QTableWidgetItem(str(assets["Дата начала"][i])))
            str_index += 1
            logging.info("ЗВР "+str(assets["Номер ЗВР"][i])+" изменил статус на "+ str(assets["Статус ЗВР"][i]))
    assets.to_excel("Database/assets.xlsx")
    logging.info("База годового объёма ремонта обновлена")



def report_EAM121(fname, month):
    # Активируем возможность работы с файлами .doxc
    Word = World()
    temp_dir=os.getcwd()
    logging.info("Синхронизируюсь с базой")
    assets = pd.read_excel('Database/assets.xlsx')
    month_daes={"Январь": 31, "Февраль": 28, "Март": 31, "Апрель":30,
           "Май":31, "Июнь":30, "Июль":31, "Август":31,
           "Сентябрь":30, "Октябрь":31, "Ноябрь":30, "Декабрь":31}
    month_nom={"Январь": "01", "Февраль": "02", "Март": "03", "Апрель":"04",
           "Май":"05", "Июнь":"06", "Июль":"07", "Август":"08",
           "Сентябрь":"09", "Октябрь":"10", "Ноябрь":"11", "Декабрь":"12"}
    os.chdir(str(fname))
    try:
        os.mkdir(str(fname)+"/EAM121")
    except:
        pass
    if month=="Год":
        logging.info("Создаю ЕАМ121 на год в директорию "+str(fname))
        os.chdir(str(fname) + "/EAM121")
        for i in range(len(list(month_daes.keys()))):
            try:
                os.mkdir(str(fname) + "/EAM121/"+str(list(month_daes.keys())[i]))
            except:
                continue
        for i in range((assets.shape[0])):
            os.chdir(temp_dir)
            if assets["Операция с активом"][i]=="TO" or assets["Операция с активом"][i]=="TО" or assets["Операция с активом"][i]=="ТO" or assets["Операция с активом"][i]=="ТО":
                shutil.copy('template/EAM121/EAM121ТО.docx', str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(assets["Номер ЗВР"][i]) + ".docx")
            elif assets["Операция с активом"][i]=="TP" or assets["Операция с активом"][i]=="TР" or assets["Операция с активом"][i]=="ТP" or assets["Операция с активом"][i]=="ТР":
                shutil.copy('template/EAM121/EAM121ТР.docx', str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(assets["Номер ЗВР"][i]) + ".docx")
            elif assets["Операция с активом"][i]=="КР" or assets["Операция с активом"][i]=="КP" or assets["Операция с активом"][i]=="KР" or assets["Операция с активом"][i]=="KP":
                shutil.copy('template/EAM121/EAM121КР.docx', str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(assets["Номер ЗВР"][i]) + ".docx")
            data_report = {"date": str(assets["Дата создания"][i]), "zvr": str(assets["Номер ЗВР"][i]),
                           "division": str(assets["Организация "][i]),
                           "asset": str(assets["Номер актива"][i]),
                           "parent_asset": str(assets["Номер родительского актива"][i]),
                           "operation": str(assets["Операция с активом"][i]),
                           "laboriousness": str(assets["Трудоёмкость"][i]),
                           "basis_of_the_operation": str(assets["Основание операции"][i]),
                           "project": str(assets["Проект"][i]),
                           "planned_start_date": "01-" + str(month_nom[assets["Месяц ремонта"][i]]) + "-2023",
                           "planned_end_date": str(month_daes[assets["Месяц ремонта"][i]]) + "-" + str(month_nom[assets["Месяц ремонта"][i]]) + "-2023",
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
            Word.save(str(assets["Номер ЗВР"][i]) + ".docx")
            os.chdir(temp_dir)
            logging.info("Отчёт "+str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i])+"/"+str(assets["Номер ЗВР"][i]) + ".docx создан")
    else:
        logging.info("Создаю ЕАМ121 на "+str(month)+ " в директорию" + str(fname))
        try:
            os.mkdir(str(fname) + "/EAM121/" + str(month))
        except:
            pass
        for i in range((assets.shape[0])):
            if assets["Месяц ремонта"][i]==month:
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
                data_report={"date": str(assets["Дата создания"][i]), "zvr": str(assets["Номер ЗВР"][i]),
               "division": str(assets["Организация "][i]),
               "asset": str(assets["Номер актива"][i]), "parent_asset": str(assets["Номер родительского актива"][i]),
               "operation": str(assets["Операция с активом"][i]), "laboriousness": str(assets["Трудоёмкость"][i]),
               "basis_of_the_operation": str(assets["Основание операции"][i]), "project": str(assets["Проект"][i]),
               "planned_start_date": "01-" + str(month_nom[month]) + "-2023",
               "planned_end_date": str(month_daes[month]) + "-" + month_nom[month] + "-2023",
               "description": str(assets["Описание"][i]),
               "description_of_the_operation": str(assets["Описание операции"][i]),
               "department_description": str(assets["Описание отдела"][i]), "department": str(assets["Отдел"][i]),
               "task": str(assets["Задача"][i]), "task_description": str(assets["Описание задачи"][i]),
               "statust": str(assets["Статус ЗВР"][i]), "tsfo": str(assets["ЦФО"][i]),
               "actual_start_date": str(assets["Дата начала"][i]),
               "actual_end_date": str(assets["Дата окончания"][i]),
               "print_date": str(date.today()), "creator": str(os.environ.get("USERNAME"))}
                # Активируем файл
                os.chdir(str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]))
                Word.record(data_report)
                Word.save(str(assets["Номер ЗВР"][i]) + ".docx")
                os.chdir(temp_dir)
                logging.info("Отчёт " + str(fname) + "/EAM121/" + str(assets["Месяц ремонта"][i]) + "/" + str(
                    assets["Номер ЗВР"][i]) + ".docx создан")

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


