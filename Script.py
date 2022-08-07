import time
import logging
import os
import shutil
from docx import Document
from docx.shared import Cm

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

def configupdate(zvr=None, update_time=None):
    if zvr==None:
        config = filereader("config.config")
        zvr=int(config['zvrnomber'])
    if update_time==None:
        config = filereader("config.config")
        update_time = config['update_time']
    lines = ["zvrnomber;" + str(zvr) + "\n"+"update_time;" + str(update_time) + "\n" ]
    with open("config.config", "w") as file:
        for line in lines:
            if line != "":
                file.write(line)

def base_update(base):
    if os.path.exists("Database/assets.xlsx"):
        pass
    else:
        shutil.copyfile('template/database/assets.xlsx', 'Database/assets.xlsx')
    try:
        base.to_excel("Database/assets.xlsx")
    except:
        logging.error("Открыт файл базы данных")
    update = time.ctime(os.path.getmtime("Database/assets.xlsx"))
    configupdate(update_time=update)
    logging.info("Создана база годового объёма ремонта")

def signature(fname, temp_dir):
    document = Document(fname)
    document.add_picture(temp_dir+'/template/shtamp.jpg', width=Cm(18.0), height=Cm(5))
    document.save(fname)



