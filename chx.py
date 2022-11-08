from ast import Break
from optparse import Option
from select import select
from xml.etree.ElementPath import xpath_tokenizer
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
import random
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as W
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support import expected_conditions as E
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
import pyautogui
import pandas as pd
import shutil
import os
import csv

data = pd.read_excel(r'Copia de Etiqueta_LI_CHX.xlsx')

driver=webdriver.Firefox()
rows = []
for i in range(len(data)):
    ocs = data['LABEL'][i]
    sleep(2)
    try:
        driver.get('https://centrodeayuda.chilexpress.cl/seguimiento/'+str(ocs))
        sleep(3)
        wait_time_out = 15
        wait_variable = W(driver, wait_time_out)

        ot = driver.find_element("xpath", '/html/body/app-root/app-seguimiento/main/div[2]/div/div/div[1]/div/b').text
        fecha = driver.find_element("xpath", '/html/body/app-root/app-seguimiento/main/div[2]/div/div/div[2]/div[1]/div/div[2]/div/table/tbody/tr[1]/td[1]').text
        hora = driver.find_element("xpath", '/html/body/app-root/app-seguimiento/main/div[2]/div/div/div[2]/div[1]/div/div[2]/div/table/tbody/tr[1]/td[2]').text
        estado = driver.find_element("xpath", '/html/body/app-root/app-seguimiento/main/div[2]/div/div/div[2]/div[1]/div/div[2]/div/table/tbody/tr[1]/td[3]').text

    except:
        pass

    #print(ot)
    #print(fecha)
    #print(hora)
    #print(estado)

    
    rows.append([ot,fecha,hora,estado])
    print(rows)

cabecera = ['ot','fecha','hora','estado'] 

with open('SELLERS1.csv','w') as f:
      
    write = csv.writer(f)
    write.writerow(cabecera)
    write.writerows(rows)


