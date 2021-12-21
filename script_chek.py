from tkinter import *
import tkinter as tk
import tkinter.messagebox as mb
from tkinter import Frame
from pyrus import client
import pyrus.models
from pyrus.models import responses
from selenium import webdriver
import time
from selenium.webdriver.chrome import options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
from pprint import pprint
import httplib2
import apiclient
from oauth2client.service_account import ServiceAccountCredentials


# Файл, полученный в Google Developer Console
CREDENTIALS_FILE = 'cred.json'
# ID Google Sheets документа (можно взять из его URL)
spreadsheet_id = 'your_id_sheets'

# Авторизуемся и получаем service — экземпляр доступа к API
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    ['https://www.googleapis.com/auth/spreadsheets',
     'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

pyrus_client = client.PyrusAPI(login='v.kolomeets@docsinbox.ru', security_key='your_key_pyrus')

auth_response = pyrus_client.auth()
if auth_response.success:
    pass

body = {
    'values' : [
        [" ", " "],  
    ]
}

url = "https://cerberus.vetrf.ru/cerberus/actualObject/pub/actualInfo/"
options = Options()
options.headless = True
while True:
    try: 
            values_guid = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range='A2:A2',
                majorDimension='COLUMNS'
            ).execute()
            res = values_guid['values'][0][0]

            values_id1 = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range='B2:B2',
                majorDimension='COLUMNS'
            ).execute()
            res_id1 = values_id1['values'][0][0]
            res_id = int(res_id1)

        
            options = Options()
            options.headless = True
            # driver = webdriver.Chrome()
            driver = webdriver.Chrome(chrome_options=options)
            
            driver.get(url=url)
            driver.find_element_by_tag_name('input').send_keys(res)
            driver.find_element_by_tag_name('button').click()
            time.sleep(0.5)

            if driver.find_elements_by_xpath('//span[@class="label label-success"][1]'):
                if True:
                    request = pyrus.models.requests.TaskCommentRequest(text="Подтверждена", action="")
                    task = pyrus_client.comment_task(res_id, request).task
                    res = print('Подтверждена')
                    resp = service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range="guid!A2",
                        valueInputOption="RAW",
                        body=body).execute()
                    

            elif driver.find_elements_by_xpath('//span[@class="label label-warning"][1]'):
                if True:
                    request = pyrus.models.requests.TaskCommentRequest(text="", action="")
                    task = pyrus_client.comment_task(res_id, request).task
                    print('')


            elif driver. find_elements_by_tag_name("tbody"):
                if True:
                    request = pyrus.models.requests.TaskCommentRequest(text="Площадка не найдена", action="")
                    task = pyrus_client.comment_task(res_id, request).task
                    print('Ошибка, площадка не найдена.')
                    resp = service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range="guid!A2",
                        valueInputOption="RAW",
                        body=body).execute()



    except KeyError:
        pass
    except ValueError:
        pass 
