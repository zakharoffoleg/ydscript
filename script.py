# -*- coding: utf-8 -*-
from time import sleep

import json
import re
import requests
import sys
from openpyxl import Workbook

if sys.version_info < (3,):
    def u(x):
        try:
            return x.encode("utf8")
        except UnicodeDecodeError:
            return x
else:
    def u(x):
        if type(x) == type(b''):
            return x.decode('utf8')
        else:
            return x

token = 'AgAAAAAJq2baAAXKPEEN2051SEebhC-XYWWB8kQ'
login = 'zakharoffffff'


class Account(object):

    def __init__(self, token, login):
        self.token = token
        self.login = login

    campaignList = []
    CampaignsURL = 'https://api-sandbox.direct.yandex.com/json/v5/campaigns'
    ReportsURL = 'https://api-sandbox.direct.yandex.com/json/v5/reports'

    def getCampaigns(self):
        headers = {"Authorization": "Bearer " + self.token,
                   "Client-Login": self.login,
                   "Accept-Language": "ru",
                   }

        body = {"method": "get",
                "params": {"SelectionCriteria": {},
                           "FieldNames": ["Id", "Name"]
                           }}

        jsonBody = json.dumps(body, ensure_ascii=False).encode('utf8')

        try:
            result = requests.post(self.CampaignsURL, jsonBody, headers=headers)

            # Отладочная информация
            #print("Заголовки запроса: {}".format(result.request.headers))
            #print("Запрос: {}".format(u(result.request.body)))
            #print("Заголовки ответа: {}".format(result.headers))
            #print("Ответ: {}".format(u(result.text)))

            # Обработка запроса
            if result.status_code != 200 or result.json().get("error", False):
                print("Произошла ошибка при обращении к серверу API Директа.")
                print("Код ошибки: {}".format(result.json()["error"]["error_code"]))
                print("Описание ошибки: {}".format(u(result.json()["error"]["error_detail"])))
                print("RequestId: {}".format(result.headers.get("RequestId", False)))
            else:
                print("RequestId: {}".format(result.headers.get("RequestId", False)))
                print("Информация о баллах: {}".format(result.headers.get("Units", False)))

                for campaign in result.json()["result"]["Campaigns"]:
                    self.campaignList.append(str(campaign['Id']))
                    print("Рекламная кампания: {} №{}".format(u(campaign['Name']), campaign['Id']))

                if result.json()['result'].get('LimitedBy', False):
                    print("Получены не все доступные объекты.")
            print(self.campaignList)

        except ConnectionError:
            print("Произошла ошибка соединения с сервером API.")

        except:
            print("Произошла непредвиденная ошибка.")

    def getCosts(self):
        if self.campaignList:
            self.campaignCosts = {}
            for i in self.campaignList:
                print(i)
                headers = {
                    "Authorization": "Bearer " + self.token,
                    "Client-Login": self.login,
                    "Accept-Language": "ru",
                    "processingMode": "auto"
                    # Формат денежных значений в отчете
                    # "returnMoneyInMicros": "false",
                    # Не выводить в отчете строку с названием отчета и диапазоном дат
                    # "skipReportHeader": "true",
                    # Не выводить в отчете строку с названиями полей
                    # "skipColumnHeader": "true",
                    # Не выводить в отчете строку с количеством строк статистики
                    # "skipReportSummary": "true"
                }

                body = {
                    "params": {
                        "SelectionCriteria": {
                            "DateFrom": "2019-07-22",
                            "DateTo": "2019-07-23",
                            "Filter": [{
                                "Field": "CampaignId",
                                "Operator": "EQUALS",
                                "Values": [i],
                            }],
                        },
                        "FieldNames": [
                            #"Date",
                            "CampaignId",
                            "CampaignName",
                            #"LocationOfPresenceName",
                            #"Impressions",
                            #"Clicks",
                            "Cost"
                        ],
                        "ReportName": u("НАЗВАНИЕ_ОТЧЕТА"),
                        "ReportType": "CUSTOM_REPORT",
                        "DateRangeType": "CUSTOM_DATE",
                        "Format": "TSV",
                        "IncludeVAT": "YES",
                        "IncludeDiscount": "NO"
                    }
                }

                # Кодирование тела запроса в JSON
                body = json.dumps(body, indent=4)

                # --- Запуск цикла для выполнения запросов ---
                # Если получен HTTP-код 200, то выводится содержание отчета
                # Если получен HTTP-код 201 или 202, выполняются повторные запросы
                while True:
                    try:
                        req = requests.post(self.ReportsURL, body, headers=headers)
                        req.encoding = 'utf-8'  # Принудительная обработка ответа в кодировке UTF-8
                        if req.status_code == 400:
                            print("Параметры запроса указаны неверно или достигнут лимит отчетов в очереди")
                            print("RequestId: {}".format(req.headers.get("RequestId", False)))
                            print("JSON-код запроса: {}".format(u(body)))
                            print("JSON-код ответа сервера: \n{}".format(u(req.json())))
                            break
                        elif req.status_code == 200:
                            print("Отчет создан успешно")
                            # print("RequestId: {}".format(req.headers.get("RequestId", False)))
                            # print("Содержание отчета: \n{}".format(u(req.text)))
                            self.campaignCosts.update({i: req.text})
                            break
                        elif req.status_code == 201:
                            print("Отчет успешно поставлен в очередь в режиме офлайн")
                            retryIn = int(req.headers.get("retryIn", 60))
                            print("Повторная отправка запроса через {} секунд".format(retryIn))
                            print("RequestId: {}".format(req.headers.get("RequestId", False)))
                            sleep(retryIn)
                        elif req.status_code == 202:
                            print("Отчет формируется в режиме офлайн")
                            retryIn = int(req.headers.get("retryIn", 60))
                            print("Повторная отправка запроса через {} секунд".format(retryIn))
                            print("RequestId:  {}".format(req.headers.get("RequestId", False)))
                            sleep(retryIn)
                        elif req.status_code == 500:
                            print("При формировании отчета произошла ошибка. Пожалуйста, попробуйте повторить запрос позднее")
                            print("RequestId: {}".format(req.headers.get("RequestId", False)))
                            print("JSON-код ответа сервера: \n{}".format(u(req.json())))
                            break
                        elif req.status_code == 502:
                            print("Время формирования отчета превысило серверное ограничение.")
                            print("Пожалуйста, попробуйте изменить параметры запроса - уменьшить период и количество запрашиваемых данных.")
                            print("JSON-код запроса: {}".format(body))
                            print("RequestId: {}".format(req.headers.get("RequestId", False)))
                            print("JSON-код ответа сервера: \n{}".format(u(req.json())))
                            break
                        else:
                            print("Произошла непредвиденная ошибка")
                            print("RequestId:  {}".format(req.headers.get("RequestId", False)))
                            print("JSON-код запроса: {}".format(body))
                            print("JSON-код ответа сервера: \n{}".format(u(req.json())))
                            break

                    # Обработка ошибки, если не удалось соединиться с сервером API Директа
                    except requests.exceptions.ConnectionError:
                        # В данном случае мы рекомендуем повторить запрос позднее
                        print("Произошла ошибка соединения с сервером API")
                        # Принудительный выход из цикла
                        break

                    # Если возникла какая-либо другая ошибка
                    except requests.exceptions.RequestException as e:
                        # В данном случае мы рекомендуем проанализировать действия приложения
                        print(e)
                        print("Произошла непредвиденная ошибка")
                        # Принудительный выход из цикла
                        break
            print(self.campaignCosts)
        else:
            print("Список кампаний пуст!")

    def exportToExcel(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Отчёт YD"

        for i in self.campaignCosts.items():
            rows = [a.start() for a in list(re.finditer('\n', i[1]))]
            cols = [a.start() for a in list(re.finditer('\t', i[1]))]
            print(rows)
            print(cols)
            if rows and cols:
                for n in range(len(rows) + len(cols) - 2):
                    if rows[0] < cols[0]:
                        rows.pop(0)
                    else:
                        cols.pop(0)
                    print(rows)
                    print(cols)
            else:
                print("Нет списка расходов по кампаниям!")

        #wb.save(filename="asdkljbl.xslx")


User = Account(token, login)

User.getCampaigns()
User.getCosts()
User.exportToExcel()
