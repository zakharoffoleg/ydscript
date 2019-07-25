# -*- coding: utf-8 -*-
import requests, json
from requests.exceptions import ConnectionError
from time import sleep
import sys

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

    campaignlist = []
    CampaignsURL = 'https://api-sandbox.direct.yandex.com/json/v5/campaigns'
    ReportsURL = 'https://api-sandbox.direct.yandex.com/json/v5/reports'

    def getcampaigns(self):
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
            #print("\n")

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
                    self.campaignlist.append(str(campaign['Id']))
                    print("Рекламная кампания: {} №{}".format(u(campaign['Name']), campaign['Id']))

                if result.json()['result'].get('LimitedBy', False):
                    print("Получены не все доступные объекты.")
            print(self.campaignlist)


        except ConnectionError:
            print("Произошла ошибка соединения с сервером API.")

        except:
            print("Произошла непредвиденная ошибка.")

    def getcosts(self):
        if self.campaignlist:
            for i in self.campaignlist:
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
                            "CampaignName",
                            "CampaignId",
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
                            print("RequestId: {}".format(req.headers.get("RequestId", False)))
                            print("Содержание отчета: \n{}".format(u(req.text)))
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
                    except ConnectionError:
                        # В данном случае мы рекомендуем повторить запрос позднее
                        print("Произошла ошибка соединения с сервером API")
                        # Принудительный выход из цикла
                        break

                    # Если возникла какая-либо другая ошибка
                    except:
                        # В данном случае мы рекомендуем проанализировать действия приложения
                        print("Произошла непредвиденная ошибка")
                        # Принудительный выход из цикла
                        break
        else:
            print("Список кампаний пуст!")


User = Account(token, login)

User.getcampaigns()
User.getcosts()
