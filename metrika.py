import json
import time

import openpyxl
import requests

start_time = time.time()

wb = openpyxl.load_workbook('report.xlsx')

YDAuth = {'za4aroff.ol': 'AgAAAAAzN9ktAAXAInwzJfwlY0E4taHqCQUC33I',
          'macsim.reshetar': 'AgAAAAArGY4IAAXAIq_OsOIJXEgIvBjrf0dpP0s'}

direct_client_logins = ["za4aroff.ol", "za4aroff.ol", "macsim.reshetar"]  # НП СПБ, НП МСК, 1001
clientTokens = {"za4aroff.ol": "AgAAAAAzN9ktAAXAInwzJfwlY0E4taHqCQUC33I",
                "macsim.reshetar": "AgAAAAArGY4IAAXAIq_OsOIJXEgIvBjrf0dpP0s"}
date1 = "2019-09-01"
date2 = "yesterday"
ids = ["35656455", "37340325", "42890799"]  # НП СПБ, НП МСК, 1001
metrics = ["ym:s:goal39547444visits", "ym:s:goal39545641visits", "ym:s:goal44259842visits"]  # НП СПБ, НП МСК, 1001
dimensions = "ym:s:goal,ym:s:lastsignDirectClickOrder,ym:s:lastsignDirectBannerGroup"
filters = ["ym:s:goal==39547444", "ym:s:goal==39545641", "ym:s:goal==44259842"]
url = "https://api-metrika.yandex.net/stat/v1/data"


def saveConversions(login_list, clientTokens, directIDs, metrics, startDate, endDate, filters):
    if 'Конверсии' in wb.sheetnames:
        report = wb['Конверсии']
    else:
        report = wb.create_sheet('Конверсии')

    for n, field in enumerate(['Кампания', 'Группа', 'Конверсии']):  # Вводим поля таблицы
        _ = report.cell(column=1 + n, row=1, value=field)

    for n, acc in enumerate(login_list):
        requestURL = url + "?direct_client_logins=%s&ids=%s&metrics=%s&date1=%s&date2=%s&dimensions=%s&filters=%s" \
                           "&pretty=true" % (acc, directIDs[n], metrics[n], startDate, endDate, dimensions, filters[n])

        headers = {"Authorization": "OAuth " + clientTokens.get(acc), "Client-Login": acc, }
        req = requests.get(requestURL, headers=headers)
        # print(req.text)
        jsonReport = json.loads(req.text)

        maxRow = report.max_row
        for index, item in enumerate(jsonReport["data"]):
            if item['dimensions'][1]['name']:
                _ = report.cell(row=index + maxRow + 1, column=1, value=item['dimensions'][1]['name'])
                _ = report.cell(row=index + maxRow + 1, column=2, value=item['dimensions'][2]['name'])
                _ = report.cell(row=index + maxRow + 1, column=3, value=int(item['metrics'][0]))
                print("Найдены конверсии в кампании %s, в группе %s в количестве: %s." % \
                      (item['dimensions'][1]['name'], item['dimensions'][2]['name'], int(item['metrics'][0])))
        wb.save(filename="report.xlsx")


def connectConversionsWithCampaigns():  #Добавляем конверсии в таблицу с данными
    data = wb["Данные"]
    conversions = wb["Конверсии"]
    for n, field in enumerate(["Конверсии", "% конверсии", "CPA"]):  # Вводим поля таблицы
        _ = data.cell(column=7 + n, row=1, value=field)

    for conv in range(1, conversions.max_row):
        for camp in range(1, data.max_row):
            if data["C"][camp].value == conversions["B"][conv].value:
                if data["B"][camp].value == conversions["A"][conv].value:
                    print("Выгружаю конверсии для %s из кампании %s." %
                          (conversions["B"][conv].value, conversions["A"][conv].value))
                    _ = data.cell(row=camp + 1, column=7, value=conversions["C"][conv].value)
                    if data["E"][camp].value > 0:
                        _ = data.cell(row=camp + 1, column=8,
                                      value=conversions["C"][conv].value / data["E"][camp].value)
                    else:
                        _ = data.cell(row=camp + 1, column=8, value=0)
                    _ = data.cell(row=camp + 1, column=9, value=float(data["F"][camp].value)
                                                                / conversions["C"][conv].value)

    for row in range(1, data.max_row):
        data["H"][row].number_format = '0.00%'
        data["I"][row].number_format = '0.00'
    wb.save(filename="report.xlsx")


def mainFunc(login_list, clientTokens, directIDs, metrics, startDate, endDate, filters):
    saveConversions(login_list, clientTokens, directIDs, metrics, startDate, endDate, filters)
    connectConversionsWithCampaigns()


mainFunc(direct_client_logins, clientTokens, ids, metrics, date1, date2, filters)
print("\n--- %s seconds ---" % (time.time() - start_time))
