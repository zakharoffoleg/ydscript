import json

import openpyxl
import requests

wb = openpyxl.load_workbook('C:/Users/KK/PycharmProjects/ydscript/report.xlsx')

YDAuth = {'za4aroff.ol': 'AgAAAAAzN9ktAAXAInwzJfwlY0E4taHqCQUC33I',
          'macsim.reshetar': 'AgAAAAArGY4IAAXAIq_OsOIJXEgIvBjrf0dpP0s'}

direct_client_logins = "za4aroff.ol"
ids = "35656455"
goalID = "39547444"
metrics = "ym:s:goal%svisits" % goalID
dimensions = "ym:s:goal,ym:s:lastsignDirectClickOrder,ym:s:lastsignDirectBannerGroup"
filters = "ym:s:goal==%s" % goalID
url = "https://api-metrika.yandex.net/stat/v1/data"
requestURL = url + "?direct_client_logins=%s&ids=%s&metrics=%s&dimensions=%s&filters=%s&pretty=true" % \
             (direct_client_logins, ids, metrics, dimensions, filters)

headers = {
    "Authorization": "OAuth " + 'AgAAAAAzN9ktAAXAInwzJfwlY0E4taHqCQUC33I',
    "Client-Login": 'za4aroff.ol',
}

req = requests.get(requestURL, headers=headers)
# req.encoding = 'utf-8'  # Принудительная обработка ответа в кодировке UTF-8
print(req.text)


def saveConversions():
    jsonReport = json.loads(req.text)
    report = wb.create_sheet('Конверсии')

    for n, field in enumerate(['Кампания', 'Группа', 'Конверсии']):  # Вводим поля таблицы
        _ = report.cell(column=1 + n, row=1, value=field)

    row = 1
    for index, item in enumerate(jsonReport["data"]):
        # print(index, item)
        if item['dimensions'][1]['name'] != None:
            print(item['dimensions'][1]['name'])
            _ = report.cell(row=index + 1, column=1, value=item['dimensions'][1]['name'])
            print(item['dimensions'][2]['name'])
            _ = report.cell(row=index + 1, column=2, value=item['dimensions'][2]['name'])
            print(int(item['metrics'][0]))
            _ = report.cell(row=index + 1, column=3, value=int(item['metrics'][0]))
    wb.save(filename="report.xlsx")


def connectConversionsWithCampaigns():
    data = wb["Данные"]
    conversions = wb["Конверсии"]
    for conv in conversions["B"][1:]:
        for i in range(1, data.max_row - 1):
            if data["C"][i].value == conv.value:
                print("Совпадение найдено! %s") % conv.value


saveConversions()
connectConversionsWithCampaigns()
