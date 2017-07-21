import json
import datetime
from collections import namedtuple
import xlwt

'''
       Iniciando Arquivo excel
'''
wb = xlwt.Workbook()
ws = wb.add_sheet('dados')

data = []

data.append(
    '''{"event_type": "sent","user": 1,"complaint": false,"algorithms": ["trends", "abandoned-cart"],"email_provider": "uol.com.br","bounce": false, "subject": "abandonado-com-ofertas", "layout": "one-banner.html","partner": "1000", "datetime": 1500654961954, "event_id": "102dca44-eb90-4f90-a70d-b8e4c5790dcd" } ''')
data.append(
    '''{"event_type": "sent","user": 1,"complaint": false,"algorithms": ["trends", "abandoned-cart"],"email_provider": "uol.com.br","bounce": false, "subject": "abandonado-com-ofertas", "layout": "one-banner.html","partner": "1000", "datetime": 1500654961954, "event_id": "102dca44-eb90-4f90-a70d-b8e4c5790dcd" } ''')
data.append(
    '''{"event_type": "opened","user": 1,"complaint": false,"algorithms": ["trends", "abandoned-cart"],"email_provider": "uol.com.br","bounce": false, "subject": "abandonado-com-ofertas", "layout": "one-banner.html","partner": "1000", "datetime": 1500654961954, "event_id": "102dca44-eb90-4f90-a70d-b8e4c5790dcd" } ''')
data.append(
    '''{"event_type": "opened","user": 1,"complaint": false,"algorithms": ["trends", "abandoned-cart"],"email_provider": "uol.com.br","bounce": false, "subject": "abandonado-com-ofertas", "layout": "one-banner.html","partner": "1000", "datetime": 1500654961954, "event_id": "102dca44-eb90-4f90-a70d-b8e4c5790dcd" } ''')
data.append(
    '''{"event_type": "clicked","user": 1,"complaint": false,"algorithms": ["trends", "abandoned-cart"],"email_provider": "uol.com.br","bounce": false, "subject": "abandonado-com-ofertas", "layout": "one-banner.html","partner": "1000", "datetime": 1500654961954, "event_id": "102dca44-eb90-4f90-a70d-b8e4c5790dcd" } ''')
data.append(
    '''{"event_type": "clicked","user": 1,"complaint": false,"algorithms": ["trends", "abandoned-cart"],"email_provider": "uol.com.br","bounce": false, "subject": "abandonado-com-ofertas", "layout": "one-banner.html","partner": "1000", "datetime": 2500654961954, "event_id": "102dca44-eb90-4f90-a70d-b8e4c5790dcd" } ''')

result = {}

for resp in data:
    sent = 0
    opened = 0
    clicked = 0


    x = json.loads(resp, object_hook=lambda d: namedtuple('X', d.keys())(*d.values()))
    date = datetime.datetime.fromtimestamp(x.datetime / 1e3).date()
    print("RESULTADO.: %s , %s , %s" % (x.event_type, date, x.email_provider))

    if x.event_type == 'sent':
        sent = 1
    elif x.event_type == 'opened':
        opened = 1
    elif x.event_type == 'clicked':
        clicked = 1

    if date in result:
        sent += result[date][0]
        opened += result[date][1]
        clicked += result[date][2]

    result[date] = [sent, opened, clicked]

i = len(result)

for k, v in result.items():
    i += 1
    print("FOR.:", k)
    ws.write(i, 0, str(k))
    ws.write(i, 1, v[0])
    ws.write(i, 2, v[1])
    ws.write(i, 3, v[2])

wb.save('asd.xls')
