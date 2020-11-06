# -*- coding: utf-8 -*-
import csv
import xml.dom.minidom
import os
import time
import openpyxl
import requests
cook='foav0m0k120tcsrq32n82pj0h6'
def getImporter(name):

    name = name.replace(',', '').replace('\'', '')
    r = requests.post('https://fsrar.gov.ru/frap/frap', data={'FrapForm[name_prod]': name},
                      cookies={'PHPSESSID': cook}, verify=False)
    # print(r.text)

    Answer = r.text
    Answer = Answer.split('<td ><b>Уведомитель</b></td>')[1].split('<td ><b>Производители</b></td>')[0]
    Answer = Answer.replace('\r', '').replace('\n', '').replace('<tr>', '').replace('</td>', '').replace('<td>',
                                                                                                         '').replace(
        '    ', '')

    Answer = Answer[1:-1].split('<br />')
    Importer = Answer[0]

    if len(Importer) < 3:
     return 'http://www.fsrar.ru/frap/frap', 'ИЩИ!!!', 'На сайте'
    INN, KPP = Answer[1].split(',')
    return Importer, INN[5:], KPP[6:]


def xmlParser(file):
    dom = xml.dom.minidom.parse(file)
    dom.normalize()
    xml_StockPosition = dom.getElementsByTagName('rst:StockPosition')
    if len(xml_StockPosition) == 0:
        xml_StockPosition = dom.getElementsByTagName('rst:ShopPosition')
    AlcoList = []
    for line in xml_StockPosition:
        tList = []
        for V in ['pref:AlcCode', 'pref:FullName', 'pref:Capacity', 'pref:ProductVCode', 'oref:UL', 'rst:Quantity']:
            if V == 'oref:UL' or V == 'pref:Capacity':
                try:
                    line.getElementsByTagName(V)[0].childNodes[0].nodeValue
                except:
                    if V != 'pref:Capacity':
                        ask = getImporter(line.getElementsByTagName('pref:FullName')[0].childNodes[0].nodeValue)
                        Importer, INN, KPP = ask[0], ask[1], ask[2]
                        tList.append(Importer)
                        tList.append(INN)
                        tList.append(KPP)
                    else:
                        tList.append('Нет Тары')
                else:
                    if V != 'pref:Capacity':
                        INN = line.getElementsByTagName('oref:INN')[0].childNodes[0].nodeValue
                        KPP = line.getElementsByTagName('oref:KPP')[0].childNodes[0].nodeValue
                        Importer = line.getElementsByTagName('oref:FullName')[0].childNodes[0].nodeValue
                        tList.append(Importer)
                        tList.append(INN)
                        tList.append(KPP)
                    else:
                        node = line.getElementsByTagName(V)[0].childNodes[0].nodeValue
                        tList.append(node)

            else:
                node = line.getElementsByTagName(V)[0].childNodes[0].nodeValue
                if V == 'rst:Quantity':
                    node = float(node)
                tList.append(node)
        AlcoList.append(tList)

    return AlcoList
a = str(input('Enter PHPSESSID from cookies: '))

if len(a)== 26:cook=str(a)
print('Start to find xml files....     \n')
files = os.listdir('./')
xml_f = []
for x in files:
    if x[-3:]=='xml':
        xml_f.append(x)
        print('Find xml file: '+x)

a = [['Алкокод', 'Наименование', 'Объём', 'Код вида', 'Импортер/производитель', 'ИНН', 'КПП', 'Остатки']]
print('\n'*2)
for fn in xml_f:
    print('Start xml parsing....:     '+fn)
    temp = xmlParser(fn)

    tf = open(fn+'.csv', "w", newline='')
    csv.writer(tf, delimiter=';').writerow(['Алкокод', 'Наименование', 'Объём', 'Код вида', 'Импортер/производитель', 'ИНН', 'КПП', 'Остатки'])
    for i in temp:
        csv.writer(tf, delimiter=';').writerow(i)
        for y in a:
            if i[1] == y[1]:
                y[7] += i[7]
            else:
                a.append(i)
                break
    print('End xml parsing....:     '+fn)
    print('---\n')
print('Save to Exel....')
xls_file = openpyxl.Workbook()
sheet = xls_file.active
for i in a:
    sheet.append(i)
xls_file.save(time.strftime("%Y%m%d%H%M%S", time.localtime()) + '_result.xlsx')
print('Save file to '+time.strftime("%Y%m%d%H%M%S", time.localtime()) + '_result.xlsx')
print('-------------------------------------\n')
input('All Done!.... Press any key')
