# -*- coding: utf-8 -*-
import csv
import xml.dom.minidom
import os
import time

import base36
import openpyxl
import requests


def getImporter(alcoCode):
    b36alcoCode = '000000000000000000' + base36.dumps(int(alcoCode))
    b36alcoCode = b36alcoCode[-16:]

    getEgaisik = requests.get(
        'https://xn--80affoam1c.xn--p1ai/api/testscan.php?barcode=22N' + b36alcoCode + '0115606220020050311D1WNUMCGE12WXYYUNNLTPAE1SNK9UQ')
    ansEgaisik = getEgaisik.text.split('</br>')
    if len(ansEgaisik) < 20:
        return 'http://www.fsrar.ru/frap/frap', 'ИЩИ!!!', 'На сайте'
    Importer = ansEgaisik[-7][10:]
    INN = ansEgaisik[-4][5:]
    KPP = ansEgaisik[-3][5:]
    return Importer, INN, KPP


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
                        ask = getImporter(line.getElementsByTagName('pref:AlcCode')[0].childNodes[0].nodeValue)
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
