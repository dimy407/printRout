# -- coding: utf-8 --
import httplib2
import base64
import subprocess
import os
import json
from xml.dom.minidom import parseString


# locationWS = 'http://testshop.status-m.com.ua/ws/RouteList.1cws'
# printerName = 'Microsoft XPS Document Writer'
# AcroRd32exe = 'C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe'
###############################################################
optionFile = open('option.json', 'r+')
option = optionFile.read()
optionFile.close()
if option == '':
    print('Ошибка получения настроек!')
optionDict = json.loads(option)
locationWS = optionDict['locationWS']
printerName = optionDict['printerName']
AcroRd32exe = optionDict['AcroRd32exe']
###############################################################
fileCausesName = 'causes.txt'
###############################################################
username = ''
password = ''

filePDF = 'filePDF.pdf'


def getText(nodelist):
    rc = []
    for node in nodelist:
        if node.nodeType == node.TEXT_NODE:
            rc.append(node.data)
    return ''.join(rc)


def getRoutes(reprint=False):
    routes = []
    h = httplib2.Http()
    h.add_credentials(username, password)
    body = """
    <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:avt="http://www.avtotehcenter.loc">
       <soapenv:Header/>
       <soapenv:Body>
          <avt:GetLists>
            <avt:RepeatedPrint>%s</avt:RepeatedPrint>
          </avt:GetLists>
       </soapenv:Body>
    </soapenv:Envelope>
    """

    resp, content = h.request(locationWS, method='POST',
                              body=body % reprint,
                              headers={"Content-Type": "application/soap+xml;charset=UTF-8;\
                                    action=\"http://www.avtotehcenter.loc#RouteList:GetLists\""})

    if resp.status == 200:
        resultsXML = content.decode('utf-8')
        dom = parseString(resultsXML)
        routeNodeList = dom.getElementsByTagName("m:RouteSearchResults")
        for routeNodeListElement in routeNodeList:
            route = {
                'Code': getText(routeNodeListElement.getElementsByTagName("m:Code")[0].childNodes),
                'Date': getText(routeNodeListElement.getElementsByTagName("m:Date")[0].childNodes),
                'RouteName': getText(routeNodeListElement.getElementsByTagName("m:RouteName")[0].childNodes),
                'ShippingDate': getText(routeNodeListElement.getElementsByTagName("m:ShippingDate")[0].childNodes)
            }
            routes.append(route)
    else:
        print('Ошибка при получении данных' + content.decode('utf-8'))
    return routes


def showRoutes(routes):
    if len(routes) == 0:
        input('# Нет маршрутов для печати.')
        return 0
    i = 0
    while i < len(routes):
        print(' %s. %s / %s' % (i+1, routes[i]['ShippingDate'], routes[i]['RouteName']))
        i += 1

def showList(list):
    i = 0
    while i < len(list):
        print(' %s. %s' % (i, list[i]))
        i += 1

def printRoute(Code, Date, idDocs, causeReprint=''):
    h = httplib2.Http()
    h.add_credentials(username, password)
    body = """
    <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:avt="http://www.avtotehcenter.loc">
       <soapenv:Header/>
       <soapenv:Body>
          <avt:PrintDocumentPackage>
             <avt:Documents>
                <item>
                    <key>RouteList</key>
                    <value>%s</value>
                </item>
                <item>
                    <key>RouteList2</key>
                    <value>%s</value>
                </item>
                <item>
                    <key>Nakl</key>
                    <value>%s</value>
                </item>
                <item>
                    <key>ReglNakl</key>
                    <value>%s</value>
                </item>
                <item>
                    <key>All</key>
                    <value>%s</value>
                </item>
                <item>
                    <key>All2Ml</key>
                    <value>%s</value>
                </item>
                <item>
                    <key>TTN</key>
                    <value>False</value>
                </item>
                <item>
                    <key>Displacement</key>
                    <value>False</value>
                </item>
             </avt:Documents>
             <avt:Num>%s</avt:Num>
             <avt:Date>%s</avt:Date>
             <avt:Cause>'%s'</avt:Cause>
          </avt:PrintDocumentPackage>
       </soapenv:Body>
    </soapenv:Envelope>
    """

    #soapAction = 'http://www.avtotehcenter.loc#RouteList:GetLists'

    resp, content = h.request(locationWS, method='POST',
                              body=body % (idDocs.count('1') > 0,
                                           idDocs.count('2') > 0,
                                           idDocs.count('3') > 0,
                                           idDocs.count('4') > 0,
                                           idDocs.count('5') > 0,
                                           idDocs.count('6') > 0,
                                           Code, Date, causeReprint),
                              headers={"Content-Type": "application/soap+xml;charset=UTF-8;\
                                    action=\"http://www.avtotehcenter.loc#RouteList:GetLists\""})

    if resp.status == 200:
        resultsXML = content.decode('utf-8')
        dom = parseString(resultsXML)

        list_result = dom.getElementsByTagName("m:return")[0].childNodes

        for index in range(list_result.length):
            if list_result.item(index).localName == 'Value':
                # base64PDF = getText(dom.getElementsByTagName("m:return")[0].childNodes)
                base64PDF = getText(list_result.item(index).childNodes)
                docPDF = base64.b64decode(base64PDF)
                binfile = open(filePDF, 'wb')
                binfile.write(docPDF)
                binfile.close()

                cmd = '"%s" /N /T "%s" "%s"' % (AcroRd32exe, filePDF, printerName)
                proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                stdout, stderr = proc.communicate()
                exit_code = proc.wait()
                os.remove(filePDF)
    else:
        print('Ошибка при получении данных' + content.decode('utf-8'))


def addCause(cause):
    cause.replace(';', ' ')
    f = open(fileCausesName, 'a')
    f.write(cause + ';')
    f.close()

def getCause():
    return ['Замятие бумаги', 'Ошибка', 'Другое']
    try:
        file = open(fileCausesName, 'r', encoding='UTF-8')
    except IOError as e:
        return ['Другое']
    else:
        with file:
            causes = file.read().split(';')
            causes.remove('')  # Удаляем последний пустой єлемент
            causes.append('Другое')
            return causes


if __name__ == '__main__':
    print('####################################################################')
    print('#                            АВТОРИЗАЦИЯ                           #')
    print('####################################################################')
    print('')
    username = str(input(' Логин  : '))
    password = str(input(' Пароль : '))
    # username = 'wser'
    # password = 'a1s2d3f4g5h6'

    causeReprint = ''
    while True:
        print('####################################################################')
        print('#                      ПЕЧАТЬ МАРШРУТНЫХ ЛИСТОВ                    #')
        print('####################################################################')
        print('')
        print(' 0. Повторить печать маршрута')
        routes = getRoutes()
        showRoutes(routes)
        print('')
        print('####################################################################')
        idRoute = int(input(u'Введите номер маршрута для печати: '))
        if idRoute == 0:
            # Запросить причину повторной печати
            # 1. Получить из файла список причин
            #   - по умолчание в списке первый и последний элемент 1. Замятие бумаги ... n.Другое
            #   -
            # 2. Показать список причин
            # 3. Выбрать причину по порядковому номеру и отправить
            # 4. Организовать запись в файл причины "другое"
            causes = getCause()
            showList(causes)
            causeReprint = causes[int(input(u'Укажите причину повторной печати: '))]
            """
            if causeReprint == 'Другое':

                inputCause = True
                while inputCause:
                    causeReprint = input(u'Введите причину повторной печати: ')
                    if causeReprint !='':
                        inputCause = False
                    else:
                        print('Причина печати не должна быть пустой!')

                addCause(causeReprint)
            """
            routes = getRoutes(True)
            showRoutes(routes)
            idRoute = int(input(u'Введите номер маршрута для печати: '))
        if 1 <= idRoute <= len(routes):
            print('')
            print('####################################################################')
            print('Что печатаем?')
            print(' 1. Маршрутный лист')
            print(' 2. Маршрутный лист в двух экземплярах')
            print(' 3. Накладная')
            print(' 4. Накладная (регл)')
            print(' 5. Пакет документов')
            print(' 6. Пакет документов (Маршрутный лист в двух экземплярах)')

            print('####################################################################')
            idDocs = str(input('Вводить через пробел: ')).split(' ')

            if causeReprint == 'Другое':
                causeReprint = 'other'
            elif causeReprint == 'Замятие бумаги':
                causeReprint = 'paper jam'
            elif causeReprint == 'Ошибка':
                causeReprint = 'error'
            else:
                causeReprint = ''

            printRoute(routes[idRoute-1]['Code'], routes[idRoute-1]['Date'], idDocs, causeReprint)
        else:
            input('Неверно указан номер маршрута(нажмите любую клавишу)')
            continue
        continueWork = str(input('Продолжить работу с программой? [Да/Нет][y/n]:'))
        if continueWork == 'y' or continueWork == 'Да' or continueWork == 'yes':
            continue
        else:
            break

# todo: разнести по функциям
# Скомпилировать "C:\Python34(32bit)\python.exe" setup.py build