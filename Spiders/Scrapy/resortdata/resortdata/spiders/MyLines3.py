#!/usr/bin/env python

import os
import re
import sys
import smtplib
import datetime
import requests
import traceback
from random import choice
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

def SaveLstIntoFile(prodURLs, fileName):
    iFile = open(fileName, "ab")
    for prodURL in prodURLs:
        iFile.write(prodURL + '\r\n')
    iFile.close()

def loadProxies(fileName):
    return loadFileIntoLIst(fileName)

def loadFileIntoLIst(fileName):
    with open(fileName) as f:
        _proxies = f.read().splitlines()
    return _proxies

def PickRandomProxy(proxies):
    pickOne = choice(proxies)
    return pickOne

def removeDataFile(dataFileName):
    try:
        os.remove(dataFileName)
    except OSError:
        pass

def sendmail(receivers, body, attachement):
    # Sending mail
    username = 'kath@perfect-macro-recorder.com'
    password = 'N8]UCX)BC6NK'
    server = 'mail.perfect-macro-recorder.com'
    sender = 'kath@perfect-macro-recorder.com'
    #receivers = 'cprinahmed@yahoo.com'

    try:
        msg = MIMEMultipart()
        msg['Subject'] = 'New Data File'
        msg['From'] = sender
        msg['To'] = receivers
        msg.preamble = 'New Data File'

        fp = open(attachement, 'rb')
        afile = MIMEApplication(fp.read())
        filename = "Data-" + datetime.datetime.now().strftime("%y-%m-%d-%H-%M") + ".csv"
        afile.add_header('Content-Disposition', 'attachment', filename=filename)
        fp.close()
        body = MIMEText(body, 'plain')
        msg.attach(body)
        msg.attach(afile)

        smtpObj = smtplib.SMTP(server)
        smtpObj.login(username, password)
        smtpObj.sendmail(msg['From'], msg['To'], msg.as_string())
        print "Successfully sent email"
    except:
        print "Error: unable to send email"
        traceback.print_exc(file=sys.stdout)

def request_page(target_url, headers, proxies):
    if proxies and len(proxies) > 0:
        rndProxy = 'http://' + PickRandomProxy(proxies)
        #print 'request using proxy::' + rndProxy
        req = requests.get(target_url, headers=headers, proxies={'http': rndProxy})
    else:
        req = requests.get(target_url, headers=headers)
    return req

def prepare_field(field):
    field = re.sub('\n|\r|\t', ' ', field)
    field = field.encode('utf-8').strip()
    return field

def find_property(soup, name):
    property_tag = soup.find('td', text=re.compile(name))
    if property_tag:
        property_tag = property_tag.find_parent()
        if property_tag:
            property_value_tag = property_tag.find('td', {'class': 'value'})
            return property_value_tag.text
        else:
            return ''
    else:
        return ''