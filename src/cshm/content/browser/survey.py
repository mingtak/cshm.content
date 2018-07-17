# -*- coding: utf-8 -*- 
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from plone.protect.auto import safeWrite
from mingtak.ECBase.browser.views import SqlObj
import json
import csv
import base64
import qrcode
import datetime
from plone.namedfile.field import NamedBlobImage,NamedBlobFile
from plone import namedfile
from StringIO import StringIO
import requests
from email.mime.text import MIMEText
import xlsxwriter
import inspect


class CheckSurvey(BrowserView):
    template = ViewPageTemplateFile('template/check_survey.pt')
    #finished = ViewPageTemplateFile('template/finished.pt')
    def __call__(self):
        request = self.request
        portal = api.portal.get()
        context = self.context
        abs_url = portal.absolute_url()
        seat_number = request.get('seat_number', '')
        if seat_number:
            now = datetime.datetime.now()
            alreadyWrite = []
            execSql = SqlObj()
            execStr = """SELECT uid FROM satisfaction WHERE seat = '%s'""" %(seat_number)
            uidList = execSql.execSql(execStr)

            for item in uidList:
                alreadyWrite.append(item[0])
            for content in api.content.find(context=context, depth=1, sort_on='startDateTime', sort_order='descending'):
                obj = content.getObject()
                uid = obj.UID()
                if uid not in alreadyWrite and now >= content.startDateTime:
                    period = context.title
                    subjectName = obj.title
                    startDateTime = obj.startDateTime.strftime('%Y-%m-%d %H:%M:%S')
                    teacher = obj.teacher.to_object.title
                    if obj.isQuiz:
                        url = """{}/@@satisfaction_sec?subject_name={}&date={}&teacher={}&uid={}&period={}&seat_number={}
                            """.format(abs_url, subjectName, startDateTime, teacher, uid, period, seat_number)
                        url = """{}/@@satisfaction_first?subject_name={}&date={}&teacher={}&uid={}&period={}&seat_number={}
                            """.format(abs_url, subjectName, startDateTime, teacher, uid, period, seat_number)
                        break;
            request.response.redirect(url)
            return

        return self.template()



class EchelonView(BrowserView):
    template = ViewPageTemplateFile('template/echelon_view.pt')
    def __call__(self):
#尚未完成
        context = self.context
        execSql = SqlObj()
	import pdb;pdb.set_trace()
        data = []
        abs_url = api.portal.get().absolute_url()
	context_uid = context.UID()

	for child in context.getChildNodes():
            child_uid = child.UID()
            execStr = """SELECT seat FROM satisfaction WHERE uid = '%s' ORDER BY seat""" %child_uid
            seat = execSql.execSql(execStr)
            content = api.content.get(UID = child_uid)
            subject_name = content.title

        url = """{}/check_surver?course_name={}&period={}""".format(abs_url, course_name, period)
        # 製作qrcode
        qr = qrcode.QRCode()
        qr.add_data(url)
        qr.make_image().save('url.png')
        img = open('url.png', 'rb')
        b64_img = base64.b64encode(img.read())
        self.url = url
        self.b64_img = b64_img
        self.data = data
        return self.template()

