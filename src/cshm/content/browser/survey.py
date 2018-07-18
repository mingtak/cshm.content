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


class ResultSatisfaction(BrowserView):
    def __call__(self):
        request = self.request
        portal = api.portal.get()
        abs_url = portal.absolute_url()
        uid = request.get('uid')
        subjectName = request.get('subjectName')
        period = request.get('period')
        startDateTime = request.get('startDateTime')
        teacher = request.get('teacher')
        seat_number = request.get('seat_number')
        question1 = request.get('question1')
        question2 = request.get('question2')
        question3 = request.get('question3')
        question4 = request.get('question4')
        question5 = request.get('question5')
        question6 = request.get('question6')
        question7 = request.get('question7')
        question8 = request.get('question8', 0)
        question9 = request.get('question9', '')
        question10 = request.get('question10', '')
        question11 = request.get('question11', '')
        question12 = request.get('question12', '')
        execSql = SqlObj()

        execStr = """SELECT id FROM satisfaction WHERE uid = '{}' AND seat = '{}'""".format(uid, seat_number)

        if execSql.execSql(execStr):
            api.portal.show_message(message='請勿重複填寫問卷'.decode('utf-8'), type='error', request=request)
        else:
            execStr = """INSERT INTO `satisfaction`(uid, seat, `date`, 
                `teacher`, `question1`, `question2`, `question3`, `question4`, `question5`, 
                `question6`, `question7`, `question8`,question9,question10,question11,question12) 
                VALUES ('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}',
                '{}','{}')""".format(uid, int(seat_number), startDateTime, teacher,
                question1, question2, question3, question4, question5,question6,
                question7, question8, question9, question10, question11, question12)
            execSql.execSql(execStr)
            # 寄信通知
            if question9 or question10 or question11 or question12:
                urlList = api.content.get(UID=uid).absolute_url_path().split('/')
                course = api.content.find(path="/%s/%s/%s" %(urlList[1], urlList[2], urlList[3]), depth=0)[0].Title
                period = api.content.find(path="/%s/%s/%s/%s" %(urlList[1], urlList[2], urlList[3], urlList[4]), depth=0)[0].Title

                body_str = """科目:%s<br>課程:%s<br>期數:%s<br/>座號:%s<br>意見提供:<br>%s<br/>%s<br/>%s<br>%s
                    """ %(course, subjectName, period, seat_number, question9, question10, question11, question12)
                mime_text = MIMEText(body_str, 'html', 'utf-8')
                api.portal.send_email(
                    #recipient="lin@cshm.org.tw",
                    recipient="ah13441673@gmail.com",
                    sender="henry@mingtak.com.tw",
                    subject="%s-%s  意見提供" %(course, period),
                   body=mime_text.as_string(),
                )

            api.portal.show_message(message='填寫完成'.decode('utf-8'), type='info', request=request)
        contentURL = api.content.get(UID=uid).absolute_url().split('/')
        contentURL.pop()
        url = ''
        for i in contentURL:
            url += i + '/'
        request.response.redirect('%scheck_survey?seat_number=%s' %(url, seat_number))


class CheckSurvey(BrowserView):
    template = ViewPageTemplateFile('template/check_survey.pt')
    #finished = ViewPageTemplateFile('template/finished.pt')
    satisfaction_first = ViewPageTemplateFile('template/satisfaction_first.pt')
    satisfaction_sec = ViewPageTemplateFile('template/satisfaction_sec.pt')

    def __call__(self):
        request = self.request
        portal = api.portal.get()
        context = self.context
        abs_url = portal.absolute_url()
        date = request.get('date', '')
        period = request.get('period', '')
        teacher = request.get('teacher', '')
        subject_name = request.get('subject_name', '')
        seat_number = request.get('seat_number', '')
        uid = request.get('uid', '')

        if seat_number:
            now = datetime.datetime.now()
            alreadyWrite = []
            ex_data = []

            execSql = SqlObj()
            execStr = "SELECT uid FROM satisfaction WHERE seat = '%s'" %(seat_number)
            uidList = execSql.execSql(execStr)

            for item in uidList:
                alreadyWrite.append(item[0])
            for content in api.content.find(context=context, depth=1, sort_on='startDateTime', sort_order='descending'):
                obj = content.getObject()
                contentUid = obj.UID()
                if contentUid not in alreadyWrite and now >= content.startDateTime:
                    url = "%s/check_survey?seat_number=%s&uid=%s" %(context.absolute_url(), seat_number, contentUid)
                    if uid:
                        if contentUid != uid:
                            ex_data.append(['%s-%s' %(obj.startDateTime, obj.title), url])
                        else:
                            self.setData(context, obj, seat_number)
                            ex_data.insert(0, ['請選擇', '請選擇'])
                    else:
                        if ex_data:
                            ex_data.append(['%s-%s' %(obj.startDateTime, obj.title), url])
                        else:
                            self.setData(context, obj, seat_number)
                            ex_data = [['請選擇', '請選擇']]

            self.ex_data = ex_data
            if self.firstTarget:
                return self.satisfaction_sec()
            else:
                return self.satisfaction_first()

        return self.template()

    def setData(self, context, obj, seat_number):
        self.period = context.title.encode('utf-8').split('期')[0]
        self.subjectName = obj.title
        self.startDateTime = obj.startDateTime.strftime('%Y-%m-%d %H:%M:%S')
        self.teacher = obj.teacher.to_object.title
        self.seat_number = seat_number
        self.firstTarget = obj.isQuiz
        self.uid = obj.UID()


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
            execStr = "SELECT seat FROM satisfaction WHERE uid = '%s' ORDER BY seat" %child_uid
            seat = execSql.execSql(execStr)
            content = api.content.get(UID = child_uid)
            subject_name = content.title

        url = "%s/check_surver?course_name=%s&period=%s" %(abs_url, course_name, period)
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

