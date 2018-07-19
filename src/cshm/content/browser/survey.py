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


class CalculateSatisfaction(BrowserView):
    template = ViewPageTemplateFile('template/calculate_satisfaction.pt')
    def __call__(self):
        request = self.request
        course = request.get('course')
        period = request.get('period')
        execSql = SqlObj()
        uidList = []
        for child in api.content.get(UID=course).getChildNodes():
            if child.title.encode('utf-8').split('期')[0] == period:
                for item in api.content.get(UID=child.UID()).getChildNodes():
                    uid = item.UID()
                    uidList.append(uid)
                break;

        execStr = """SELECT * FROM `satisfaction` WHERE uid in {}""".format(tuple(uidList))
        result = execSql.execSql(execStr)
        if not result:
            return 'error'
        tmp_data = {}
        anw_set = {}
        anw_set['count_A'] = []
        anw_set['count_B'] = []
        anw_set['count_C'] = []
        anw_set['count_D'] = []
        anw_set['count_E'] = []
        anw_set['count_F'] = []
        option1_data = []
        option2_data = []
        option3_data = []
        option4_data = []

        for item in result:
            tmp = dict(item)
            teacher = tmp['teacher'].strip()
            # question 1,2,3,4,5,8 為基本問題
            # 6,7 為輔導員及場地茶水問題
            # 9,10,11,12 為意見
            anwA = tmp['question1']
            anwB = tmp['question2']
            anwC = tmp['question3']
            anwD = tmp['question4']
            anwE = tmp['question5']
            anwF = tmp['question8']
            # 統計意見
            option1 = tmp['question9']
            option2 = tmp['question10']
            option3 = tmp['question11']
            option4 = tmp['question12']
            option1_data.append(option1)
            option2_data.append(option2)
            option3_data.append(option3)
            option4_data.append(option4)
            # 統計各題的回答
            anw_set['count_A'].append(anwA)
            anw_set['count_B'].append(anwB)
            anw_set['count_C'].append(anwC)
            anw_set['count_D'].append(anwD)
            anw_set['count_E'].append(anwE)
            anw_set['count_F'].append(anwF)
            # 統計老師的評分狀況
            if tmp_data.has_key(teacher):
                tmp_data[teacher].append(anwA)
                tmp_data[teacher].append(anwB)
                tmp_data[teacher].append(anwC)
                tmp_data[teacher].append(anwD)
                tmp_data[teacher].append(anwE)
                tmp_data[teacher].append(anwF)
            else:
                tmp_data[teacher] = [anwA, anwB, anwC, anwD, anwE, anwF]
        self.option1_data = option1_data
        self.option2_data = option2_data
        self.option3_data = option3_data
        self.option4_data = option4_data
        count_data = {}
        tmp_teacher_point = 0
        each_teacher_data = {}
        for k,v in tmp_data.items():
            count_5 = v.count(5)
            count_4 = v.count(4)
            count_3 = v.count(3)
            count_2 = v.count(2)
            count_1 = v.count(1)
            weight_5 = count_5 * 5
            weight_4 = count_4 * 4
            weight_3 = count_3 * 3
            weight_2 = count_2 * 2
            weight_1 = count_1 * 1
            point = round((float(weight_5) + float(weight_4) + float(weight_3) + float(weight_2) + float(weight_1)) / (float(count_5) 
		+ float(count_4) + float(count_3) + float(count_2) + float(count_1)),2)
            # 講師平均權值，加權分數再pt算
            count_data[k] = point
            # 總講師權值分數
            tmp_teacher_point += point * 20
            # 圓餅圖要顯示每個老師的個別資料
            each_teacher_data[k] = [count_5, count_4, count_3, count_2, count_1]
        self.each_teacher_data = json.dumps(each_teacher_data)
        self.count_data = count_data
        # 總講師權值分數
        self.point_teacher = round(float(tmp_teacher_point) / float(len(count_data)),2)

        tmp_space = [0, 0, 0, 0, 0]
        tmp_envir = [0, 0, 0, 0, 0]
        for item in result:
            tmp = dict(item)
            space = tmp['question6']
            environment = tmp['question7']
            if space == 5:
                tmp_space[0] += 1

            elif space == 4:
                tmp_space[1] += 1

            elif space == 3:
                tmp_space[2] += 1

            elif space == 2:
                tmp_space[3] += 1

            elif space == 1:
                tmp_space[4] += 1

            if environment == 5:
                tmp_envir[0] += 1

            elif environment == 4:
                tmp_envir[1] += 1

            elif environment == 3:
                tmp_envir[2] += 1

            elif environment == 2:
                tmp_envir[3] += 1

            elif environment == 1:
                tmp_envir[4] += 1

        self.envir_data = [tmp_envir[0], tmp_envir[1], tmp_envir[2], tmp_envir[3], tmp_envir[4]]
        self.space_data = [tmp_space[0], tmp_space[1], tmp_space[2], tmp_space[3], tmp_space[4]]

        # 計算環境分數
        origin_space = tmp_space[0] + tmp_space[1] + tmp_space[2] + tmp_space[3] + tmp_space[4] 
        weight_space = tmp_space[0] * 5 + tmp_space[1] * 4 + tmp_space[2] * 3 + tmp_space[3] * 2 + tmp_space[4] * 1
        self.point_space = round(float(weight_space) / float(origin_space) * 20, 2)

        origin_envir = tmp_envir[0] + tmp_envir[1] + tmp_envir[2] + tmp_envir[3] + tmp_envir[4] 
        weight_envir = tmp_envir[0] * 5 + tmp_envir[1] * 4 + tmp_envir[2] * 3 + tmp_envir[3] * 2 + tmp_envir[4] * 1
        self.point_envir = round(float(weight_envir) / float(origin_envir) * 20, 2)

        self.point_total = round((float(self.point_space * 10) + float(self.point_envir * 20) + float(self.point_teacher * 70)) / 100,2)
        anw_data = {}
        anw_5 = 0
        anw_4 = 0
        anw_3 = 0
        anw_2 = 0
        anw_1 = 0
        for k,v in anw_set.items():
            # 全部問題的元餅圖資料
            for item in v:
                if item == 5:
                    anw_5 += 1
                elif item == 4:
                    anw_4 += 1
                elif item == 3:
                    anw_3 += 1
                elif item == 2:
                    anw_2 += 1
                elif item == 1:
                    anw_1 += 1
            total_anw = [anw_5, anw_4, anw_3, anw_2, anw_1]
            # 單問題的圓餅圖資料
            anw_A = int(v.count(5))
            anw_B = int(v.count(4))
            anw_C = int(v.count(3))
            anw_D = int(v.count(2))
            anw_E = int(v.count(1))
            anw_data[k] = [anw_A, anw_B, anw_C, anw_D, anw_E]
        self.anw_data = anw_data
        self.total_anw = total_anw
        self.period = period
        self.course = api.content.get(UID=course).title

        return self.template()


class ShowStatistics(BrowserView):
    template = ViewPageTemplateFile('template/show_statistics.pt')
    def __call__(self):
        execSql = SqlObj()

        execStr = """SELECT DISTINCT(uid) FROM `satisfaction`"""
        uidList = execSql.execSql(execStr)
        result = []
        for i in uidList:
            uid = api.content.get(UID=i[0]).getParentNode().getParentNode().UID()
            course = api.content.get(UID=i[0]).getParentNode().getParentNode().Title()
            result.append([course, uid])

        self.result = result
        return self.template()


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

