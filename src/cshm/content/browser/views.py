# -*- coding: utf-8 -*-
from cshm.content import _
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from Products.CMFPlone.utils import safe_unicode
from Products.CMFPlone.PloneBatch import Batch
from Acquisition import aq_inner
from plone.app.contenttypes.browser.folder import FolderView

from plone.app.contentlisting.interfaces import IContentListing
from Products.CMFCore.utils import getToolByName

from mingtak.ECBase.browser.views import SqlObj
from docxtpl import DocxTemplate, Listing
from plone.protect.interfaces import IDisableCSRFProtection
from zope.interface import alsoProvides
from DateTime import DateTime

import xlwt, xlrd
from xlutils.copy import copy
import base64
import requests
import re
import json
import logging
import datetime
import csv
from StringIO import StringIO
import random

logger = logging.getLogger("cshm.content")


class BasicBrowserView(BrowserView):
    """ 通用method """

    no_permission_template = ViewPageTemplateFile('template/no_permission_template.pt')

    def checkPermission(self):
        """ 檢查權限: 配合 mingtak.viewPermissions 使用 """

        userId = api.user.get_current().id
        viewName = self.__name__

        if userId == 'admin': #TODO: admin or role 有 Manager 即可？
            return True

        permissions_table = api.portal.get_registry_record('mingtak.viewPermissions.browser.configlet.IViewPermissions.permissions_table')
        permissions_table = json.loads(permissions_table)
        userList = permissions_table.get(viewName, [])
        if userId in userList:
            return True
        else:
            return False


    def checkIDCard(self, ID):
        """ 檢查身份證正確性 """

        first = {'a':10, 'b':11, 'c':12, 'd':13, 'e':14, 'f':15, 'g':16, 'h':17, 'j':18, 'k':19, 'l':20, 'm':21, 'n':22, 'p':23, 'q':24, 'r':25, 's':26, 't':27, 'u':28, 'v':29, 'w':32, 'x':30, 'y':31, 'z':33, 'i':34, 'o':35 }
        sec = [1,9,8,7,6,5,4,3,2,1,1]

        if len(ID) != 10:               #字數檢查 length check
            return False

        if not ID[1:9].isnumeric():   #2～10碼檢查 [1:9] not str check 
            return False

        if ID[0].isnumeric():         #1碼檢查 [0] not int check
            return False
        ID = ID.lower()
        ID = list(ID)

        if int(ID[1])!=1 and int(ID[1])!=2:         #2碼檢查 [1] check
            return False

        if len(ID) == 10:             #驗證 (N0 十位數 + (N0 個位數 x 9) + (N1 x 8) + (N2 x 7) +  (N3 x 6) +  (N4 x 5) +  (N5 x 4) +  (N6 x 3) +  (N7 x 2) + N8 + N9)
            ID[0] = first[ID[0]]
            ID00 = ID[0]
            del ID[0]
            ID.insert(0,int(ID00 % 10))
            ID.insert(0,int(ID00 / 10))
            total = 0

            for x,y in zip(ID,sec):
                s = int(x) * int(y)
                total+=s

            if total%10 ==0:
                return True
            else:
                return False

    def twYear2Ad(self, twStr):
        """ 民國date轉西元 """
        year, month, day = twStr.split('/')
        year = int(year) + 1911
        return '%s/%s/%s' % (str(year), month, day)

    def adYear2Tw(self, adStr):
        """ 西元date轉民國 """
        year, month, day = adStr.split('/')
        year = int(year) - 1911
        return '%s/%s/%s' % (str(year), month, day)

    def sendMessage(self, cell, message):
        settingStr = api.portal.get_registry_record('cshm.content.browser.configlet.IOffice.cell_msg_url')
        uid, pwd, url = settingStr.split(',')
        requests.get('%s?UID=%s&PWD=%s&SB=簡訊測試&MSG=%s&DEST=%s&ST=' % (url, uid, pwd, message, cell))

    def getOnTraining(self):
        """ 取得 on_training_status 選項資訊 """
        sqlStr = "SELECT * FROM `on_training_status` WHERE 1"
        sqlInstance = SqlObj()
        result = sqlInstance.execSql(sqlStr)
        return result

    def getRegCourse(self, id):
        """ 取得報名表詳細資訊 """
        uid = self.context.UID()

        sqlStr = "SELECT * FROM `reg_course` WHERE uid = '%s' and id = %s" % (uid, id)
        sqlInstance = SqlObj()
        result = sqlInstance.execSql(sqlStr)
        return result[0]

    def updateContactLog(self, regCourse, text):
        """ 更新聯絡記錄 """
        currentName = self.currentName()
        if regCourse['contactLog']:
            logData = '%s\n%s / %s / %s' % \
                (regCourse['contactLog'], currentName, DateTime().strftime('%Y-%m-%d %H:%M'), safe_unicode(text))
        else:
            logData = '%s / %s / %s' % (currentName, DateTime().strftime('%Y-%m-%d %H:%M'), safe_unicode(text))

        sqlInstance = SqlObj()
        sqlStr = "update reg_course set contactLog = '%s' where id = %s" % (logData, regCourse['id'])
        sqlInstance.execSql(sqlStr)

    def currentName(self):
        """ 取得登入者名稱 """
        current = api.user.get_current()
        return current.getProperty('fullname')

    def getTrainingStatusCode(self):
        """ 取得受訓狀態碼 """
        sqlStr = """SELECT * FROM training_status_code WHERE 1 ORDER BY id"""
        sqlInstance = SqlObj()
        return sqlInstance.execSql(sqlStr)

    def getCityList(self):
        """ 取得縣市列表 """
        return api.portal.get_registry_record('mingtak.ECBase.browser.configlet.ICustom.citySorted')

    def getDistList(self):
        """ 取得鄉鎮市區列表及區碼 """
        return api.portal.get_registry_record('mingtak.ECBase.browser.configlet.ICustom.distList')

    def isFrontend(self):
        """ 檢查是否前端頁面 """
        isFrontendView = api.content.get_view(name='is_frontend', context=self.portal, request=self.request)
        return isFrontendView(self)

    def checkCell(self, value):
        """ 檢查手機號碼是否09開頭 """
        if value.isdigit() and value.startswith('09'):
            return True
        else:
            return False

    def checkEmail(self, value):
        """ 檢查email格式 """
        pattern = re.compile(r"[^@\s]+@[^@\s]+\.[a-zA-Z0-9]+$")
        try:
            if bool(re.match(pattern, value)):
                return True
            else:
                return False
        except:
            return False

    def getEducation(self):
        """ 取得學歷代碼 """
        sqlStr = "SELECT * FROM `education_code` WHERE 1"
        sqlInstance = SqlObj()
        return sqlInstance.execSql(sqlStr)

    def insertOldStudent(self, form):
        """ 新增舊學員資料 """
        portal = api.portal.get()
        context = self.context
        request = self.request

        sqlStr = "INSERT INTO `old_student`(`studId`, `name`, `birthday`, `phone`, `cellphone`, `fax`, `priv_email`, `priv_email2`, `address`) \
                  VALUES ('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')" % \
                  ( form.get('studId').upper(), form.get('name'), form.get('birthday'), form.get('phone'), form.get('cellphone'),
                    form.get('fax'), form.get('priv_email'), form.get('priv_email2'), form.get('address') )
        sqlInstance = SqlObj()
        try:
            sqlInstance.execSql(sqlStr)
        except:pass

    def getUidCourseData(self, uid):
        """抓符合UID的reg_course資料"""
        sqlStr = """SELECT reg_course.*,training_status_code.status,education_code.degree FROM reg_course,training_status_code,
                    education_code WHERE uid = '{}' and reg_course.training_status = training_status_code.id AND
                    reg_course.education_id = education_code.id
                    """.format(uid)
        sqlInstance = SqlObj()
        return sqlInstance.execSql(sqlStr)



# 證照費用核銷狀態追蹤
class LicenseFeeTrace(BasicBrowserView):

    template = ViewPageTemplateFile("template/license_fee_trace.pt")

    def everExam(self, reg_id):
        sqlInstance = SqlObj()
        # 找最近一次考試狀態
        sqlStr = """SELECT id
                    FROM grade_mana
                    WHERE reg_id = {}
                      AND status in (2, 3)""".format(reg_id)
        result = sqlInstance.execSql(sqlStr)
        if result:
            return True
        else:
            return False


    def getDetail(self, reg_id):

        sqlInstance = SqlObj()
        # 找最近一次考試狀態
        sqlStr = """SELECT *, reg_course.id AS reg_course_id, grade_mana.id AS grade_mana_id
                    FROM reg_course, grade_mana
                    WHERE reg_course.id = grade_mana.reg_id
                    AND reg_course.id = {}
                    ORDER BY grade_mana.id DESC""".format(reg_id)
# TODO:要確認是否只會有一開始會有「未報考」的狀態？
        return sqlInstance.execSql(sqlStr)


    def __call__(self):

        sqlInstance = SqlObj()
        # 找到要被追蹤的學生名單
        sqlStr = """SELECT reg_course.id
                    FROM reg_course, grade_mana
                    WHERE (path LIKE 'cshm/mana_course/a005%%'
                      OR path LIKE 'cshm/mana_course/a007%%'
                      OR path LIKE 'cshm/mana_course/a008%%'
                      OR path LIKE 'cshm/mana_course/a009%%'
                      OR path LIKE 'cshm/mana_course/a011%%'
                      OR path LIKE 'cshm/mana_course/a012%%'
                      OR path LIKE 'cshm/mana_course/a013%%'
                      OR path LIKE 'cshm/mana_course/d001%%'
                      OR path LIKE 'cshm/mana_course/d003%%'
                      OR path LIKE 'cshm/mana_course/a001%%')
                      AND reg_course.id = grade_mana.reg_id
                    GROUP BY reg_course.id""" #a001是測試用的

        self.students = sqlInstance.execSql(sqlStr)

        return self.template()


# 成績追蹤管理系統
class GradeManage(BasicBrowserView):

    template = ViewPageTemplateFile("template/grade_manage.pt")

    def checkAddNew(self):
        request = self.request

        status = [0, 0]
        for item in request.form.items():
            if item[0].startswith('new'):
                logger.info('key:%s, value:%s' % (item[0], item[1]))
                status[0] += 1
                if item[1]:
                    status[1] += 1
        if status[1] == 0:
            return True
        elif status[0] == status[1] and status[0] > 1:
            return True
        else:
            return False


    def addNewGrade(self):
        """ 新增一梯次 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()

        sqlStr = """SELECT MAX(exam_step) AS max_step FROM grade_mana WHERE uid = '{}'""".format(context.UID())
        result = sqlInstance.execSql(sqlStr)
        step = result[0]['max_step'] + 1 if result[0]['max_step'] else 1
        newdate = request.form.get('newdate')

        for item in request.form.keys():
            if item.startswith('new-'):
                sqlStr = """INSERT INTO grade_mana (exam_date, exam_step, reg_id, `status`, uid)
                            VALUES ('{}', {}, {}, {}, '{}')
                         """.format(newdate, step, item.split('-')[1], request.form.get(item), context.UID())
                sqlInstance.execSql(sqlStr)


    def updateGrade(self):
        """ 更新梯次表 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()

        for item in request.form.keys():
            if item.startswith('old-'):
                step = item.split('-')[1]
                reg_id = item.split('-')[2]
                id = item.split('-')[3]
                exam_date = request.form.get('olddate-%s' % step)
                status = request.form.get(item)

                if not id:
                    sqlStr = """INSERT INTO grade_mana (exam_date, exam_step, reg_id, `status`, uid)
                                 VALUES ('{}', {}, {}, {}, '{}')
                             """.format(exam_date, step, reg_id, status, context.UID())
                    sqlInstance.execSql(sqlStr)
                    continue

                sqlStr = """UPDATE `grade_mana`
                            SET `reg_id`={}, `exam_step`={}, `exam_date`='{}', `status`={}, `uid`='{}'
                            WHERE id={}
                         """.format(reg_id, step, exam_date, request.form.get(item), context.UID(), id)
                sqlInstance.execSql(sqlStr)


    def getStep(self):
        """ 取得梯次數 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()

        uid = context.UID()
        sqlStr = """SELECT DISTINCT `exam_step` FROM grade_mana WHERE uid = '{}'""".format(uid)
        result = sqlInstance.execSql(sqlStr)
        return result


    def getExamDate(self, exam_step):
        """ 取得各梯次日期 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()
        sqlStr = """SELECT exam_date FROM grade_mana WHERE exam_step = {}""".format(exam_step)
        result = sqlInstance.execSql(sqlStr)
        return result


    def getGrade(self, reg_id):
        """ 取得學員梯次成績 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()
        sqlStr = """SELECT exam_step, status, id FROM grade_mana WHERE reg_id = {}""".format(reg_id)
        result = sqlInstance.execSql(sqlStr)

        value = {}
        for item in result:
            value[item['exam_step']] = [item['id'], item['status']]
        return value


    def __call__(self):
        request = self.request

        if not self.checkPermission():
            return self.no_permission_template()

        if request.form.has_key('update'):
            if self.checkAddNew():
                self.addNewGrade()
                self.updateGrade()
                api.portal.show_message(message='Update Grade Complete!', request=request)
                return self.template()
            else:
                return '新增資料輸入有誤，請回上一頁'

        return self.template()


# 成績追蹤管理系統-匯出表格(一般)
class GradeManageExportNormal(GradeManage):

    template = ViewPageTemplateFile("template/grade_manage_export_normal.pt")

    def getSum(self, step):
        sqlInstance = SqlObj()
        sqlStr = """SELECT status FROM grade_mana WHERE uid='{}' and exam_step={}""".format(self.context.UID(), step)
        result = sqlInstance.execSql(sqlStr)
        value = {1:0, 2:0, 3:0}
        for item in result:
            value[item['status']] += 1
        return value


    def __call__(self):
        return self.template()


# 成績追蹤管理系統-匯出表格(操作類)
class GradeManageExportOperator(GradeManageExportNormal):

    template = ViewPageTemplateFile("template/grade_manage_export_operator.pt")



# 成績追蹤管理系統2
class GradeManage2(GradeManage):
    template = ViewPageTemplateFile("template/grade_manage2.pt")

    def getScore(self, id):
        sqlInstance = SqlObj()
        sqlStr = """SELECT * FROM `grade_mana2` WHERE reg_id={} order by id desc""".format(id)
        result = sqlInstance.execSql(sqlStr)
        if result:
            return [result[0]['status'], result[0]['score_1'], result[0]['score_2'], result[0]['score_3'], result[0]['exam_date']]
        else:
            return [1, 0, 0, 0, None]


    def getStep(self):
        """ 取得梯次數 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()

        uid = context.UID()
        sqlStr = """SELECT DISTINCT `exam_step` FROM grade_mana2 WHERE uid = '{}'""".format(uid)
        result = sqlInstance.execSql(sqlStr)
        return result


    def getExamDate(self, exam_step):
        """ 取得各梯次日期 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()
        sqlStr = """SELECT exam_date FROM grade_mana2 WHERE exam_step = {}""".format(exam_step)
        result = sqlInstance.execSql(sqlStr)
        return result


    def getGrade(self, reg_id):
        """ 取得學員梯次成績 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()
        sqlStr = """SELECT exam_step, status, score_1, score_2, score_3, id FROM grade_mana2 WHERE reg_id = {}""".format(reg_id)
        result = sqlInstance.execSql(sqlStr)

        value = {}
        for item in result:
            value[item['exam_step']] = [item['id'], item['score_1'], item['score_2'], item['score_3'], item['status']]
        return value


    def getStatus(self, reg_id):
        """ 取得學員考試追蹤狀態 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()
        sqlStr = """SELECT MAX(exam_step) as max_exam_step FROM grade_mana2 WHERE reg_id = {}""".format(reg_id)
        result = sqlInstance.execSql(sqlStr)

        max_exam_step = result[0]['max_exam_step']
        sqlStr = """SELECT grade_mana2.id, grade_mana2.status AS value, grade_mana2_status.status AS status
                    FROM grade_mana2, grade_mana2_status
                    WHERE grade_mana2.reg_id = {} and grade_mana2.status = grade_mana2_status.id
                    ORDER BY grade_mana2.id DESC""".format(reg_id)
        return sqlInstance.execSql(sqlStr)


    def getStatusList(self):
        """ 取得狀態列表 """
        request = self.request
        context = self.context
        sqlInstance = SqlObj()
        sqlStr = """SELECT id, status FROM grade_mana2_status"""
        return sqlInstance.execSql(sqlStr)


    def addNewGrade(self):
        """ 新增一梯次 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()

        sqlStr = """SELECT MAX(exam_step) AS max_step FROM grade_mana2 WHERE uid = '{}'""".format(context.UID())
        result = sqlInstance.execSql(sqlStr)
        step = result[0]['max_step'] + 1 if result[0]['max_step'] else 1
        newdate = request.form.get('newdate')

        for item in request.form.keys():
            logger.info('%s: %s' % (item, str(request.form.get(item))))
            if item.startswith('new-'):
                score_1 = request.form.get(item)[0] if request.form.get(item)[0] else 'null'
                score_2 = request.form.get(item)[1] if request.form.get(item)[1] else 'null'
                score_3 = request.form.get(item)[2] if request.form.get(item)[2] else 'null'

                reg_id = item.split('-')[1]
                status = request.form.get('status-%s' % reg_id)

                logger.info('%s: %s, %s, %s' % (item, score_1, score_2, score_3))
                if score_1 == 'null' and score_2 == 'null' and score_3 == 'null':
                    continue

                sqlStr = """INSERT INTO grade_mana2 (exam_date, exam_step, reg_id, score_1, score_2, score_3, uid, status)
                            VALUES ('{}', {}, {}, {}, {}, {}, '{}', {})
                         """.format(newdate, step, reg_id, score_1, score_2, score_3, context.UID(), status)
                sqlInstance.execSql(sqlStr)


    def updateGrade(self):
        """ 更新梯次表 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()

        for item in request.form.keys():
            if item.startswith('old-'):
                step = item.split('-')[1]
                reg_id = item.split('-')[2]
                id = item.split('-')[3]
                exam_date = request.form.get('olddate-%s' % step)
                score_1 = request.form.get(item)[0] if request.form.get(item)[0] else 'null'
                score_2 = request.form.get(item)[1] if request.form.get(item)[1] else 'null'
                score_3 = request.form.get(item)[2] if request.form.get(item)[2] else 'null'
                status = request.form.get('status-%s' % reg_id)

                if not id:
                    if score_1 == 'null' and score_2 == 'null' and score_3 == 'null':
                        continue
                    sqlStr = """INSERT INTO grade_mana2 (exam_date, exam_step, reg_id, score_1, score_2, score_3, uid, status)
                                 VALUES ('{}', {}, {}, {}, {}, {}, '{}', {})
                             """.format(exam_date, step, reg_id, score_1, score_2, score_3, context.UID(), status)
                    sqlInstance.execSql(sqlStr)
                    continue

                if score_1 == 'null' and score_2 == 'null' and score_3 == 'null':
                    sqlStr = """DELETE FROM `grade_mana2`
                                WHERE id={}
                             """.format(id)
                else:
                    sqlStr = """UPDATE `grade_mana2`
                                SET `reg_id`={}, `exam_step`={}, `exam_date`='{}', `score_1`={}, `score_2`={}, `score_3`={}, `uid`='{}', `status`={}
                                WHERE id={}
                             """.format(reg_id, step, exam_date, score_1, score_2, score_3, context.UID(), status, id)
                sqlInstance.execSql(sqlStr)


    def checkAddNew(self):
        request = self.request

        status = [0, 0]
        for item in request.form.items():
            if item[0].startswith('new'):
                logger.info('key:%s, value:%s' % (item[0], item[1]))

                if type(item[1]) == type([]) and (item[1][0] or item[1][1] or item[1][2]):
                    status[0] += 1
                    status[1] += 1

                if type(item[1]) == type(''):
                    status[0] += 1
                    if item[1]:
                        status[1] += 1

        logger.info(str(status))
        if status[1] == 0:
            return True
        elif status[0] == status[1] and status[0] > 1:
            return True
        else:
            return False


# 成績追蹤管理系統2-匯出 學員成績通知單
class GradeManage2Export(BasicBrowserView):

    def __call__(self):

        request = self.request
        context = self.context

        filePath = '/home/andy/cshm/zeocluster/src/cshm.content/src/cshm/content/browser/static/grade_mana2_export.docx'
        doc = DocxTemplate(filePath)

        parameter = {}
        name = request.form.get('name')
        parameter['notify_date'] = DateTime().strftime('%Y年%m月%d日')
        parameter['start_date'] = context.courseStart.strftime('%Y年%m月%d日')
        parameter['end_date'] = context.courseEnd.strftime('%Y年%m月%d日')
        parameter['step'] = context.id
        parameter['exam_date'] = request.form.get('exam_date')
        parameter['siteno'] = request.form.get('siteno')
        parameter['name'] = name
        parameter['pass'] = 'V' if request.form.get('status') == '2' else '□'
        parameter['nopass'] = 'V' if request.form.get('status') in ['3', '4', '5', '6'] else '□'
        parameter['nopass_1'] = 'V' if request.form.get('status') in ['4', '5'] else '□'
        parameter['nopass_2'] = 'V' if request.form.get('status') in ['3', '6'] else '□'
        parameter['score_1'] = request.form.get('score_1')
        parameter['score_2'] = request.form.get('score_2')
        parameter['score_3'] = request.form.get('score_3')

        doc.render(parameter)
        doc.save("/tmp/temp.docx")

        self.request.response.setHeader('Content-Type', 'application/docx')
        self.request.response.setHeader(
            'Content-Disposition',
            'attachment; filename="notify_%s.docx"' % (name)
        )

        with open('/tmp/temp.docx', 'rb') as file:
            docs = file.read()
            return docs


# 成績追蹤管理系統3: 工地主任220小時
class GradeManage3(GradeManage):
    template = ViewPageTemplateFile("template/grade_manage3.pt")

    def getScore(self, id):
        sqlInstance = SqlObj()
        sqlStr = """SELECT * FROM `grade_mana3` WHERE reg_id={} order by id desc""".format(id)
        result = sqlInstance.execSql(sqlStr)
        if result:
            return [result[0]['status'], result[0]['score_1'], result[0]['score_2'], result[0]['exam_date']]
        else:
            return [1, 0, 0, None]


    def getStep(self):
        """ 取得梯次數 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()

#        uid = context.UID()
        sqlStr = """SELECT DISTINCT `exam_step` FROM grade_mana3"""
        result = sqlInstance.execSql(sqlStr)
        return result


    def getExamDate(self, exam_step):
        """ 取得各梯次日期 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()
        sqlStr = """SELECT exam_date FROM grade_mana3 WHERE exam_step = {}""".format(exam_step)
        result = sqlInstance.execSql(sqlStr)
        return result


    def getGrade(self, reg_id):
        """ 取得學員梯次成績 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()
        sqlStr = """SELECT exam_step, status, score_1, score_2, id FROM grade_mana3 WHERE reg_id = {}""".format(reg_id)
        result = sqlInstance.execSql(sqlStr)

        value = {}
        for item in result:
            value[item['exam_step']] = [item['id'], item['score_1'], item['score_2'], item['status']]
        return value


    def getStatus(self, reg_id):
        """ 取得學員考試追蹤狀態 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()
        sqlStr = """SELECT MAX(exam_step) as max_exam_step FROM grade_mana3 WHERE reg_id = {}""".format(reg_id)
        result = sqlInstance.execSql(sqlStr)

        max_exam_step = result[0]['max_exam_step']
        sqlStr = """SELECT grade_mana3.id, grade_mana3.status AS value, grade_mana3_status.status AS status
                    FROM grade_mana3, grade_mana3_status
                    WHERE grade_mana3.reg_id = {} and grade_mana3.status = grade_mana3_status.id
                    ORDER BY grade_mana3.id DESC""".format(reg_id)
        return sqlInstance.execSql(sqlStr)


    def getStatusList(self):
        """ 取得狀態列表 """
        request = self.request
        context = self.context
        sqlInstance = SqlObj()
        sqlStr = """SELECT id, status FROM grade_mana3_status"""
        return sqlInstance.execSql(sqlStr)


    def addNewGrade(self):
        """ 新增一梯次 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()

        sqlStr = """SELECT MAX(exam_step) AS max_step FROM grade_mana3"""
        result = sqlInstance.execSql(sqlStr)
        step = result[0]['max_step'] + 1 if result[0]['max_step'] else 1
        newdate = request.form.get('newdate')

        for item in request.form.keys():
            logger.info('%s: %s' % (item, str(request.form.get(item))))
            if item.startswith('new-'):
                reg_id = item.split('-')[1]
                score_1 = request.form.get(item)[0] if request.form.get(item)[0] else 'null'
                score_2 = request.form.get(item)[1] if request.form.get(item)[1] else 'null'
                study_no = request.form.get('study_no-%s' % reg_id)
                first_exam_sn = request.form.get('first_exam_sn-%s' % reg_id)
                exam_sn = request.form.get('exam_sn-%s' % reg_id)
                status = request.form.get('status-%s' % reg_id)

                sqlStr = """select id, uid from reg_course where id={}""".format(reg_id)
                result = sqlInstance.execSql(sqlStr)
                # 求考區
                course_uid = result[0]['uid']
                exam_area = api.content.find(UID=course_uid)[0].getObject().trainingCenter.to_object.title

                # 求首考日期
                sqlStr = """select first_exam_date from grade_mana3 where reg_id={}""".format(reg_id)
                result = sqlInstance.execSql(sqlStr)
                if result:
                    first_exam_date = sqlInstance.execSql(sqlStr)[0]['first_exam_date']
                else:
                    first_exam_date = newdate

                logger.info('%s: %s, %s' % (item, score_1, score_2))
                if score_1 == 'null' and score_2 == 'null':
                    continue

                sqlStr = """INSERT INTO grade_mana3 (first_exam_date, exam_date, exam_step, reg_id, score_1, score_2,
                                                     uid, status, exam_area, study_no, first_exam_sn, exam_sn)
                            VALUES ('{}', '{}', {}, {}, {}, {}, '{}', {}, '{}', '{}', '{}', '{}')
                         """.format(first_exam_date, newdate, step, reg_id, score_1, score_2, context.UID(),
                                    status, exam_area, study_no, first_exam_sn, exam_sn)
                sqlInstance.execSql(sqlStr)


    def updateGrade(self):
        """ 更新梯次表 """
        request = self.request
        context = self.context

        sqlInstance = SqlObj()

        for item in request.form.keys():
            if item.startswith('old-'):
                step = item.split('-')[1]
                reg_id = item.split('-')[2]
                id = item.split('-')[3]
                exam_date = request.form.get('olddate-%s' % step)
                study_no = request.form.get('study_no-%s' % reg_id)
                first_exam_sn = request.form.get('first_exam_sn-%s' % reg_id)
                exam_sn = request.form.get('exam_sn-%s' % reg_id)

                score_1 = request.form.get(item)[0] if request.form.get(item)[0] else 'null'
                score_2 = request.form.get(item)[1] if request.form.get(item)[1] else 'null'
                status = request.form.get('status-%s' % reg_id)

                sqlStr = """select id, uid from reg_course where id={}""".format(reg_id)
                result = sqlInstance.execSql(sqlStr)

                # 求考區
                course_uid = result[0]['uid']
                exam_area = api.content.find(UID=course_uid)[0].getObject().trainingCenter.to_object.title

                # 求首考日期
                sqlStr = """select first_exam_date from grade_mana3 where reg_id={}""".format(reg_id)
                result = sqlInstance.execSql(sqlStr)
                if result:
                    first_exam_date = sqlInstance.execSql(sqlStr)[0]['first_exam_date']
                else:
                    first_exam_date = exam_date

                if not id:
                    if score_1 == 'null' and score_2 == 'null':
                        continue
                    sqlStr = """INSERT INTO grade_mana3 (first_exam_date, exam_date, exam_step, reg_id, score_1, score_2,
                                                         uid, status, study_no, first_exam_sn, exam_sn, exam_area)
                                 VALUES ('{}', '{}', {}, {}, {}, {}, '{}', {}, '{}', '{}', '{}', '{}')
                             """.format(first_exam_date, exam_date, step, reg_id, score_1, score_2, course_uid,
                                        status, study_no, first_exam_sn, exam_sn, exam_area)
                    sqlInstance.execSql(sqlStr)
                    continue

                if score_1 == 'null' and score_2 == 'null':
                    sqlStr = """DELETE FROM `grade_mana3`
                                WHERE id={}
                             """.format(id)
                else:
                    sqlStr = """UPDATE grade_mana3
                                SET reg_id={}, exam_step={}, exam_date='{}', score_1={}, score_2={}, uid='{}', status={},
                                    first_exam_date='{}', study_no='{}', first_exam_sn='{}', exam_sn='{}'
                                WHERE id={}
                             """.format(reg_id, step, exam_date, score_1, score_2, course_uid, status,
                                        first_exam_date, study_no, first_exam_sn, exam_sn, id)
                sqlInstance.execSql(sqlStr)


    def checkAddNew(self):
        request = self.request

        status = [0, 0]
        for item in request.form.items():
            if item[0].startswith('new'):
                logger.info('key:%s, value:%s' % (item[0], item[1]))

                if type(item[1]) == type([]) and (item[1][0] or item[1][1]):
                    status[0] += 1
                    status[1] += 1

                if type(item[1]) == type(''):
                    status[0] += 1
                    if item[1]:
                        status[1] += 1

        logger.info(str(status))
        if status[1] == 0:
            return True
        elif status[0] == status[1] and status[0] > 1:
            return True
        else:
            return False

    def getUidCourseData(self):
        """抓所有 k005 的資料"""
        portal = api.portal.get()

        courseStart = {
            'query': DateTime()-1096, # 3年含潤年
            'range': 'min',
        }
        courseEnd = {
            'query': DateTime(),
            'range': 'max',
        }

        k005_course = api.content.find(
            context=portal['mana_course']['k005'],
            portal_type='Echelon',
            courseStart=courseStart,
            courseEnd= courseEnd,
        )

        uidStr = ''
        for item in k005_course:
            uidStr += "'%s'," % item.UID
        uidStr = uidStr[:-1]

        # 3年前開始的課，today之前結束
        sqlStr = """SELECT reg_course.*, training_status_code.status, education_code.degree
                    FROM reg_course, training_status_code, education_code
                    WHERE path LIKE 'cshm/mana_course/k005%%'
                      AND reg_course.training_status = training_status_code.id
                      AND reg_course.education_id = education_code.id
                      AND reg_course.isAlt = 0
                      AND reg_course.uid IN ({})
                 """.format(uidStr)
        sqlInstance = SqlObj()
        return sqlInstance.execSql(sqlStr)


    def getGradeInfo(self, reg_id):
        """ 抓考生基本資料, study_no, first_exam_sn, exam_sn """
        sqlStr = """SELECT study_no, first_exam_sn, exam_sn
                    FROM grade_mana3
                    WHERE reg_id = {}
                    ORDER BY id DESC
                 """.format(reg_id)
        sqlInstance = SqlObj()
        result = sqlInstance.execSql(sqlStr)
        if result:
            return result[0]
        else:
            return {'study_no': '', 'first_exam_sn':'', 'exam_sn':''}


    def __call__(self):
        request = self.request

        if not self.checkPermission():
            return self.no_permission_template()

        if request.form.has_key('update'):
            if self.checkAddNew():
                self.addNewGrade()
                self.updateGrade()
                api.portal.show_message(message='Update Grade Complete!', request=request)
                return self.template()
            else:
                return '新增資料輸入有誤，請回上一頁'

        return self.template()


# 工地主任220小時-匯出
class GradeManage3Export(BasicBrowserView):

    template = ViewPageTemplateFile("template/grade_manage3_export.pt")

    def getHighestsocre(self, reg_id):
        sqlStr = """SELECT max(score_1) as score_1, max(score_2) as score_2
                    FROM `grade_mana3`
                    WHERE reg_id={}""".format(reg_id)
        sqlInstance = SqlObj()
        return sqlInstance.execSql(sqlStr)


    def __call__(self):

        request = self.request
        exam_date = request.form.get('exam_date')

        self.result = None
        if exam_date:
            sqlStr = """select *, reg_course.id as reg_course_id, grade_mana3.id as grade_mana3_id
                        from reg_course, grade_mana3
                        where reg_course.id = grade_mana3.reg_id
                            and exam_date = '{}'""".format(exam_date)
            sqlInstance = SqlObj()
            self.result = sqlInstance.execSql(sqlStr)

        return self.template()



# 補課Detail
class MakeUpDetail(BasicBrowserView):

    template = ViewPageTemplateFile("template/make_up_detail.pt")

    def addMakeUpDetail(self, subject, id, origin_reg_id):
        sqlInstance = SqlObj()
        sqlStr = """INSERT INTO `make_up_detail`(`make_up_id`, `origin_reg_id`, `subject_name`) VALUES ({}, {}, '{}')
                 """.format(id, origin_reg_id, subject)
        sqlInstance.execSql(sqlStr)


    def delMakeUpDetail(self, make_up_detail_id):
        sqlInstance = SqlObj()
        sqlStr = """DELETE FROM `make_up_detail` WHERE id={}
                 """.format(make_up_detail_id)
        sqlInstance.execSql(sqlStr)


    def regMakeUp(self, uid, origin_reg_id):
        sqlInstance = SqlObj()
        sqlStr ="""
                INSERT INTO `reg_course`(`cellphone`, `fax`, `tax_no`, `name`, `com_email`,
                    `company_name`, `industry`, `invoice_title`, `company_address`, `priv_email`,
                    `phone`, `birthday`, `address`, `job_title`, `studId`, `uid`, `path`, `isAlt`,
                    `seatNo`, `contactLog`, `training_status`, `isReserve`, `invoice_tax_no`,
                    `education_id`, `edu_school`, `edu_major`, `edu_date`, `city`, `zip`, `home_city`,
                    `home_zip`, `home_address`, `company_city`, `company_zip`, `company_undertaker`,
                    `company_undertaker_job`, `company_undertaker_phone`, `retraining_from`,
                    `on_training`, `is_print`, `reTrainingCat`, `reTrainingCode`, `reTrainingHour`,
                    `training_hour`, `license_unit`, `license_date`, `license_code`, `grade1`, `grade2`, `apply_time`)
                SELECT `cellphone`, `fax`, `tax_no`, `name`, `com_email`, `company_name`, `industry`,
                    `invoice_title`, `company_address`, `priv_email`, `phone`, `birthday`, `address`,
                    `job_title`, `studId`, `uid`, `path`, `isAlt`, `seatNo`, `contactLog`, `training_status`,
                    `isReserve`, `invoice_tax_no`, `education_id`, `edu_school`, `edu_major`, `edu_date`,
                    `city`, `zip`, `home_city`, `home_zip`, `home_address`, `company_city`, `company_zip`,
                    `company_undertaker`, `company_undertaker_job`, `company_undertaker_phone`, `retraining_from`,
                    `on_training`, `is_print`, `reTrainingCat`, `reTrainingCode`, `reTrainingHour`, `training_hour`,
                    `license_unit`, `license_date`, `license_code`, `grade1`, `grade2`, `apply_time`
                FROM `reg_course`
                WHERE id = {};""".format(origin_reg_id)
        sqlInstance.execSql(sqlStr)

        sqlStr = """SELECT MAX(id) AS 'id' FROM reg_course;"""
        result = sqlInstance.execSql(sqlStr)
        insertId = result[0]['id']

        path = api.content.find(UID=uid)[0].getPath()
        sqlStr = """UPDATE `reg_course`
                    SET `uid`='{}',`isAlt`=99,`path`='{}',`seatNo`=null, training_status=3
                    WHERE `id`={}
                 """.format(uid, path, insertId)
        sqlInstance.execSql(sqlStr)
        return


    def getSubjects(self, title):
        # 取得同名課程的所有科目名稱
        portal = api.portal.get()
        brain = api.content.find(Title=title, context=portal['resource']['course_template'])
        for item in brain:
            if safe_unicode(title) == safe_unicode(item.Title):
                course = item.getObject()
                break
        subjects = course.getChildNodes()
        return subjects


    def getSameNameCourse(self, title): # 取得可補課列表
        portal = api.portal.get()
        brain = api.content.find(Title=title, context=portal['mana_course'])

        for item in brain:
            if safe_unicode(title) == safe_unicode(item.Title):
                course = item.getObject()
                break
        return course.getChildNodes() #回傳各期 TODO 要判斷可用期數


    def getCourse(self, uid): #取得當期
        return api.content.find(UID=uid)


    def getMakeUpList(self):
        # 取得補課清單
        request = self.request

        id = request.get('id')
        sqlInstance = SqlObj()
        sqlStr = """SELECT * FROM `make_up_detail` WHERE make_up_id = {}""".format(id)
        return sqlInstance.execSql(sqlStr)


    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        request = self.request
        sqlInstance = SqlObj()

        id = request.form.get('id')

        if request.form.has_key('add_make_up'):
            origin_reg_id = request.form.get('origin_reg_id')
            self.addMakeUpDetail(request.form.get('subject'), id, origin_reg_id)
            request.response.redirect('%s?id=%s' % (request.URL, id))
            return
        elif request.form.has_key('del_make_up'):
            make_up_detail_id = request.form.get('make_up_detail_id')
            self.delMakeUpDetail(make_up_detail_id)
            request.response.redirect('%s?id=%s' % (request.URL, id))
            return
        elif request.form.has_key('reg_make_up'):
            uid = request.form.get('uid')
            origin_reg_id = request.form.get('origin_reg_id')
            self.regMakeUp(uid, origin_reg_id)
            api.portal.show_message(message='Set Make up is Finish!', request=request)
            request.response.redirect('%s?id=%s' % (request.URL, id))
            return

        sqlStr = """SELECT *, wait_make_up.id AS wait_make_up_id, reg_course.id AS reg_course_id
                    FROM `wait_make_up`, `reg_course`
                    WHERE reg_course.id = wait_make_up.origin_reg_id
                        AND wait_make_up.id = {}
                 """.format(id)
        self.result = sqlInstance.execSql(sqlStr)

        return self.template()


# 補課清單
class MakeUpList(BasicBrowserView):

    template = ViewPageTemplateFile("template/make_up_list.pt")

    def get_course(self, uid):
        return api.content.find(UID=uid)

    def get_training_status(self, id):
        sqlStr = """SELECT * FROM `training_status_code` WHERE id = {}""".format(id)
        sqlInstance = SqlObj()
        return sqlInstance.execSql(sqlStr)

    def get_reg_data(self, origin_reg_id):
        sqlStr = """SELECT * FROM `reg_course` WHERE id = {}""".format(origin_reg_id)
        sqlInstance = SqlObj()
        return sqlInstance.execSql(sqlStr)

    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        context = self.context
        request = self.request

        sqlStr = """SELECT * FROM `wait_make_up` WHERE 1"""
        sqlInstance = SqlObj()
        self.result = sqlInstance.execSql(sqlStr)

        return self.template()


# 建立報價單
class CreateQuotation(BasicBrowserView):

    template = ViewPageTemplateFile("template/create_quotation.pt")

    def getQuote(self):
        request = self.request

        id = request.form.get('id')
        sqlInstance = SqlObj()
        sqlStr = """SELECT * FROM `quote_record` WHERE id = {}""".format(id)
        return sqlInstance.execSql(sqlStr)

    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        self.quote = self.getQuote()

        courseUID = self.quote[0]['course_uid']
        self.course = api.content.find(UID=courseUID)

        return self.template()



# 報價管理
class QuotationManage(BasicBrowserView):

    template = ViewPageTemplateFile("template/quotation_manage.pt")

    def getStatus(self, status):
        sqlInstance = SqlObj()
        sqlStr = """SELECT * FROM `quote_status` WHERE id = {}""".format(status)
        return sqlInstance.execSql(sqlStr)

    def updateQuotation(self):
        context = self.context
        request = self.request

        id = request.form.get('id')
        people_number = request.form.get('people_number')
        fax = request.form.get('fax')
        tax_no = request.form.get('tax_no')
        address = request.form.get('address')
        # use courseUID
        course_uid = request.form.get('course_uid')
        course_name = api.content.find(UID=course_uid)[0].Title
        contact = request.form.get('contact')
        cell = request.form.get('cell')
        phone = request.form.get('phone')
        training_center = request.form.get('training_center')
        company_name = request.form.get('company_name')
        email = request.form.get('email')
        capital = request.form.get('capital')
        owner = request.form.get('owner')
        main_product = request.form.get('main_product')
        staff_amount = request.form.get('staff_amount')
        dep_title = request.form.get('dep_title')
        turnover = request.form.get('turnover')
        contactLog = request.form.get('contactLog')
        status = request.form.get('status')

        sqlInstance = SqlObj()
        sqlStr = """UPDATE `quote_record`
                    SET `company_name`='{}', `owner`='{}', `address`='{}', `tax_no`='{}',
                        `main_product`='{}', `capital`='{}', `staff_amount`='{}', `turnover`='{}',
                        `phone`='{}', `fax`='{}', `contact`='{}', `cell`='{}', `email`='{}',
                        `dep_title`='{}', `training_center`='{}', `course_name`='{}', `course_uid`='{}',
                        `people_number`='{}', `contactLog`='{}', `status`={} WHERE id={}
                 """.format(company_name, owner, address, tax_no, main_product, capital, staff_amount, turnover,
                            phone, fax, contact, cell, email, dep_title, training_center, course_name, course_uid,
                            people_number, contactLog, status, id)

        return sqlInstance.execSql(sqlStr)

    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        request = self.request
        portal = api.portal.get()

        if not request.form.has_key('id'):
            request.response.redirect('%s/quotation_list' % portal.absolute_url())
            return

        if request.form.has_key('update'):
            self.updateQuotation()

        id = request.form.get('id')
        sqlInstance = SqlObj()
        sqlStr = """SELECT * FROM `quote_record` WHERE id = {}""".format(id)
        self.result = sqlInstance.execSql(sqlStr)
        return self.template()


# 報價列表
class QuotationList(BasicBrowserView):

    template = ViewPageTemplateFile("template/quotation_list.pt")

    def getStatus(self, status):
        sqlInstance = SqlObj()
        sqlStr = """SELECT * FROM `quote_status` WHERE id = {}""".format(status)
        return sqlInstance.execSql(sqlStr)

    def getList(self):
        sqlInstance = SqlObj()
        sqlStr = """SELECT * FROM `quote_record` WHERE 1 ORDER BY id DESC"""
        return sqlInstance.execSql(sqlStr)

    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        return self.template()


# 報價請求(企業內訓)
class QuoteRequest(BasicBrowserView):

    template = ViewPageTemplateFile("template/quote_request.pt")

    def addQuoteRequestByUID(self, courseUID):
        context = self.context
        request = self.request

        people_number = request.form.get('people-number')
        fax = request.form.get('fax')
        tax_no = request.form.get('tax-no')
        address = request.form.get('address')
        # use courseUID
        course_uid = courseUID
#        import pdb; pdb.set_trace()
        course_name = api.content.find(UID=course_uid)[0].Title
        contact = request.form.get('contact')
        cell = request.form.get('cell')
        phone = request.form.get('phone')
        training_center = request.form.get('training-center')
        company_name = request.form.get('company-name')
        email = request.form.get('email')
        capital = request.form.get('capital')
        owner = request.form.get('owner')
        main_product = request.form.get('main-product')
        staff_amount = request.form.get('staff-amount')
        dep_title = request.form.get('dep_title')
        turnover = request.form.get('turnover')
        created = DateTime().strftime('%Y-%m-%d')

        sqlInstance = SqlObj()
        sqlStr = """INSERT INTO `quote_record`
                        (`company_name`, `owner`, `address`, `tax_no`, `people_number`,
                         `main_product`, `capital`, `staff_amount`, `turnover`, `phone`, `fax`,
                         `contact`, `cell`, `email`, `dep_title`, `training_center`, `course_name`, `course_uid`, `created`)
                    VALUES ('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')
                 """.format(company_name, owner, address, tax_no, people_number, main_product,
                            capital, staff_amount, turnover, phone, fax, contact, cell, email,
                            dep_title, training_center, course_name, course_uid, created)
        return sqlInstance.execSql(sqlStr)


    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        context = self.context
        request = self.request

        if request.form.has_key('course-name'):
            for course in request.form.get('course-name'):
#                import pdb; pdb.set_trace()
                self.addQuoteRequestByUID(course)
#            import pdb; pdb.set_trace()
            api.portal.show_message(message=_('Send Quote Request Finish!'), request=request)

        return self.template()


# 開新班
class CreateClass(BasicBrowserView):

    template = ViewPageTemplateFile("template/create_class.pt")

    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        request = self.request
        portal = api.portal.get()

        courseUID = request.form.get('course')
        if not courseUID:
            return self.template()

        obj = api.content.find(UID=courseUID)[0].getObject() # 課程模版
        courseId = obj.id
        if portal['mana_course'].has_key(courseId):
            container = api.content.create(type='Echelon', title='New Course', container=portal['mana_course'][courseId])
            childNodes = obj.getChildNodes()
            for item in childNodes:
                api.content.copy(source=item, target=container)
        else:
            newCourse = api.content.copy(source=obj, target=portal['mana_course'])
            childNodes = newCourse.getChildNodes()
            container = api.content.create(type='Echelon', title='New Course', container=newCourse)
            for item in childNodes:
                api.content.move(source=item, target=container)


class ClassroomOverview(BasicBrowserView):

    template = ViewPageTemplateFile("template/classroom_overview.pt")

    def getScheduling(self, centerName):
        request = self.request
        date = request.form.get('date', DateTime().strftime('%Y-%m-%d'))

        sqlInstance = SqlObj()
        sqlStr = """SELECT * FROM class_scheduling
                    WHERE start LIKE '{}%%' AND classroom LIKE '{}%%'
                 """.format(date, centerName)
        return sqlInstance.execSql(sqlStr)


    def __call__(self):
        context = self.context
        request = self.request

        return self.template()


# 排課管理
class ClassScheduling(BasicBrowserView):

    template = ViewPageTemplateFile("template/class_scheduling.pt")

    def alreadyScheduling(self, uid):
        sqlStr = """SELECT *
                    FROM class_scheduling
                    WHERE uid = '{}'
                    """.format(uid)
        sqlInstance = SqlObj()
        return sqlInstance.execSql(sqlStr)

    def addClassScheduling(self):
        context = self.context
        request = self.request
        uid = request.form.get('uid')
        date = request.form.get('date')
        start = request.form.get('start')
        end = request.form.get('end')
        teacher_fee = request.form.get('teacher_fee')
        traffic_fee = request.form.get('traffic_fee')
        classroom = request.form.get('classroom')
        if classroom in ['other', 'special']:
            classroomId = classroom
        else:
            classroomId = api.content.find(UID=classroom)[0].id
            # traning center
        trainingCenter = context.trainingCenter.to_object.title
        classroomName = '%s-%s' % (trainingCenter, classroomId)

        if not (date and start and end):
            return _(u'Data lost, please fill all field')

        obj = api.content.find(UID=uid)[0].getObject()
        subject_code = obj.id
        subject_name = obj.title
        teacher = obj.teacher
        if teacher:
            teacher_name = teacher.to_object.title
        hours = obj.hours

        date = self.twYear2Ad(date)

        sqlInstance = SqlObj()
        sqlStr = """INSERT INTO `class_scheduling`(`uid`, `subject_code`, `subject_name`, `start`, `end`, `classroom`, `teacher_fee`, `traffic_fee`)
                    VALUES ('{}', '{}', '{}', '{}', '{}', '{}', {}, {})
                 """.format(uid, subject_code, subject_name,
                            '%s %s:%s' % (date, start[0:2], start[2:]),
                            '%s %s:%s' % (date, end[0:2], end[2:]), classroomName, teacher_fee, traffic_fee)
        try:
            sqlInstance.execSql(sqlStr)
            return _(u'Class Scheduling Success.')
        except:
            return _(u'Fail, Please check data format!!')


    def updateClassScheduling(self):
        sqlInstance = SqlObj()
        context = self.context
        request = self.request
        id = request.form.get('uid')
        date = request.form.get('date')
        start = request.form.get('start')
        end = request.form.get('end')
        teacher_fee = request.form.get('teacher_fee')
        traffic_fee = request.form.get('traffic_fee')

        classroom = request.form.get('classroom')
        if classroom in ['other', 'special']:
            classroomId = classroom
        else:
            classroomId = api.content.find(UID=classroom)[0].id

        # traning center
        trainingCenter = context.trainingCenter.to_object.title
        classroomName = '%s-%s' % (trainingCenter, classroomId)

        if not (date and start and end):
            return _(u'Data lost, please fill all field')

        date = self.twYear2Ad(date)

        sqlStr = """UPDATE `class_scheduling`
                  SET `start`='{}',
                      `end`='{}',
                      `classroom`='{}',
                      `teacher_fee`={},
                      `traffic_fee`={}
                  WHERE id={}""".format('%s %s:%s' % (date, start[0:2], start[2:]),
                                        '%s %s:%s' % (date, end[0:2], end[2:]),
                                        classroomName, teacher_fee, traffic_fee, id)
        try:
            sqlInstance.execSql(sqlStr)
            return _(u'Class Scheduling Success.')
        except:
            return _(u'Fail, Please check data format!!')


    def checkOverTime(self):
        """ 檢查是否有時間重疊 """
        context = self.context
        request = self.request

        id = request.form.get('uid')
        date = request.form.get('date')
        start = request.form.get('start')
        end = request.form.get('end')

        classroom = request.form.get('classroom')
        if classroom in ['other', 'special']:
            classroomId = classroom
        else:
            classroomId = api.content.find(UID=classroom)[0].id

        trainingCenter = context.trainingCenter.to_object.title
        classroomName = '%s-%s' % (trainingCenter, classroomId)

        if not (date and start and end):
            return _(u'Data lost, please fill all field')

        date = self.twYear2Ad(date)

        start = '%s %s:%s' % (date, start[0:2], start[2:])
        end = '%s %s:%s' % (date, end[0:2], end[2:])

        if not id.isdigit():
            id = '0'

        sqlInstance = SqlObj()
        # 查詢教室時間衝突
        if classroom not in ['other', 'special']:
            sqlStr = """SELECT id, uid, start, subject_name
                        FROM `class_scheduling`
                        WHERE classroom='{}'
                          AND ( (end > '{}') AND (start < '{}') )
                          AND id != {}""".format(classroomName, start, end, id)
            result = sqlInstance.execSql(sqlStr)
            if result:
                return sqlInstance.execSql(sqlStr)

        # 同課程時間衝突
        uids = []
        for item in context.getChildNodes():
            uids.append(item.UID())
        uidString = "'%s'" % "', '".join(uids)

        sqlStr = """SELECT id, uid, start, subject_name
                    FROM class_scheduling
                    WHERE uid in ({})
                      AND ( (end > '{}') AND (start < '{}') )
                      AND id != {}""".format(uidString, start, end, id)
        return sqlInstance.execSql(sqlStr)


    def checkFitHours(self):
        """ 檢查是否超時, return 0:剛好符合, 1:不足時數, 2:超過時數 """
        request = self.request
        sqlInstance = SqlObj()

        id = request.form.get('uid')
        date = request.form.get('date')
        start = request.form.get('start')
        end = request.form.get('end')

        if not (date and start and end):
            return _(u'Data lost, please fill all field')

        date = self.twYear2Ad(date)

        # 確認uid
        if id.isdigit():
            sqlStr = """SELECT uid
                        FROM class_scheduling
                        WHERE id = '{}'""".format(id)
            result = sqlInstance.execSql(sqlStr)
            uid = result[0]['uid']
        else:
            uid = id

        # check ignoreSchedule
        if api.content.find(UID=uid)[0].getObject().ignoreSchedule:
            return

        # 開始時間
        s_year, s_month, s_day = date.split('/')
        s_hour = start[0:2]
        s_minute = start[2:]

        # 結束時間
        e_year, e_month, e_day = date.split('/')
        e_hour = end[0:2]
        e_minute = end[2:]

        timedelta = datetime.datetime(
            int(e_year), int(e_month), int(e_day), int(e_hour), int(e_minute)) - datetime.datetime(int(s_year), int(s_month), int(s_day), int(s_hour), int(s_minute)
        )
        timedeltaSec = timedelta.seconds

        # 計算現有時數
        sqlInstance = SqlObj()
        sqlStr = """SELECT id, uid, start, end
                    FROM class_scheduling
                    WHERE uid = '{}'""".format(uid)
        result = sqlInstance.execSql(sqlStr)

        totalHours = 0
        for item in result:
            if int(id) == item['id']:
                continue
            if item['end'] and item['start']:
                totalHours += float( (item['end']-item['start']).seconds ) / 3600

        totalHours += float(timedeltaSec) / 3600

        hours = api.content.find(UID=uid)[0].getObject().hours

        # 檢查是否超時, return 0:剛好符合, 1:不足時數, 2:超過時數
        if totalHours == hours:
            return 0
        elif totalHours < hours:
            return 1
        else:
            return 2


    def deleteRow(self):
        """ 刪除一列 """
        request = self.request
        uid = request.form.get('uid')

        sqlInstance = SqlObj()
        sqlStr = """DELETE FROM `class_scheduling`
                    WHERE id={}""".format(uid)

        sqlInstance.execSql(sqlStr)


    def addRow(self):
        """ 增加一列 """
        request = self.request
        uid = request.form.get('uid')

        obj = api.content.find(UID=uid)[0].getObject()
        subject_code = obj.id
        subject_name = obj.title

        sqlInstance = SqlObj()
        sqlStr = """INSERT INTO `class_scheduling`(`uid`, `subject_code`, `subject_name`)
                    VALUES ('{}', '{}', '{}')""".format(uid, subject_code, subject_name)

        sqlInstance.execSql(sqlStr)


    def checkHours(self):
        context = self.context
        request = self.request
        sqlInstance = SqlObj()

        uids = []
        for item in context.getChildNodes():
#            if item.UID() not in uids:
            uids.append(item.UID())
        uidString = "'%s'" % "', '".join(uids)

        sqlStr = """SELECT *
                    FROM class_scheduling
                    WHERE uid in ({})""".format(uidString)
        sqlResult = sqlInstance.execSql(sqlStr)

        result = {}
        for item in sqlResult:
            if item['uid'] not in result:
                obj = api.content.find(UID=item['uid'])[0].getObject()
                ignoreSchedule = obj.ignoreSchedule
                hours = obj.hours
                if ignoreSchedule:
                    continue
                result[item['uid']] = [hours, 0]

            if item['end'] and item['start']:
                result[item['uid']][1] += float( (item['end']-item['start']).seconds ) / 3600
                result[item['uid']].append(item['id'])
        return json.dumps(result)


    def download(self):
        context = self.context
        request = self.request

        sqlInstance = SqlObj()
        # 檢查是否可以下載
        childNodes = context.getChildNodes()
        for item in childNodes:
            if item.ignoreSchedule:
                continue
            hours = item.hours
            sqlStr = """SELECT start, end
                    FROM class_scheduling
                    WHERE uid = '{}'""".format(item.UID())
            result = sqlInstance.execSql(sqlStr)

            if not result:
                return _(u'Download not Allowed')

            totalHours = 0
            for row in result:
                if row['start'] is None or row['end'] is None:
                    return _(u'Download not Allowed')
                totalHours += float( (row['end']-row['start']).seconds ) / 3600

            if totalHours != hours:
                return _(u'Download not Allowed')

        uids = []
        for item in context.getChildNodes():
            uids.append(item.UID())
        uidString = "'%s'" % "', '".join(uids)

        sqlStr = """SELECT *
                    FROM class_scheduling
                    WHERE uid in ({})""".format(uidString)
        result = sqlInstance.execSql(sqlStr)

        output = StringIO()
        spamwriter = csv.writer(output)
        spamwriter.writerow(['科目編號', '科目名稱', '教室', '講師', '排課', '講酬', '車馬費'])

        for item in result:
            obj = api.content.find(UID=item['uid'])[0].getObject()
            startYear = item['start'].strftime('%Y/%m/%d')
            endYear = item['end'].strftime('%Y/%m/%d')
            spamwriter.writerow([item['subject_code'],
                item['subject_name'], item['classroom'], obj.teacher.to_object.title,
                '%s %s~%s' % (startYear, item['start'].strftime('%H%M'), item['end'].strftime('%H%M')),
                item['teacher_fee'], item['traffic_fee'],
            ])
        contents = output.getvalue()

        request.response.setHeader('Content-Type', 'application/csv')
        request.response.setHeader('Content-Disposition', 'attachment; filename="course_schedule.csv"')

        return output.getvalue()


    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        request = self.request
        uid = request.form.get('uid')

        # download
        if request.form.has_key('download'):
            return self.download()

        # Check HOurs
        if request.form.has_key('checkhours'):
            return self.checkHours()

        # 新增一列
        if request.form.has_key('addrow'):
            self.addRow()
            return

        # 刪除一列
        if request.form.has_key('deleterow'):
            self.deleteRow()
            return

        if request.form.has_key('update_class_scheduling'):

            # 檢查時間是否重疊
            overtimeResult = self.checkOverTime()
            if overtimeResult:
                course = api.content.find(UID=overtimeResult[0]['uid'])[0].getObject().getParentNode().getParentNode().title
                return 'overtime:tr-%s:%s %s(%s/%s)' % (overtimeResult[0]['id'],
                   course, overtimeResult[0]['subject_name'],
                   int(overtimeResult[0]['start'].strftime('%m')),
                   int(overtimeResult[0]['start'].strftime('%d'))
                )

            # 檢查時數
            fitHours = self.checkFitHours()
            if fitHours == 2: # 超時
                return _(u'Over Hours')

            if uid.isdigit():
                result = self.updateClassScheduling()
                return result
            else:
                result = self.addClassScheduling()
                return result

        return self.template()


# 新增教師
class AddTeacher(BasicBrowserView):

    template = ViewPageTemplateFile("template/add_teacher.pt")

    def addTeacher(self):
        request = self.request
        portal = api.portal.get()

        name = safe_unicode(request.form.get('teacher-name'))
        idCardNo = request.form.get('id-card-no')
        birthday = request.form.get('birthday')
        birthday = self.twYear2Ad(birthday)
        year, month, day = birthday.split('/')

        homePhone = request.form.get('contact-phone')
        cellPhone = request.form.get('cell-phone')
        contactAddr = safe_unicode(request.form.get('contact-addr'))

        obj = api.content.create(
            type='Teacher',
            title=name,
            container=portal['resource']['mana_teacher'],
            idCardNo = idCardNo,
            birthday = datetime.datetime(int(year), int(month), int(day)),
            homePhone = homePhone,
            cellPhone = cellPhone,
            contactAddr = contactAddr
        )
        return obj

    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        request = self.request

        if request.form.has_key('add-teacher'):
            try:
                obj = self.addTeacher()
                api.content.transition(obj=obj, transition='publish')
                api.portal.show_message(message=_(u"Add Teacher Success."), request=request, type='info')
            except:
                api.portal.show_message(message=_(u"Add Teacher Fail."), request=request, type='error')
        return self.template()


# 教師管理
class TeacherMana(BasicBrowserView):

    template = ViewPageTemplateFile("template/teacher_mana.pt")

    def updateTeacher(self):
        request = self.request
        portal = api.portal.get()

        uid = request.form.get('uid')
        isMember = request.form.get('isMember')
        cellPhone = request.form.get('cellPhone')
        contactAddr = request.form.get('contactAddr')

        teacher = api.content.find(context=portal, UID=uid)[0].getObject()

        teacher.isMember = True if isMember == 'true' else False
        teacher.cellPhone = cellPhone
        teacher.contactAddr = contactAddr


    def getActivationTeachers(self, activation=True):
        request = self.request
        portal = api.portal.get()
        catalog = portal.portal_catalog

        idCardNo = request.form.get('idCardNo')
        title = request.form.get('title')
        data = {
            'Type':'Teacher',
            'activation':activation,
        }
        if title:
            data['Title'] = title
        if idCardNo:
            data['idCardNo'] = idCardNo

        brain = catalog(data, sort_on='id')
        return brain


    def outputCSV(self):
        request = self.request
        output = StringIO()

        spamwriter = csv.writer(output)
        spamwriter.writerow(['序號', '身份證字號', '姓名'])
        brain = self.getActivationTeachers()
        count = 1
        for item in brain:
            spamwriter.writerow([count, item.idCardNo, item.Title])
            count += 1
        contents = output.getvalue()

        request.response.setHeader('Content-Type', 'application/csv')
        request.response.setHeader('Content-Disposition', 'attachment; filename="teacher_list.csv"')

        return output.getvalue()


    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        request = self.request

        if request.form.has_key('output'):
            return self.outputCSV()

        if request.form.has_key('update_teacher'):
            return self.updateTeacher()

        return self.template()


class NotOnTime(BasicBrowserView):

    template_add = ViewPageTemplateFile("template/not_on_time_add.pt")
    template_list = ViewPageTemplateFile("template/not_on_time_list.pt")

    def getSubjects(self, uid):

        teachSubjects = api.content.find(UID=uid)[0].getObject().teachSubjects
        result = []
        for item in teachSubjects:
            obj = item.to_object
            course = obj.getParentNode()
            uid = obj.UID()
            result.append([uid, str('%s-%s') % (course.title.encode('utf-8'), obj.title.encode('utf-8'))])

        return json.dumps(result)


    def getTeachers(self):
        return api.content.find(portal_type='Teacher')


    def getNotOnTimeStatus(self):
        sqlInstance = SqlObj()
        sqlStr = """SELECT *
                    FROM not_on_time_status
                    WHERE 1"""
        return sqlInstance.execSql(sqlStr)

    def insertRecord(self):
        request = self.request

        current = api.user.get_current()
        logger = current.getProperty('fullname')
        loggerId = current.getProperty('id')
        subject_uid = request.form.get('subject-name')
        subject_path = api.content.find(UID=subject_uid)[0].getPath()
        teacher_uid = request.form.get('teacher-name')
        teacher_name = api.content.find(UID=teacher_uid)[0].Title
        status_code = request.form.get('status_code')
        detail = request.form.get('detail')
        in_scope = request.form.has_key('in_scope')
        date = request.form.get('date')

        date = self.twYear2Ad(date)

        sqlInstance = SqlObj()
        sqlStr = """INSERT INTO `not_on_time_record`(`subject_uid`, `subject_path`, `teacher_uid`, `teacher_name`, `status_code`, `detail`, \
                    `in_scope`, `date`, `logger`, `logger_id`)
                    VALUES ('{}', '{}', '{}', '{}', {}, '{}', {}, '{}', '{}', '{}')
                 """.format(subject_uid, subject_path, teacher_uid, teacher_name, int(status_code), detail, in_scope, date, logger, loggerId)

        return sqlInstance.execSql(sqlStr)

    def getLog(self):
        current = api.user.get_current()
        loggerId = current.getProperty('id')

        sqlInstance = SqlObj()
        sqlStr = """SELECT *
                    FROM `not_on_time_record`
                    WHERE `logger_id` = '{}'
                    ORDER BY `id`
                    DESC
                 """.format(loggerId)

        return sqlInstance.execSql(sqlStr)


    def getNotOnTimeList(self):
        sqlStr = """INSERT INTO `not_on_time_record`(`subject_uid`, `subject_path`, `teacher_uid`, `teacher_name`, `status_code`, `detail`, \
                    `in_scope`, `date`, `logger`)
                    VALUES ('{}', '{}', '{}', '{}', {}, '{}', {}, '{}', '{}')
                 """.format(subject_uid, subject_path, teacher_uid, teacher_name, int(status_code), detail, in_scope, date, logger)

    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        request = self.request

        if request.form.has_key('getSubject'):
            return self.getSubjects(request.form.get('uid'))

        if request.form.has_key('list'):
            return self.template_list()

        if request.form.has_key('_authenticator'):
            try:
                self.insertRecord()
                api.portal.show_message(message=_(u"Insert Not On Time Record Success."), request=request, type='info')
            except:
                api.portal.show_message(message=_(u"Insert Not On Time Fail."), request=request, type='error')
        return self.template_add()



class TeacherView(BrowserView):

    template = ViewPageTemplateFile("template/teacher_view.pt")

    def __call__(self):
        return self.template()


# 期別管理(詳細)
class EchelonDetail(BasicBrowserView):

    template = ViewPageTemplateFile("template/echelon_detail.pt")

    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        return self.template()


# 仿寫自 folderListing,
class CoverListing(BrowserView):

    def __call__(self, batch=False, b_size=20, b_start=0, orphan=0, **kw):
        query = {}
        query.update(kw)

        # 限定範圍: ['News Item', 'Link']
        query['portal_type'] = ['News Item', 'Link']
        query['path'] = {
            'query': '/'.join(self.context.getPhysicalPath()),
            'depth': 1,
        }

        # if we don't have asked explicitly for other sorting, we'll want
        # it by position in parent
        if 'sort_on' not in query:
            query['sort_on'] = 'getObjPositionInParent'

        # Provide batching hints to the catalog
        if batch:
            query['b_start'] = b_start
            query['b_size'] = b_size + orphan

        catalog = getToolByName(self.context, 'portal_catalog')
        results = catalog(query)
        return IContentListing(results)


class CoverView(FolderView):

    def results(self, **kwargs):
        # Extra filter
        portal = api.portal.get()

        kwargs.update(self.request.get('contentFilter', {}))
        if 'object_provides' not in kwargs:  # object_provides is more specific
            kwargs.setdefault('portal_type', self.friendly_types)
        kwargs.setdefault('batch', True)
        kwargs.setdefault('b_size', self.b_size)
        kwargs.setdefault('b_start', self.b_start)

        #
        listing = aq_inner(portal['news']).restrictedTraverse(
            '@@coverListing', None)

        if listing is None:
            return []
        results = listing(**kwargs)

        return results


class ExportEmailCell(BasicBrowserView):
    """ 匯出email, 手機清單 """

    template = ViewPageTemplateFile("template/export_email_cell.pt")

    def __call__(self):

        if not self.checkPermission():
            return self.no_permission_template()

        portal = api.portal.get()
        context = self.context
        request = self.request

        if request.form.has_key('sendmsg'):
            message = request.form.get('message')
            for item in request.form.get('cell').split(','):
                if item:
                    self.sendMessage(item, message)
            api.portal.show_message(message=_(u"Send cell message success."), request=request, type='info')

        if request.form.has_key('sendmail'):
            title = request.form.get('title')
            message = request.form.get('message')
            for item in request.form.get('email').split(';'):
                if item:
                    api.portal.send_email(
                        recipient=item,
                        sender="andyfang51@gmail.com",
                        subject=title,
                        body=message,
                    )
            api.portal.show_message(message=_(u"Send mail success."), request=request, type='info')

        uid = context.UID()

        sqlInstance = SqlObj()
        sqlStr = """SELECT `name`, `priv_email`, `cellphone`
                    FROM reg_course
                    WHERE uid = '{}' and isAlt = 0 and isReserve is null""".format(uid)
        result = sqlInstance.execSql(sqlStr)

        self.rightEmails = []
        self.wrongEmails = []
        self.rightCells = []
        self.wrongCells = []
        for item in result:
            if self.checkEmail(item['priv_email']):
                self.rightEmails.append(item['priv_email'])
            else:
                self.wrongEmails.append([item['name'], item['priv_email']])

            if self.checkCell(item['cellphone']):
                self.rightCells.append(item['cellphone'])
            else:
                self.wrongCells.append([item['name'], item['cellphone']])

        self.email_template = api.portal.get_registry_record('cshm.content.browser.configlet.IOffice.email_template')
        self.msg_template = api.portal.get_registry_record('cshm.content.browser.configlet.IOffice.msg_template')
        return self.template()


class DownloadGroupReg(BrowserView):
    """ 下載團報報名表 """

    def __call__(self):
        portal = api.portal.get()
        context = self.context
        request = self.request
        # TODO template file path, use configlet
        filePath = '/home/andy/cshm/zeocluster/src/cshm.content/src/cshm/content/browser/static/group_reg.xls'

        wb = xlrd.open_workbook(filePath, formatting_info=True)
        wb_copy = copy(wb)

        # write(row, column, text)
        wb_copy.get_sheet(0).write(3,3,"%s-%s" % (context.getParentNode().title, context.title))
        wb_copy.get_sheet(0).write(4,3, request['HTTP_REFERER'])

        wb_copy.save('/tmp/temp.xls')

        self.request.response.setHeader('Content-Type', 'application/xls')
        self.request.response.setHeader(
            'Content-Disposition',
            'attachment; filename="%s-%s-%s.xls"' %
                ("團體報名表", context.getParentNode().title.encode('utf-8'), context.title.encode('utf-8'))
        )
        with open('/tmp/temp.xls', 'rb') as file:
            docs = file.read()
            return docs


class RegCourse(BasicBrowserView):

    template = ViewPageTemplateFile("template/reg_course.pt")

    def makeSqlStr(self, idCard_filename, document_1_filename, document_2_filename, document_3_filename):
        form = self.request.form
#        import pdb; pdb.set_trace()
        uid = self.context.UID()
        path = self.context.virtual_url_path()
        isAlt = self.isAlt()
        is_print =  1 if form.get('check_print') == 'on' else 0
        sqlStr = """INSERT INTO `reg_course`(`cellphone`, `fax`, `tax_no`, `name`, `com_email`, `company_name`,\
                    `invoice_title`, `company_address`, `priv_email`, `phone`, `birthday`, `address`, `job_title`, \
                    `studId`, `uid`, `path`, `isAlt`, `invoice_tax_no`, `education_id`, `city`, `zip`, `industry`, `is_print`, `edu_date`,
                    `edu_school`, `edu_major`, `home_city`, `home_zip`, `home_address`, `company_undertaker`,
                    `company_undertaker_job`, `company_undertaker_phone`, `idCard_filename`, `document_1_filename`, `document_2_filename`, `document_3_filename`)
                    VALUES ('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}',
                    '{}', '{}', '{}', '{}', {}, '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')
                 """.format(form.get('cellphone'), form.get('fax'), form.get('tax_no'), form.get('name'), form.get('com_email'),
            form.get('company_name'), form.get('invoice_title'), form.get('company_address'), form.get('priv_email'), form.get('phone'),
            form.get('birthday'), form.get('address'), form.get('job_title'), form.get('studId'), uid, path, isAlt, form.get('invoice_tax_no'),
            form.get('education_id'), form.get('city'), form.get('zip'), form.get('industry'), is_print, form.get('edu_date'),
            form.get('edu_school'), form.get('edu_major'), form.get('home_city'), form.get('home_zip'), form.get('home_address'),
            form.get('company_undertaker'), form.get('company_undertaker_job'),form.get('company_undertaker_phone'),
            idCard_filename, document_1_filename, document_2_filename, document_3_filename)
        return sqlStr


    def isAlt(self):
        context = self.context
        quota = context.quota

        uid = context.UID()
        sqlInstance = SqlObj()
        sqlStr = """SELECT COUNT(id), MAX(isAlt) FROM reg_course WHERE uid = '{}'""".format(uid)
        result = sqlInstance.execSql(sqlStr)
        nowCount = result[0]['COUNT(id)']
        max_isAlt = result[0]['MAX(isAlt)']
        if max_isAlt is None:
            max_isAlt = 0

        if quota > nowCount:
            return 0
        else:
            return max_isAlt + 1


    def checkFull(self):
        context = self.context
        quota = context.quota

        uid = context.UID()
        sqlInstance = SqlObj()
        sqlStr = """SELECT COUNT(id) FROM reg_course WHERE uid = '{}'""".format(uid)
        result = sqlInstance.execSql(sqlStr)

        studentCount = result[0]['COUNT(id)']
        totalQuota = context.quota + context.altCount

        if studentCount >= context.quota and context.quota > 0:
            context.classStatus = u'fullCanAlt'
        if studentCount >= totalQuota:
            context.classStatus = u'altFull'


    def checkAltFull(self):
        context = self.context
        if context.classStatus == 'altFull':
            return True
        else:
            return False


    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request
        alsoProvides(request, IDisableCSRFProtection)
        form = request.form

        # 檢查是進入報名頁，或者為報名程序
        if not form.get('studId', False):
            return self.template()

        if context.courseStart:
            courseStart = DateTime(context.courseStart.strftime('%Y-%m-%d'))
            birthday = request.form.get('birthday')
            byear, bmonth, bday = birthday.split('-')
            byear = str(int(byear)+18)
            stud_18y = DateTime('%s-%s-%s' % (byear, bmonth, bday))
            if courseStart < stud_18y:
                api.portal.show_message(message=_(u"Not yat 18 years old."), request=request, type='error')
                request.response.redirect(self.portal['training']['courselist'].absolute_url())
                return

        if self.checkAltFull():
            api.portal.show_message(message=_(u"Quota Full include Alternate."), request=request, type='error')
            request.response.redirect(self.portal['training']['courselist'].absolute_url())
            return


        # 證件照存檔
        idCard = request.form.get('idCard').read()
        idCard_filename = 'id_%s%s.%s' % (DateTime().strftime('%Y%m%d%H%M%S'), random.randint(100, 999), request.form.get('idCard').filename.split('.')[-1])
        document_1 = request.form.get('document_1').read()
        document_1_filename = 'doc1_%s%s.%s' % (DateTime().strftime('%Y%m%d%H%M%S'), random.randint(100, 999), request.form.get('document_1').filename.split('.')[-1])
        document_2 = request.form.get('document_2').read()
        document_2_filename = 'doc2_%s%s.%s' % (DateTime().strftime('%Y%m%d%H%M%S'), random.randint(100, 999), request.form.get('document_2').filename.split('.')[-1])
        document_3 = request.form.get('document_3').read()
        document_3_filename = 'doc3_%s%s.%s' % (DateTime().strftime('%Y%m%d%H%M%S'), random.randint(100, 999), request.form.get('document_3').filename.split('.')[-1])

        images_folder_path = api.portal.get_registry_record('cshm.content.browser.configlet.IOffice.images_folder_path')
        with open('%s/%s' % (images_folder_path, idCard_filename), 'w') as file:
            file.write(idCard)
        with open('%s/%s' % (images_folder_path, document_1_filename), 'w') as file:
            file.write(document_1)
        with open('%s/%s' % (images_folder_path, document_2_filename), 'w') as file:
            file.write(document_2)
        with open('%s/%s' % (images_folder_path, document_3_filename), 'w') as file:
            file.write(document_3)

        # 報名寫入 DB
        sqlInstance = SqlObj()
        sqlStr = self.makeSqlStr(idCard_filename, document_1_filename, document_2_filename, document_3_filename)
        sqlInstance.execSql(sqlStr)

        # 寫入舊學員名單
        self.insertOldStudent(request.form)

        # 檢查額滿(含正備取)
        self.checkFull()

        # reindex
        context.reindexObject(idxs=['studentCount', 'classStatus'])

        api.portal.show_message(message=_(u"Registry success."), request=request, type='info')


        # 發送報名成功簡訊
        courseName = context.getParentNode().title
        messageStr = '%s您好，您報名的 %s 課程，已報名成功。' % (form.get('name'), courseName)

        # 客製化內容
        reg_ok_message = api.portal.get_registry_record('cshm.content.browser.configlet.IOffice.reg_ok_message')
        if reg_ok_message:
            messageStr = reg_ok_message.replace('name', form.get('name')).replace('course', courseName)

        self.sendMessage(form.get('cellphone'), messageStr)

        # 發送報名成功email
        api.portal.send_email(
            recipient=form.get('priv_email'),
            sender="andyfang51@gmail.com",
            subject="報名成功通知信-中國勞工安全衛生管理學會",
            body=messageStr,
        )

        if api.user.is_anonymous():
            reg_finish_alert_message = api.portal.get_registry_record('cshm.content.browser.configlet.IOffice.reg_finish_alert_message')
            api.portal.show_message(message=reg_finish_alert_message, request=request, type='info')

        if api.user.is_anonymous():
            request.response.redirect('%s?reg_alert=1&msg=%s' % (self.portal['training']['courselist'].absolute_url(), messageStr))
        else:
            request.response.redirect('%s?reg_alert=1&msg=%s' % (request['ACTUAL_URL'], messageStr))
        return


class GroupRegCourse(RegCourse):

    template = ViewPageTemplateFile("template/group_reg_course.pt")

    def makeSqlStr(self, form):

        uid = self.context.UID()
        path = self.context.virtual_url_path()
# 正確性待確認
        isAlt = self.isAlt()

        form['invoice_title'] = 'TODO'
        form['job_title'] = 'TODO'

#        import pdb; pdb.set_trace()
        sqlStr = """INSERT INTO `reg_course`(`cellphone`, `fax`, `tax_no`, `name`, `com_email`, `company_name`,
               `invoice_title`, `company_address`, `priv_email`, `phone`, `birthday`, `address`, `job_title`,
               `studId`, `uid`, `path`, `isAlt`)
           VALUES ('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')
        """.format(form.get('cellphone'), form.get('fax'), form.get('tax_no'),
            form.get('name').encode('utf-8'), form.get('com_email'), form.get('company_name').encode('utf-8'),
            form.get('invoice_title').encode('utf-8'), form.get('company_address').encode('utf-8'),
            form.get('priv_email'), form.get('phone'), form.get('birthday'), form.get('address').encode('utf-8'),
            form.get('job_title').encode('utf-8'), form.get('studId'), uid, path, isAlt)
        return sqlStr


    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request
        alsoProvides(request, IDisableCSRFProtection)
        form = request.form

        # 檢查是進入報名頁，或者為報名程序
        if not form.get('file_data', False):
            return self.template()

        # 檢查還有名額能報否
        if self.checkAltFull():
            api.portal.show_message(message=_(u"Quota Full include Alternate."), request=request, type='error')
            request.response.redirect(self.portal['training']['courselist'].absolute_url())
            return

        file_data = request.get('file_data')
        file_data = file_data.split(',')[1]
        text = base64.b64decode(file_data)
        """
        try:
            text = text.decode('utf-8')
        except:
            text = text.decode('big5')
        """

        extFilename = request.form.get('file').split('.')[-1]
        fullFilename = '/tmp/group_reg.%s' % extFilename
        with open(fullFilename, 'wb') as file:
            file.write(text)

        wb = xlrd.open_workbook(fullFilename, formatting_info=True)
        st = wb.sheet_by_index(0)

        form = {}
        form['company_name'] = st.cell(1,3).value
        contact = st.cell(2,3).value # 承辦人未處理

        courseName = st.cell(3,3).value # 檢查課程名稱
        uploadURL = st.cell(4,3).value # 檢查上傳網址
        realCourseName = '%s-%s' % (context.getParentNode().title, context.title)
        regURL = request['ACTUAL_URL']

        if courseName != realCourseName or regURL != regURL:
            api.portal.show_message(message=_(u"Course name or upload url wrong, please check once."), request=request, type='error')
            return request.response.redirect(request['ACTUAL_URL'])

        form['company_address'] = st.cell(1,7).value
        form['company_phone'] = st.cell(2,5).value
        form['company_fax'] = st.cell(2,8).value
        form['com_email'] = st.cell(3,8).value

        errorMsg = []
        sucessMsg = []
        for index in range(len(st.col(1))):
            if st.col(1)[index].ctype != 2: # is not number
                continue

            row = st.row(index)
            form['studId'] = row[2].value
            form['name']= row[3].value
            form['birthday'] = row[4].value
            form['phone'] = row[5].value
            form['cellphone'] = row[5].value
            edu = row[6].value
            form['priv_email'] = row[7].value
            form['address'] = row[8].value

            # 所有欄位無資料，表空白行
            if not(form['studId'] or form['name'] or form['birthday'] or form['phone'] or form['cellphone'] or form['priv_email'] or form['address']):
                continue
            # 任一欄位有資料，即有填寫
            if not(form['studId'] and form['name'] and form['birthday'] and form['phone'] and form['cellphone'] and form['priv_email'] and form['address']):
                errorMsg.append('編號 %s 報名失敗:資料缺漏' % int(st.col(1)[index].value))
                continue

            # 檢查身份證號
            if not self.checkIDCard(form['studId']):
                errorMsg.append('編號 %s 報名失敗:身份證號錯誤' % int(st.col(1)[index].value))
                continue

            # 報名寫入 DB
            sqlInstance = SqlObj()
            sqlStr = self.makeSqlStr(form)
            sqlInstance.execSql(sqlStr)
            self.insertOldStudent(form)
            sucessMsg.append('編號%s: %s 報名成功' % (int(st.col(1)[index].value), form['name']))


        # 檢查額滿(含正備取)
        self.checkFull()

        # reindex
        context.reindexObject(idxs=['studentCount', 'classStatus'])

        for msg in sucessMsg:
            api.portal.show_message(message=msg, request=request, type='info')
        for msg in errorMsg:
            api.portal.show_message(message=msg, request=request, type='error')

        request.response.redirect(request['ACTUAL_URL'])
        return


class AdminRegCourse(RegCourse):

    def checkAltFull(self):
        return False


class AdminGroupRegCourse(GroupRegCourse):

    def checkAltFull(self):
        return False


class CourseInfo(BrowserView):

    template = ViewPageTemplateFile('template/course_info.pt')

    def __call__(self):
        self.portal = api.portal.get()
        return self.template()


class CourseListing(BrowserView):

    template = ViewPageTemplateFile("template/course_listing.pt")

    def __call__(self):
        self.portal = api.portal.get()
        request = self.request
        location = request.get('location')

        self.statusList = ['willStart', 'fullCanAlt', 'planed', 'registerFirst']
        date_range = {
            'query': DateTime(DateTime().strftime('%Y-%m-%d')),
            'range': 'min',
        }

        self.echelonBrain = {}
        for status in self.statusList:
            if location:
                self.echelonBrain[status] = api.content.find(
                    context=self.portal['mana_course'],
                    portal_type='Echelon',
                    regDeadline=date_range,
                    classStatus=status,
                    trainingCenterId=location
                )
                for center in api.content.find(context=self.portal['resource']['training_center'], portal_type='TrainingCenter'):
                    if location == center.id:
                        obj = center.getObject()
                        self.trainingCenter = obj
            else:
                self.echelonBrain[status] = api.content.find(
                    context=self.portal['mana_course'],
                    portal_ype='Echelon',
                    regDeadline=date_range,
                    classStatus=status,
                )
                self.trainingCenter = False
        return self.template()


class EchelonListingOperation(BrowserView):

    """ 班別列表 / 辦班作業管理，辦班前/中/後"""

    template = ViewPageTemplateFile("template/echelon_listing_operation.pt")

    def __call__(self):
        self.portal = api.portal.get()
        #TODO 條件尚未明確 
        self.echelonBrain = api.content.find(context=self.portal['mana_course'], Type='Echelon')

        return self.template()


class EchelonListingRegister(BrowserView):

    """ 班別列表 / 報名作業管理 """

    template = ViewPageTemplateFile("template/echelon_listing_register.pt")

    def __call__(self):
        self.portal = api.portal.get()

        self.statusList = ['willStart', 'fullCanAlt', 'planed', 'registerFirst', 'altFull']
        self.echelonBrain = {}
        for status in self.statusList:
            self.echelonBrain[status] = api.content.find(context=self.portal['mana_course'], Type='Echelon', classStatus=status)

        return self.template()


class RegisterDetail(BrowserView):

    """ 報名作業 / 開班程序 """

    template = ViewPageTemplateFile("template/register_detail.pt")

    def licenseDate(self):
        """ 發證日期 """
        context = self.context

        uid = context.UID()
        sqlInstance = SqlObj()
        sqlStr = """SELECT license_date FROM reg_course WHERE uid = '{}'  ORDER BY license_date DESC""".format(uid)
        result = sqlInstance.execSql(sqlStr)
        return result[0]['license_date']

    def regAmount(self):
        """ 報名人數 """
        context = self.context

        uid = context.UID()
        sqlInstance = SqlObj()
        sqlStr = """SELECT COUNT(id) FROM reg_course WHERE uid = '{}' and isAlt = 0""".format(uid)
        result = sqlInstance.execSql(sqlStr)
        return result[0]['COUNT(id)']

    def finishAmount(self):
        """ 結訓人數, training_status in [1,2,3,4] """
        context = self.context

        today = datetime.date.today()
        if not context.courseEnd or today < context.courseEnd:
            return ''

        uid = context.UID()
        sqlInstance = SqlObj()
        sqlStr = """SELECT COUNT(id) FROM reg_course WHERE uid = '{}' and training_status in (1, 2, 3, 4)""".format(uid)
        result = sqlInstance.execSql(sqlStr)
        return result[0]['COUNT(id)']

    def __call__(self):
        self.portal = api.portal.get()

        return self.template()


class StudentsList(BasicBrowserView):

    """ 報名作業 / 學生列表 """

    template = ViewPageTemplateFile("template/students_list.pt")

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context

        uid = context.UID()
        sqlInstance = SqlObj()
        # 正取名單
        sqlStr = """SELECT reg_course.*, status
                    FROM reg_course, training_status_code
                    WHERE uid = '{}' and
                          isAlt = 0 and
                          reg_course.training_status = training_status_code.id
                    ORDER BY isnull(isReserve),isReserve, isnull(seatNo), seatNo""".format(uid)
        self.admit = sqlInstance.execSql(sqlStr)

        self.admitForJS = []
        numberCount = 0
        for item in self.admit:
            temp = dict(item)
            numberCount += 1
            temp['id'] = int(item['id'])
            temp['apply_time'] = temp['apply_time'].strftime('%Y/%m/%d %H:%M:%S')
            if temp['birthday']:
                temp['birthday'] = temp['birthday'].strftime('%Y/%m/%d')
            if temp['license_date']:
                temp['license_date'] = temp['license_date'].strftime('%Y/%m/%d')
            if temp['edu_date']:
                temp['edu_date'] = temp['edu_date'].strftime('%Y/%m/%d')

            self.admitForJS.append(temp)
        self.admitForJS = json.dumps(self.admitForJS)

        # 備取名單
        sqlStr = """SELECT * FROM reg_course WHERE uid = '{}' and isAlt > 0 and on_training != 3 ORDER BY isAlt""".format(uid)
        self.waiting = sqlInstance.execSql(sqlStr)

        # 取得受訓狀態代碼
        self.trainingStatus = self.getTrainingStatusCode()

        # 計算繳費人數
        sqlStr = """SELECT user_id,money FROM receipt WHERE uid = '{}'""".format(uid)
        receipt = sqlInstance.execSql(sqlStr)
        count = 0
        totalMoney = 0
        for item in receipt:
            obj = dict(item)
            # 分割字會多一個空字串
            count += (len(obj['user_id'].split(',')) - 1)
            totalMoney += obj['money']
        self.count = count
        self.totalMoney = totalMoney
        self.numberCount = numberCount
        return self.template()


class UpdateContactLog(BrowserView):

    """ 更新聯絡記錄 contactLog """

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request

        id = request.form.get('id')
        contactLog = request.form.get('contactLog')

        sqlInstance = SqlObj()
        sqlStr = """update reg_course set contactLog = '{}' where id = {}""".format(contactLog, id)
        sqlInstance.execSql(sqlStr)
        return


class AdmitBatchNumbering(BrowserView):

    """ 正取學員批次座位編碼 """

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request

        uid = context.UID()
        sqlInstance = SqlObj()

        sqlStr = """SELECT MAX(seatNo) FROM `reg_course` WHERE uid = '{}'""".format(uid)
        result = sqlInstance.execSql(sqlStr)
        seatNo = (int(result[0]['MAX(seatNo)']) + 1) if result[0]['MAX(seatNo)'] else 1

        # 只能執行批次一次
#        if seatNo > 1:
#            return '1'

        sqlStr = """SELECT id, seatNo
                    FROM `reg_course`
                    WHERE uid = '{}' and isAlt = 0 and isReserve is null and seatNo is null""".format(uid)
        result = sqlInstance.execSql(sqlStr)

        for item in result:
            id = item['id']
            sqlStr = """update `reg_course` set seatNo = {} where id = {}""".format(seatNo, id)
            sqlInstance.execSql(sqlStr)
            seatNo += 1


class AllTransToAdmit(BrowserView):

    """ 備取全部轉正取 """

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request

        uid = context.UID()
        sqlInstance = SqlObj()
        sqlStr = """SELECT id FROM `reg_course` WHERE uid = '{}' and isAlt = 0""".format(uid)
        result = sqlInstance.execSql(sqlStr)
        if len(result):
            return 1 # 若正取名單有人，就不能全轉，回傳1
        else:
            sqlStr = """SELECT id FROM `reg_course` WHERE uid = '{}'""".format(uid)
            result = sqlInstance.execSql(sqlStr)

            for item in result:
                id = item['id']
                sqlStr = """update `reg_course` set isAlt = 0 where id = {}""".format(id)
                sqlInstance.execSql(sqlStr)


class WaitingTransToAdmit(BrowserView):

    """ 備取轉正取(單筆) """

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request

        id = request.form.get('id')
        sqlInstance = SqlObj()
        sqlStr = """update `reg_course` set isAlt = 0, on_training = 1 where id = {}""".format(id)
        sqlInstance.execSql(sqlStr)


class UpdateSeatNo(BrowserView):

    """ 更新座位編號 """

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request

        sqlStr = ''
        for item in request.form.keys():
            if item.startswith('seat-'):
                id = item.split('-')[-1]
                sqlStr += "update `reg_course` set seatNo = {} where id = {};".format(request.form['seat-%s' % id], id)
            if item.startswith('status-'):
                id = item.split('-')[-1]
                sqlStr += "update `reg_course` set training_status = {} where id = {};".format(request.form['status-%s' % id], id)

        sqlInstance = SqlObj()
        sqlInstance.execSql(sqlStr)


class ReserveSeat(BrowserView):

    """ 預約保留座位 """

    def __call__(self):
        portal = api.portal.get()
        context = self.context
        request = self.request

        alsoProvides(request, IDisableCSRFProtection)

        quota = request.form.get('quota')
        company_name = request.form.get('company_name')
        uid = context.UID()
        path = context.virtual_url_path()

        sqlStr = "INSERT INTO `reg_course`(`company_name`, `uid`, `path`, `isAlt`, `isReserve`) VALUES "
        for i in range(int(quota)):
            sqlStr += "('%s', '%s', '%s', 0, 1), " % (company_name, uid, path)
        sqlStr = sqlStr[:-2]

        sqlInstance = SqlObj()
        sqlInstance.execSql(sqlStr)

        api.portal.show_message(message=_(u"Reserve Seat is OK."), request=request, type='info')
        request.response.redirect('%s/@@students_list' % context.absolute_url())
        return


class ExportToDownloadSite(BrowserView):

    """ 匯出帳號到下載專區 """

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request

        url = request.form.get('url')
        uid = context.UID()
        sqlInstance = SqlObj()
        sqlStr = """select studId, seatNo from `reg_course` where uid = '{}' and isAlt = 0""".format(uid)
        result = sqlInstance.execSql(sqlStr)
        data = ''
        for item in result:
            data += '%s:%s\n' % (item['studId'], item['seatNo'])
        data = data[:-1]

        req = requests.patch(url,
            headers={'Accept': 'application/json', 'Content-Type': 'application/json'},
            json={'description': data},
            auth=('editor', 'editor')
        )
        if req.ok:
            return 1
        else:
            return 0


class TeacherSelector(BrowserView):

    """ 教師遴選作業 """

    template = ViewPageTemplateFile("template/teacher_selector.pt")

    def __call__(self):
        portal = api.portal.get()
        context = self.context
        request = self.request

        uid = request.form.get('uid', None)
        if not uid:
            return self.template()

        brain = api.content.find(UID=uid)
        self.refObj = brain[0].getObject()
        return self.template()


class TeacherAppointment(BrowserView):

    """ 教師聘書 套表作業(docx) """

    template = ViewPageTemplateFile("template/teacher_appointment.pt")

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request
        childNodes = context.getChildNodes()

        teacher_name = request.get('teacher_name')
        type = request.get('type')
        if teacher_name and type:
            if type == 'inner':
                filePath = '/home/andy/cshm/zeocluster/src/cshm.content/src/cshm/content/browser/static/teacher_appointment_inner.docx'
            elif type == 'organization':
                filePath = '/home/andy/cshm/zeocluster/src/cshm.content/src/cshm/content/browser/static/teacher_appointment_organization.docx'
            elif type == 'not_organization':
                filePath = '/home/andy/cshm/zeocluster/src/cshm.content/src/cshm/content/browser/static/teacher_appointment_not_organization.docx'
            doc = DocxTemplate(filePath)
            data = []

            for child in childNodes:
                teacher = child.teacher.to_object.title
                if teacher == teacher_name:
                    subject = child.title
                    hours = child.hours
                    startDate = child.startDateTime.strftime('%m月%d日')
                    startTime = '%s至%s' %(child.startDateTime.strftime('%H:%M'),
                                (child.startDateTime + datetime.timedelta(hours=hours)).strftime('%H:%M'))
                    data.append([subject, startDate, startTime, hours])

            parameter = {}
            courseStart = context.courseStart
            trainingCenter = context.trainingCenter.to_object.title

            parameter['title'] = context.getParentNode().Title()
            parameter['courseStart'] = courseStart.strftime('%m月%d日')
            parameter['trainingCenter'] = trainingCenter
            parameter['data'] = data
            parameter['now'] = datetime.datetime.now().strftime('%Y年%m月%d日')

            doc.render(parameter)
            doc.save("/tmp/temp.docx")

            self.request.response.setHeader('Content-Type', 'application/docx')
            self.request.response.setHeader(
                'Content-Disposition',
                'attachment; filename="%s-%s.docx"' % ("教師聘書", teacher_name)
            )
            with open('/tmp/temp.docx', 'rb') as file:
                docs = file.read()
                return docs
        else:
            teacherList = []
            for child in childNodes:
                teacher = child.teacher.to_object.title
                if teacher not in teacherList:
                    teacherList.append(teacher)
            self.teacherList = teacherList

            return self.template()


class QueryCompany(BrowserView):

    """ 即時搜尋公司名稱 """

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request

        taxNo = request.form.get('tax_no')
        sqlStr = "SELECT * FROM `company_basic_info` WHERE `tax_no` = '%s'" % taxNo
        sqlInstance = SqlObj()
        result = sqlInstance.execSql(sqlStr)
        if result:
            return '%s,%s' % (result[0]['company_name'], result[0]['company_address'])


class QueryTaxNo(BrowserView):

    """ 即時搜尋統一編號 """

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request

        companyName = request.form.get('company_name')
        sqlStr = "SELECT * FROM `company_basic_info` WHERE `company_name` = '%s'" % companyName
        sqlInstance = SqlObj()
        result = sqlInstance.execSql(sqlStr)
        if result:
            return '%s,%s' % (result[0]['tax_no'], result[0]['company_address'])


class UpdateStudentReg(BasicBrowserView):

    """ 修改學員個人詳細報名資訊 """
    template = ViewPageTemplateFile("template/update_student_reg.pt")
    regForm = ViewPageTemplateFile("template/reg_course_print_form.pt")

    def add_make_up(self, id, training_status):

        # 檢查狀態,若不為 待補課(3)，return
        if training_status != '3':
            return

        # 檢查是否已存在
        sqlInstance = SqlObj()

        sqlStr = """SELECT id
                    FROM wait_make_up
                    WHERE origin_reg_id = {}
                 """.format(id)
        result = sqlInstance.execSql(sqlStr)

        # 若已有補課資料，return
        if result:
            return

        # 新增 make_up
        sqlStr = """SELECT uid
                    FROM reg_course
                    WHERE id = {}
                 """.format(id)
        result = sqlInstance.execSql(sqlStr)

        origin_reg_id = id
        begin_date = DateTime().strftime('%Y-%m-%d')
        status = 3
        make_up_course_uid = result[0]['uid']

        sqlStr = """INSERT INTO `wait_make_up`(`origin_reg_id`, `begin_date`, `status`, `make_up_course_uid`)
                    VALUES ({}, '{}', {}, '{}')
                 """.format(origin_reg_id, begin_date, status, make_up_course_uid)
        sqlInstance.execSql(sqlStr)
        return


    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request

        id = request.form.get('id')
        self.regCourse = self.getRegCourse(id)

        # 套印報名表
        if request.form.has_key('print_form'):
            return self.regForm()

        if not request.form.has_key('updateinfo'):
            return self.template()

        # 檢查受訓狀態，若改為待補課，多一筆
        training_status = request.form.get('training_status')
        self.add_make_up(id, training_status)

        form = request.form
        sqlInstance = SqlObj()
        sqlStr = "UPDATE `reg_course` \
                  SET `cellphone`='%s', `fax`='%s',\
                      `tax_no`='%s',`com_email`='%s', \
                      `company_name`='%s',`invoice_title`='%s', \
                      `company_address`='%s',`priv_email`='%s',`industry`='%s',\
                      `phone`='%s',`birthday`='%s',`address`='%s', \
                      `job_title`='%s', \
                      `training_status`=%s, \
                      `invoice_tax_no`='%s', \
                      `education_id`=%s,`edu_school`='%s', \
                      `city`='%s',`zip`='%s',`company_city`='%s',`company_zip`='%s',`reTrainingCat`='%s',`reTrainingCode`='%s',\
                      `reTrainingHour`=%s,training_hour=%s,license_unit='%s',license_date='%s',license_code='%s' \
                  WHERE id=%s" % \
                  (form.get('cellphone'), form.get('fax'), form.get('tax_no'), form.get('com_email'), form.get('company_name'),
                   form.get('invoice_title'), form.get('company_address'), form.get('priv_email'), form.get('industry'),
                   form.get('phone'), form.get('birthday'), form.get('address'), form.get('job_title'), form.get('training_status'),
                   form.get('invoice_tax_no'), form.get('education_id'), form.get('edu_school'), form.get('city'), form.get('zip'),
                   form.get('company_city'), form.get('company_zip'), form.get('reTrainingCat'), form.get('reTrainingCode'),
                   form.get('reTrainingHour'), form.get('training_hour'), form.get('license_unit'), form.get('license_date'),
                   form.get('license_code'), form.get('id'))
        sqlInstance.execSql(sqlStr)

        if form.get('contactLog'):
            self.updateContactLog(self.regCourse, form.get('contactLog'))

        request.response.redirect('%s?id=%s' % (request['ACTUAL_URL'], request.form.get('id')))
        return self.template()


class SeatTable(BasicBrowserView):

    """ 套印座位表 """
    template = ViewPageTemplateFile("template/seat_table.pt")

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request

        uid = context.UID()
        sqlStr = "SELECT `seatNo`, `name` FROM `reg_course` WHERE isAlt = 0 and isReserve is null and uid = '%s' ORDER BY seatNo" % uid
        sqlInstance = SqlObj()
        self.result = sqlInstance.execSql(sqlStr)
        return self.template()


class BatchUpdateBeforeTraining(BasicBrowserView):

    """ 批次更新(訓前) """
    template = ViewPageTemplateFile("template/batch_update_before_training.pt")

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request
        sqlInstance = SqlObj()

        # 沒有update==1,出列表
        if request.form.get('update') != '1':
            uid = context.UID()
            sqlStr = "SELECT * FROM `reg_course` WHERE isAlt = 0 and isReserve is null and uid = '%s' ORDER BY seatNo" % uid
            self.result = sqlInstance.execSql(sqlStr)
            return self.template()
        #import pdb; pdb.set_trace()
        id = request.form.get('id')
        form = request.form
        on_training = int(form.get('on_training'))
        seatNo = None
        isAlt = 0
        if on_training == 3:
            seatNo = 'NULL'
            isAlt = 100 # 100??


        sqlStr = "UPDATE reg_course \
                  SET phone = '%s', \
                      on_training = '%s' " % (form.get('phone'), form.get('on_training'))
        if seatNo:
            sqlStr += ", seatNo = %s, isAlt = %s " % (seatNo, isAlt)

        sqlStr += "WHERE id = %s" % id
        sqlInstance.execSql(sqlStr)

        if form.get('contactLog'):
            regCourse = self.getRegCourse(id)
            self.updateContactLog(regCourse, form.get('contactLog'))


class BatchUpdateOnTraining(BasicBrowserView):

    """ 批次更新(訓中) """
    template = ViewPageTemplateFile("template/batch_update_on_training.pt")

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request
        sqlInstance = SqlObj()

        # 沒有update==1,出列表
        if request.form.get('update') != '1':
            uid = context.UID()
            sqlStr = "SELECT * FROM `reg_course` WHERE isAlt = 0 and isReserve is null and uid = '%s' ORDER BY seatNo" % uid
            self.result = sqlInstance.execSql(sqlStr)
            return self.template()

        id = request.form.get('id')
        form = request.form
        sqlStr = "UPDATE reg_course \
                  SET phone = '%s', \
                      city = '%s', \
                      zip = '%s', \
                      address = '%s' \
                  WHERE id = %s" % (form.get('phone'), form.get('city'), form.get('zip'), form.get('address'), id)
        sqlInstance.execSql(sqlStr)

        if form.get('contactLog'):
            regCourse = self.getRegCourse(id)
            self.updateContactLog(regCourse, form.get('contactLog'))


class BatchUpdateAfterTraining(BasicBrowserView):

    """ 批次更新(訓後) """
    template = ViewPageTemplateFile("template/batch_update_after_training.pt")

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request
        sqlInstance = SqlObj()

        # 沒有update==1,出列表
        if request.form.get('update') != '1':
            uid = context.UID()
            sqlStr = "SELECT * FROM `reg_course` WHERE isAlt = 0 and isReserve is null and uid = '%s' ORDER BY seatNo" % uid
            self.result = sqlInstance.execSql(sqlStr)
            return self.template()

        id = request.form.get('id')
        form = request.form
        sqlStr = "UPDATE reg_course \
                  SET phone = '%s', \
                      city = '%s', \
                      zip = '%s', \
                      address = '%s' \
                  WHERE id = %s" % (form.get('phone'), form.get('city'), form.get('zip'), form.get('address'), id)
        sqlInstance.execSql(sqlStr)

        if form.get('contactLog'):
            regCourse = self.getRegCourse(id)
            self.updateContactLog(regCourse, form.get('contactLog'))


class DelReserve(BrowserView):

    """ 刪除預約 """

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request

        id = request.form.get('id')

        if not id:
            return

        sqlStr = "DELETE FROM `reg_course` WHERE id = %s and isReserve = 1" % id
        sqlInstance = SqlObj()
        sqlInstance.execSql(sqlStr)


class DebugView(BrowserView):

    def __call__(self):
        portal = api.portal.get()
        context = self.context
        request = self.request

        alsoProvides(request, IDisableCSRFProtection)
        import pdb; pdb.set_trace()


class ListPrint(BrowserView):
    """選擇畫面"""
    template = ViewPageTemplateFile('template/list_print_view.pt')
    """顯示表格畫面"""
    list_print_template = ViewPageTemplateFile('template/list_print_template.pt')
    def __call__(self):
        context = self.context

        uid = context.UID()
        request = self.request
        execSql = SqlObj()
        execStr = """SELECT COUNT(reg_course.id) FROM reg_course, training_status_code WHERE uid = '%s' and 
                  isAlt = 0 and reg_course.training_status = training_status_code.id and reg_course.seatNo is not null""" %uid
        self.total_number = execSql.execSql(execStr)[0][0]
        # page為第幾頁，gap為一頁幾個
        page = request.get('page', '')
        gap = request.get('gap', '')

        self.childNodes = context.getChildNodes()

        if page and gap:
            page = int(page)
            gap = int(gap)
            sqlStr = """SELECT reg_course.name, reg_course.seatNo, reg_course.training_status, reg_course.training_hour
                        FROM reg_course, training_status_code
                        WHERE uid = '{}' and
                              isAlt = 0 and
                              reg_course.training_status = training_status_code.id and
                              reg_course.seatNo is not null
                        ORDER BY seatNo""".format(uid)
            admit = execSql.execSql(sqlStr)
            data = []
            data.append(admit[(page - 1) * gap:page * gap])

            courseList = []
            for course in context.getChildNodes():
                date = course.startDateTime.strftime('%m%d')
                title = course.title
                courseList.append([date, title])

            self.courseList = courseList
            self.data = data
            return self.list_print_template()
        else:
            return self.template()


class HasExportCount(BrowserView):
    template = ViewPageTemplateFile('template/has_export_count.pt')
    def __call__(self):
        request = self.request
        context = self.context
        execSql = SqlObj()
        data = {}
        for course in context.getChildNodes():
            courseName = course.title
            uidList = []
            for echelon in course.getChildNodes():
                echelonUID = echelon.UID()
                echelonID = echelon.id
                uidList.append(echelonUID)

            execStr = """SELECT COUNT(id), uid FROM reg_course WHERE uid in {} AND isAlt != 0 AND 
                         on_training = 3 GROUP BY uid""".format(tuple(uidList))
            result = execSql.execSql(execStr)
            for item in result:
                numbers = item[0]
                uid = str(item[1])
                duringTime = api.content.get(UID=uid).duringTime
                if data.has_key(courseName):
                    if data[courseName].has_key(duringTime):
                        data[courseName][duringTime][0] += numbers
                        uidList = json.loads(data[courseName][duringTime][1])
                        if uid not in uidList:
                            uidList.append(uid)
                            data[courseName][duringTime][1]= json.dumps(uidList)
                    else:
                        data[courseName][duringTime] = [numbers, json.dumps([uid])]
                else:
                    data[courseName] = {duringTime: [numbers, json.dumps([uid])]}
        self.data = data
        return self.template()


class HasExportView(BrowserView):
    template = ViewPageTemplateFile('template/has_export_view.pt')
    def __call__(self):
        request = self.request
        context = self.context
        execSql = SqlObj()
        data = {}
        temp = []
        uidList = json.loads(request.get('uidList'))

        if len(uidList) == 1:
            uidList.append('zzz')

        for uid in uidList:
            temp.append(str(uid))

        sqlStr = """SELECT * FROM reg_course WHERE on_training = 3 and uid in {}""".format(tuple(temp))
        result = execSql.execSql(sqlStr)
        duringTime = api.content.get(UID=result[0][16]).duringTime
        if duringTime == 'inDay':
            self.time = '日間'
        elif duringTime == 'inEvening':
            self.time = '夜間'
        elif duringTime == 'inWeekend':
            self.time = '假日'
        elif duringTime == 'inWeekendEvening':
            self.time = '假日夜間'
        elif duringTime == 'complex':
            self.time = '綜合'
        elif duringTime == 'phone':
            self.time = '電話'

        self.course_name = api.content.get(UID=result[0][16]).getParentNode().title
        self.result = result
        return self.template()


class RegisterPrint(BasicBrowserView):
    # 橫式
    template_horizontal = ViewPageTemplateFile("template/register_print_horizontal.pt")
    def __call__(self):
        request = self.request
        context = self.context
        uid = context.UID()

        self.result = self.getUidCourseData(uid)

        return self.template_horizontal()


class GradeInput(BasicBrowserView):
    template = ViewPageTemplateFile('template/grade_input.pt')
    def __call__(self):
        request = self.request
        context = self.context
        data = request.get('data')
        if data:
            execSql = SqlObj()
            for k,v in json.loads(data).items():
                sqlStr = """UPDATE reg_course SET grade1 = {}, grade2 = {} WHERE id = {}""".format(v['grade1'], v['grade2'], k)
                execSql.execSql(sqlStr)
        uid = context.UID()
        self.result = self.getUidCourseData(uid)

        return self.template()


class TestView(BrowserView):

    def __call__(self):
        request = self.request
        alsoProvides(request, IDisableCSRFProtection)
#        import pdb; pdb.set_trace()
