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

logger = logging.getLogger("cshm.content")


class BasicBrowserView(BrowserView):
    """ 通用method """

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
        portal = api.portal.get()
        context = self.context
        request = self.request

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

    def makeSqlStr(self):
        form = self.request.form

        uid = self.context.UID()
        path = self.context.virtual_url_path()
        isAlt = self.isAlt()

        sqlStr = """INSERT INTO `reg_course`(`cellphone`, `fax`, `tax_no`, `name`, `com_email`, `company_name`,\
                    `invoice_title`, `company_address`, `priv_email`, `phone`, `birthday`, `address`, `job_title`, \
                    `studId`, `uid`, `path`, `isAlt`, `invoice_tax_no`, `education_id`, `city`, `zip`, `industry`)
                    VALUES ('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{})
        """.format(form.get('cellphone'), form.get('fax'), form.get('tax_no'), form.get('name'), form.get('com_email'),
            form.get('company_name'), form.get('invoice_title'), form.get('company_address'), form.get('priv_email'), form.get('phone'),
            form.get('birthday'), form.get('address'), form.get('job_title'), form.get('studId'), uid, path, isAlt, form.get('invoice_tax_no'),
            form.get('education_id'), form.get('city'), form.get('zip'), form.get('industry'))
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

        if self.checkAltFull():
            api.portal.show_message(message=_(u"Quota Full include Alternate."), request=request, type='error')
            request.response.redirect(self.portal['training']['courselist'].absolute_url())
            return

        # 報名寫入 DB
        sqlInstance = SqlObj()
        sqlStr = self.makeSqlStr()
        sqlInstance.execSql(sqlStr)

        # 寫入舊學員名單
        self.insertOldStudent(request.form)

        # 檢查額滿(含正備取)
        self.checkFull()

        # reindex
        context.reindexObject(idxs=['studentCount', 'classStatus'])

        api.portal.show_message(message=_(u"Registry success."), request=request, type='info')

        if api.user.is_anonymous():
            api.portal.show_message(message=u"請以傳真:02-22222222或Email: service@cshm.org.tw 傳送XX/XX等證件影本(文字待確認).", request=request, type='info')

        if api.user.is_anonymous():
            request.response.redirect(self.portal['training']['courselist'].absolute_url())
        else:
            request.response.redirect(request['ACTUAL_URL'])
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

        self.statusList = ['willStart', 'fullCanAlt', 'planed', 'registerFirst']
        date_range = {
            'query': DateTime(DateTime().strftime('%Y-%m-%d')),
            'range': 'min',
        }
        self.echelonBrain = {}
        for status in self.statusList:
            self.echelonBrain[status] = api.content.find(
                context=self.portal['mana_course'],
                Type='Echelon',
                regDeadline=date_range,
                classStatus=status
            )

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

    def regAmount(self):
        context = self.context

        uid = context.UID()
        sqlInstance = SqlObj()
        sqlStr = """SELECT COUNT(id) FROM reg_course WHERE uid = '{}' and isAlt = 0""".format(uid)
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
        for item in self.admit:
            temp = dict(item)
            temp['id'] = int(item['id'])
            temp['apply_time'] = temp['apply_time'].strftime('%Y/%m/%d %H:%M:%S')
            if temp['birthday']:
                temp['birthday'] = temp['birthday'].strftime('%Y/%m/%d')
            self.admitForJS.append(temp)
        self.admitForJS = json.dumps(self.admitForJS)

        # 備取名單
        sqlStr = """SELECT * FROM reg_course WHERE uid = '{}' and isAlt > 0 and on_training != 3 ORDER BY isAlt""".format(uid)
        self.waiting = sqlInstance.execSql(sqlStr)

        # 取得受訓狀態代碼
        self.trainingStatus = self.getTrainingStatusCode()
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
        # TODO template file path, use configlet
        filePath = '/home/andy/cshm/zeocluster/src/cshm.content/src/cshm/content/browser/static/teacher_appointment.docx'

        # if no parameter, return template
        if not request.form.get('appointment_no', None):
            return self.template()

        doc = DocxTemplate(filePath)

        parameter = {}
        for key in request.form:
            parameter[key] = safe_unicode(request.form.get(key))

        parameter['memo'] = Listing(parameter['memo'])
        doc.render(parameter)
        doc.save("/tmp/temp.docx")

        self.request.response.setHeader('Content-Type', 'application/docx')
        self.request.response.setHeader(
            'Content-Disposition',
            'attachment; filename="%s-%s.docx"' % ("教師聘書", parameter.get("teacher_name").encode('utf-8'))
        )
        with open('/tmp/temp.docx', 'rb') as file:
            docs = file.read()
            return docs


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
                      reTrainingHour=%s\
                  WHERE id=%s" % \
                  (form.get('cellphone'), form.get('fax'), form.get('tax_no'), form.get('com_email'), form.get('company_name'),
                   form.get('invoice_title'), form.get('company_address'), form.get('priv_email'), form.get('industry'),
                   form.get('phone'), form.get('birthday'), form.get('address'), form.get('job_title'), form.get('training_status'),
                   form.get('invoice_tax_no'), form.get('education_id'), form.get('edu_school'), form.get('city'), form.get('zip'),
                   form.get('company_city'), form.get('company_zip'), form.get('reTrainingCat'), form.get('reTrainingCode'),
                   form.get('reTrainingHour'), form.get('id'))
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
                  isAlt = 0 and reg_course.training_status = training_status_code.id""" %uid
        self.total_number = execSql.execSql(execStr)[0][0]
        # page為第幾頁，gap為一頁幾個
        page = request.get('page', '')
        gap = request.get('gap', '')

        if page and gap:
            page = int(page)
            gap = int(gap)
            sqlStr = """SELECT reg_course.name, reg_course.seatNo, reg_course.training_status
                        FROM reg_course, training_status_code
                        WHERE uid = '{}' and
                              isAlt = 0 and
                              reg_course.training_status = training_status_code.id
                        ORDER BY seatNo""".format(uid)
            admit = execSql.execSql(sqlStr)
            data = []
            data.append(admit[(page - 1) * gap:page * gap])

            courseList = []
            for course in context.getChildNodes():
                date = course.startDateTime.strftime('%m%d')
                title = course.title
                courseList.append(date + '-' + title)

            self.courseList = courseList
            self.data = data
            return self.list_print_template()
        else:
            return self.template()


class HasExportView(BrowserView):
    template = ViewPageTemplateFile('template/has_export_view.pt')
    def __call__(self):
        request = self.request
        context = self.context
        execSql = SqlObj()
        data = {}
        nowDate = datetime.datetime.now().date()
        # 用來判斷是否要更新資料
        period = request.get('period', '')
        user_id = request.get('user_id', '')
        if period and user_id:
            path = api.content.get(UID=period).absolute_url_path()
            execStr = """UPDATE reg_course SET isAlt = 0, on_training = 1, path = '{}', uid = '{}' WHERE id = {}
                     """.format(path, period, user_id)
            execSql.execSql(execStr)
            request.response.redirect(context.absolute_url() + '/@@has_export_view')
        # 從mana_course開始
        for course in context.getChildNodes():
            uidList = []
            echelonDict = {}
            courseName = course.title
            # 所有的course
            for echelon in course.getChildNodes():
                echelonUID = echelon.UID()
                echelonID = echelon.id
                uidList.append(echelonUID)
                courseStart = echelon.courseStart
                courseEnd = echelon.courseEnd
                if nowDate >= courseStart and nowDate <= courseEnd:
                    echelonDict[echelonUID] = echelonID

            execStr = """SELECT cellphone,name,company_name,priv_email,phone,id FROM reg_course WHERE uid in {} AND isAlt != 0 AND 
                         on_training = 3""".format(tuple(uidList))
            result = execSql.execSql(execStr)
            # {uid: id}排序
            data[courseName] = [ result, sorted(echelonDict.items(), key= lambda x: x[1]) ]
        self.data = data
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


