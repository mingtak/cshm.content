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

import requests
import logging


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


class RegCourse(BrowserView):

    template = ViewPageTemplateFile("template/reg_course.pt")

    def makeSqlStr(self):
        form = self.request.form

        uid = self.context.UID()
        path = self.context.virtual_url_path()
        isAlt = self.isAlt()

        sqlStr = """INSERT INTO `reg_course`(`cellphone`, `fax`, `tax-no`, `name`, `com-email`, `company-name`, \
                    `tax-title`, `company-address`, `priv-email`, `phone`, `birthday`, `address`, `job-title`, \
                    `studId`, `uid`, `path`, `isAlt`)
                    VALUES ('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')
        """.format(form.get('cellphone'), form.get('fax'), form.get('tax-no'), form.get('name'), form.get('com-email'),
            form.get('company-name'), form.get('tax-title'), form.get('company-address'), form.get('priv-email'), form.get('phone'),
            form.get('birthday'), form.get('address'), form.get('job-title'), form.get('studId'), uid, path, isAlt)
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


    def __call__(self):
        self.portal = api.portal.get()
        context = self.context
        request = self.request
        alsoProvides(request, IDisableCSRFProtection)
        form = request.form

        # 檢查是進入報名頁，或者為報名程序
        if not form.get('studId', False):
            return self.template()

        if context.classStatus == 'altFull':
            api.portal.show_message(message=_(u"Quota Full include Alternate."), request=request, type='info')
            request.response.redirect(self.portal['training']['courselist'].absolute_url())
            return


        # 報名寫入 DB
        sqlInstance = SqlObj()
        sqlStr = self.makeSqlStr()
        sqlInstance.execSql(sqlStr)

        # 檢查額滿(含正備取)
        self.checkFull()

        # reindex
        context.reindexObject(idxs=['studentCount', 'classStatus'])

        api.portal.show_message(message=_(u"Registry success."), request=request, type='info')
        request.response.redirect(self.portal['training']['courselist'].absolute_url())
        return


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


class StudentsList(BrowserView):

    """ 報名作業 / 學生列表 """

    template = ViewPageTemplateFile("template/students_list.pt")

    def __call__(self):
        self.portal = api.portal.get()
        context = self.context

        uid = context.UID()
        sqlInstance = SqlObj()

        sqlStr = """SELECT * FROM reg_course WHERE uid = '{}' and isAlt = 0 ORDER BY seatNo""".format(uid)
        self.admit = sqlInstance.execSql(sqlStr) # 正取

        sqlStr = """SELECT * FROM reg_course WHERE uid = '{}' and isAlt > 0 ORDER BY isAlt""".format(uid)
        self.waiting = sqlInstance.execSql(sqlStr) # 備取

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

        sqlStr = """SELECT id, seatNo FROM `reg_course` WHERE uid = '{}' and isAlt = 0""".format(uid)
        result = sqlInstance.execSql(sqlStr)

        for item in result:
            id = item['id']
            if not item['seatNo']:
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
        sqlStr = """update `reg_course` set isAlt = 0 where id = {}""".format(id)
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

        sqlInstance = SqlObj()
        sqlInstance.execSql(sqlStr)


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


class DebugView(BrowserView):

    def __call__(self):
        portal = api.portal.get()
        context = self.context
        request = self.request

        alsoProvides(request, IDisableCSRFProtection)
        import pdb; pdb.set_trace()
