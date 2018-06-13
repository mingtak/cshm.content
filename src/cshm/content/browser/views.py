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

        if studentCount >= context.quota:
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
        self.echelonBrain = {}
        for status in self.statusList:
            self.echelonBrain[status] = api.content.find(context=self.portal['mana_course'], Type='Echelon', classStatus=status)

        return self.template()


class EchelonListing(BrowserView):

    """ 班別列表 / 辦班作業管理頁面 """

    template = ViewPageTemplateFile("template/echelon_listing.pt")

    def __call__(self):
        self.portal = api.portal.get()
        #TODO 條件尚未明確 
        self.echelonBrain = api.content.find(context=self.portal, Type='Echelon')

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
