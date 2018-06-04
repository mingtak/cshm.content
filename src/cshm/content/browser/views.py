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

    def __call__(self):
        self.portal = api.portal.get()
        request = self.request
        form = request.form

        if not form.get('studId', False):
            return self.template()

        sqlInstance = SqlObj()
        sqlStr = """INSERT INTO `reg_course`(`cellphone`, `fax`, `tax-no`, `name`, `com-email`, `company-name`, \
                    `tax-title`, `company-address`, `priv-email`, `phone`, `birthday`, `address`, `job-title`, `studId`, `uuid`)
                    VALUES ('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')
        """.format(form.get('cellphone'), form.get('fax'), form.get('tax-no'), form.get('name'), form.get('com-email'),
            form.get('company-name'), form.get('tax-title'), form.get('company-address'), form.get('priv-email'), form.get('phone'),
            form.get('birthday'), form.get('address'), form.get('job-title'), form.get('studId'), self.context.UID())
        sqlInstance.execSql(sqlStr)
        api.portal.show_message(message=_(u"Registry success."), request=request, type='info')
        request.response.redirect(self.portal['courselist'].absolute_url())
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
        #TODO 條件尚未明確 
        self.echelonBrain = api.content.find(context=self.portal, Type='Echelon')

        return self.template()


class EchelonListing(BrowserView):

    """ 班別列表 / 辦班作業管理頁面 """

    template = ViewPageTemplateFile("template/echelon_listing.pt")

    def __call__(self):
        self.portal = api.portal.get()
        #TODO 條件尚未明確 
        self.echelonBrain = api.content.find(context=self.portal, Type='Echelon')

        return self.template()

