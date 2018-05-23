# -*- coding: utf-8 -*-
from cshm.content import _
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from Products.CMFPlone.utils import safe_unicode
from mingtak.ECBase.browser.views import SqlObj
import logging


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


class CourseListing(BrowserView):

    template = ViewPageTemplateFile("template/course_listing.pt")

    def __call__(self):
        self.portal = api.portal.get()
        self.echelonBrain = api.content.find(context=self.portal, Type='Echelon')

        return self.template()
