# -*- coding: utf-8 -*-
from cshm.content import _
from cshm.content.browser.views import BasicBrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from Products.CMFPlone.utils import safe_unicode
from Products.CMFPlone.PloneBatch import Batch
from Acquisition import aq_inner
from plone.app.contenttypes.browser.folder import FolderView
from plone.app.contentlisting.interfaces import IContentListing
from Products.CMFCore.utils import getToolByName

from zope import component
from zope.intid.interfaces import IIntIds
#from zope.app.intid import IntIds
from z3c.relationfield import RelationValue

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
from transaction import commit

logger = logging.getLogger("cshm.content")


DATA = ['研討會', '職業安全衛生管理人員', '施工安全評估', '製程安全評估', '營造作業主管',
        '有害作業主管', '防火管理人', '急救人員', '危險物品運送人員', '具有危險之設備操作人員',
        '具有危險之機械操作人員', '高壓氣體作業主管', '特殊作業安全衛生人員', '一般安全衛生教育訓練',
        '營造業工地主任', '環境保護人員', '保稅工廠保稅業務人員', '勞工健康服務護理人員', '人因性危害評估專業人員', '其他', '在職回訓課程']


class CourseListBySubject(BasicBrowserView):

    template = ViewPageTemplateFile('template/course_list_by_subject.pt')
    def __call__(self):
        request = self.request
        portal = api.portal.get()
        alsoProvides(request, IDisableCSRFProtection)

        content = api.content.find(context=portal['resource']['course_template'], depth=1)

        courseList = {}
        for course in DATA:
            courseList[course] = []

        for item in content:
            obj = item.getObject()
            courseHour = 0
            subject = obj.subject
            for child in obj.getChildNodes():
                courseHour += child.hours

            data = {
                'id': obj.id,
                'url': obj.absolute_url(),
                'title': obj.title,
                'courseHour': courseHour
            }

            if subject:
                courseList[subject[0].encode()].append(data)
            else:
                courseList['其他'].append(data)

        self.courseList = courseList
        self.sortList = DATA
        return self.template()
