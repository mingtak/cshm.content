# -*- coding: utf-8 -*-
from cshm.content import _
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from Products.CMFPlone.utils import safe_unicode
from cshm.content.browser.configlet import IOffice
from plone.protect.interfaces import IDisableCSRFProtection
from zope.interface import alsoProvides
from datetime import datetime
import json
from mingtak.ECBase.browser.views import SqlObj


class Basic(BrowserView):

    def GetSignatureSample(self):
        sqlInstance = SqlObj()
        user = api.user.get_current().getUserName()
        sqlStr = """SELECT * FROM `signature_sample` WHERE user = '{}' or user is NULL """.format(user)
        return sqlInstance.execSql(sqlStr)


class CreateReceiveView(BrowserView):
    template = ViewPageTemplateFile('template/create_receive_view.pt')
    def __call__(self):
        request = self.request
        return self.template()


class SearchSignatureView(BrowserView):
    template = ViewPageTemplateFile('template/search_signature_view.pt')
    template2 = ViewPageTemplateFile('template/search_signature_result.pt')
    def __call__(self):
        request = self.request
        year = request.get('year')
        category = request.get('category')
        item = request.get('item')
        kw = request.get('kw')

        sqlInstance = SqlObj()
        if year and category:
            sqlStr = """SELECT * FROM signature WHERE year = {} and category = '{}' ORDER BY code DESC""".format(year, category)
            self.result = sqlInstance.execSql(sqlStr)
            return self.template2()
        elif item and kw:
            sqlStr = """SELECT * FROM signature WHERE {} = '{}' ORDER BY code DESC""".format(item, kw)
            self.result = sqlInstance.execSql(sqlStr)
            return self.template2()
        else:
            sqlStr = """SELECT DISTINCT(year) FROM `signature`"""
            self.yearList = sqlInstance.execSql(sqlStr)

            return self.template()


class SqlSignatureSample(BrowserView):
    def __call__(self):
        request = self.request
        id = request.get('id')
        action = request.get('action')
        sqlInstance = SqlObj()
        if action == 'search':
            sqlStr = """SELECT * FROM signature_sample WHERE id = {}""".format(id)
            result = dict(sqlInstance.execSql(sqlStr)[0])
            return json.dumps(result)
        elif action == 'delect':
            sqlStr = """DELETE FROM signature_sample WHERE id = {}""".format(id)
            sqlInstance.execSql(sqlStr)


class AddSignatureSample(BrowserView):
    template = ViewPageTemplateFile('template/add_signature_sample.pt')
    def __call__(self):
        request = self.request
        sample_name = request.get('sample_name')
        category = request.get('category')
        title = request.get('title')
        detail_1 = request.get('detail_1')
        detail_2 = request.get('detail_2')
        detail_3 = request.get('detail_3')
        detail_4 = request.get('detail_4')
        user = api.user.get_current().getUserName()
        if sample_name:
            sqlInstance = SqlObj()
            sqlStr = """INSERT INTO signature_sample(`sample_name`, `user`, `category`, `title`, `detail_1`, `detail_2`, `detail_3`, 
                        `detail_4`) VALUES('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')""".format(sample_name, user, category,
                         title, detail_1, detail_2, detail_3, detail_4)
            sqlInstance.execSql(sqlStr)
            api.portal.show_message(message='新增範本成功!'.decode(), request=request)
            abs_url = api.portal.get().absolute_url()
            request.response.redirect(abs_url + '/signature_view')
        else:
            return self.template()


class SignatureView(BrowserView):
    template = ViewPageTemplateFile('template/signature_view.pt')
    def __call__(self):
        sqlInstance = SqlObj()
        sqlStr = """SELECT * FROM signature  ORDER BY code desc LIMIT 50"""
        self.result = sqlInstance.execSql(sqlStr)

        return self.template()


class ManageSignature(BrowserView):
    template = ViewPageTemplateFile('template/manage_signature.pt')
    def __call__(self):
        request = self.request
        id = request.get('id')
        status = request.get('status')
        sqlInstance = SqlObj()
        if status:
            sqlStr = """UPDATE signature SET status = '{}' WHERE id = {}""".format(status, id)
            sqlInstance.execSql(sqlStr)
            url = api.portal.get().absolute_url() + '/signature_view'
            api.portal.show_message(message='更新狀態成功!'.decode(), request=request)
            request.response.redirect(url)
        else:
            sqlStr = """SELECT * FROM signature WHERE id = {}""".format(id)
            self.result = sqlInstance.execSql(sqlStr)
            return self.template()


class ModifySignature(Basic):
    template = ViewPageTemplateFile('template/modify_signature.pt')
    def __call__(self):
        request = self.request
        id = request.get('id')
        type = request.get('type')
        category = request.get('category')
        date = request.get('date')
        department = request.get('department')
        title = request.get('title')
        detail_1 = request.get('detail_1')
        detail_2 = request.get('detail_2')
        detail_3 = request.get('detail_3')
        detail_4 = request.get('detail_4')
        sqlInstance = SqlObj()
        if category and title:
            sqlStr = """UPDATE signature SET title='{}',category='{}',date='{}',type='{}',department='{}',title='{}',detail_1='{}',
                        detail_2='{}',detail_3='{}',detail_4='{}' WHERE id = {}""".format(title, category, date, type, department, title,
                        detail_1, detail_2, detail_3, detail_4, id)
            sqlInstance.execSql(sqlStr)
            url = api.portal.get().absolute_url() + '/signature_view'
            api.portal.show_message(message='更新簽呈成功!'.decode(), request=request)
            request.response.redirect(url)
        else:
            sqlStr = """SELECT * FROM signature WHERE id = {}""".format(id)
            self.result = sqlInstance.execSql(sqlStr)
            self.sample = self.GetSignatureSample()

            return self.template()



class AddSignature(Basic):
    template = ViewPageTemplateFile('template/add_signature.pt')
    def __call__(self):
        request = self.request
        type = request.get('type')
        category = request.get('category')
        date = request.get('date')
        department = request.get('department')
        title = request.get('title')
        detail_1 = request.get('detail_1')
        detail_2 = request.get('detail_2')
        detail_3 = request.get('detail_3')
        detail_4 = request.get('detail_4')
        user = api.user.get_current().getUserName()
        if type and category and date:
            sqlInstance = SqlObj()
            year = datetime.now().year - 1911
            sqlStr = """SELECT MAX(code) FROM signature WHERE year = {}""".format(year)
            code = sqlInstance.execSql(sqlStr)[0][0]
            code = int(code) + 1 if code else 1
            sqlStr = """INSERT INTO `signature`(`user`, `category`, `type`, `date`, `department`, `title`, `detail_1`, `detail_2`,
                       `detail_3`, `detail_4`, year, code) VALUES ('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}', {}, {})
                     """.format(user, category, type, date, department, title, detail_1, detail_2, detail_3, detail_4, year, code)
            sqlInstance.execSql(sqlStr)
            api.portal.show_message(message='新增簽呈成功!'.decode(), request=request)
            abs_url = api.portal.get().absolute_url()
            request.response.redirect(abs_url + '/signature_view')

        self.sample = self.GetSignatureSample()

        return self.template()


class SearchOfficial(BrowserView):
    template = ViewPageTemplateFile('template/search_official.pt')
    template2 = ViewPageTemplateFile('template/search_official_result.pt')
    def __call__(self):
        request = self.request
        docHeader = request.get('docHeader')
        docSN = request.get('docSN')
        if docHeader or docSN:
            portal = api.portal.get()
            query = {
                'context': portal['official_doc'],
                'portal_type': 'OfficialDoc',
            }
            if docHeader:
                query['docHeader_indexer'] = docHeader
            if docSN:
                query['docSN_indexer'] = docSN

            result = api.content.find(**query)
            data = []
            for item in result:
                obj = item.getObject()
                title = obj.title
                code = obj.docHeader + obj.docSN
                date = obj.docDate.strftime('%Y-%m-%d')
                recipient = obj.recipient
                data.append([title, code, date, recipient])
            self.data = data
            return self.template2()

        return self.template()


class CreateOfficialView(BrowserView):
    template = ViewPageTemplateFile('template/create_official_view.pt')
    def __call__(self):
        self.office_header = api.portal.get_registry_record('office_header', interface=IOffice)
        request = self.request
        official_header = request.get('official_header')
        title = request.get('title')
        date = request.get('date')
        recipient = request.get('recipient')
        detail_1 = request.get('detail_1')
        detail_2 = request.get('detail_2')
        detail_3 = request.get('detail_3')
        detail_4 = request.get('detail_4')
        detail_5 = request.get('detail_5')
        detail_6 = request.get('detail_6')
        detail_7 = request.get('detail_7')
        detail_8 = request.get('detail_8')
        detail_9 = request.get('detail_9')
        detail_10 = request.get('detail_10')

        if title and date:
            date = datetime.strptime(date, '%Y-%m-%d')
            alsoProvides(request, IDisableCSRFProtection)
            portal = api.portal.get()
            config_office_header = api.portal.get_registry_record('office_header', interface=IOffice)
            config_count = api.portal.get_registry_record('count_office_header', interface=IOffice)

            if config_count:
                official_header = safe_unicode(official_header)
                config_count = json.loads(config_count)

                if config_count.has_key(official_header):
                    config_count[official_header] += 1
                    count = str(config_count[official_header])
                else:
                    config_count[official_header] = 1
                    count = '1'

                config_count = json.dumps(config_count)

            else:
                config_count = {}
                config_count[official_header] = 1
                config_count = json.dumps(config_count)
                count = '1'
            id = ''
            for item in config_office_header.split('/'):
                if item.split(':')[0] == official_header:
                    id += item.split(':')[1]

            for i in range(5 - len(count)):
                count = '0' + count
            id += count

            obj = api.content.create(
                type='OfficialDoc',
                title=title,
                docDate=date,
                recipient=recipient,
                docSN=count,
                docHeader=official_header,
                detail_1=detail_1,
                detail_2=detail_2,
                detail_3=detail_3,
                detail_4=detail_4,
                detail_5=detail_5,
                detail_6=detail_6,
                detail_7=detail_7,
                detail_8=detail_8,
                detail_9=detail_9,
                detail_10=detail_10,
                container=portal['official_doc'],
                id=id
            )
            uid = obj.UID()
            api.portal.set_registry_record('count_office_header', safe_unicode(config_count), interface=IOffice)
            url = api.portal.get().absolute_url() + '/official_view?uid=%s' %uid
            request.response.redirect(url)
        else:
            return self.template()


class OfficialView(BrowserView):
    template = ViewPageTemplateFile('template/official_view.pt')
    def __call__(self):
        request = self.request
        uid = request.get('uid') if request.get('uid') else self.context.UID()
        self.uid = uid
        obj = api.content.get(UID=uid)
        self.title = obj.title

        header = obj.docHeader
        sn = obj.docSN
        for i in range(5 - len(sn)):
            header += '0'
        header += sn
        self.code = header
        self.date = obj.docDate
        self.recipient = obj.recipient
        self.detail_1 = obj.detail_1
        self.detail_2 = obj.detail_2
        self.detail_3 = obj.detail_3
        self.detail_4 = obj.detail_4
        self.detail_5 = obj.detail_5
        self.detail_6 = obj.detail_6
        self.detail_7 = obj.detail_7
        self.detail_8 = obj.detail_8
        self.detail_9 = obj.detail_9
        self.detail_10 = obj.detail_10

        return self.template()
