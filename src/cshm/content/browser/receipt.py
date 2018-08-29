# -*- coding: utf-8 -*-
from cshm.content import _
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from mingtak.ECBase.browser.views import SqlObj
import json
import datetime


class ReceiptList(BrowserView):
    template = ViewPageTemplateFile("template/receipt_list.pt")
    def __call__(self):
        self.portal = api.portal.get()
        context = self.context

        uid = context.UID()
        sqlInstance = SqlObj()
        # 正取名單
        sqlStr = """SELECT reg_course.*,receipt_money.money FROM reg_course LEFT JOIN receipt_money ON receipt_money.user_id = reg_course.id 
                    WHERE reg_course.uid = '{}' and isAlt = 0 and on_training = 1 ORDER BY reg_course.seatNo""".format(uid)
        self.admit = sqlInstance.execSql(sqlStr)
        sqlStr = """SELECT user_id FROM receipt"""
        user_id = sqlInstance.execSql(sqlStr)
        idList = []
        for ids in user_id:
            for id in ids[0].split(','):
                if id:
                    idList.append(int(id))
        self.idList = idList
        return self.template()


class ReceiptCreateView(BrowserView):
    template = ViewPageTemplateFile("template/receipt_create.pt")
    def __call__(self):
        request = self.request
        userList = request.get("userList", "")
        if userList:
            userList = json.loads(userList)
            if type(userList) == int:
                userList = [userList, 'aa']
            sqlInstance = SqlObj()
            sqlStr = """SELECT name,tax_no,company_name,receipt_money.money FROM reg_course LEFT JOIN receipt_money ON receipt_money.user_id = 
                        reg_course.id WHERE reg_course.id in {}""".format(tuple(userList))
            result = sqlInstance.execSql(sqlStr)
            title = ''
            totalMoney = 0
            for item in result:
                name = item[0]
                self.tax_no = item[1] if item[1] != 'None' and item[1] else ''
                money = item[3]
                company_name = item[2]

                if money:
                    totalMoney += money

                if company_name:
                    title = company_name
                elif title:
                    title+= ',%s' %name
                else:
                    title = name

            self.title = title
            self.totalMoney = totalMoney
            try:
                userList.remove('aa')
            except:
                pass
            self.userList = userList

        return self.template()


class DoReceiptCreate(BrowserView):
    def __call__(self):
        request = self.request
        uid = self.context.UID()
        sqlInstance = SqlObj()
        user_id = json.loads(request.get('user_id'))
        total_money = request.get('total_money')
        type = request.get('type')
        apply_date = request.get('apply_date')
        receipt_date = request.get('receipt_date')
        title = request.get('title')
        code = request.get('code')
        note = request.get('note')
        detail1_money = request.get('detail1_money')
        detail2_money = request.get('detail2_money')
        detail1_name = request.get('detail1_name')
        detail2_name = request.get('detail2_name')
        detail1_note = request.get('detail1_note')
        detail2_note = request.get('detail2_note')
        ids = ''
        for id in user_id:
            ids += '%s,' %id

        detail = detail1_name + '/' + detail1_money + '/' + detail1_note + ',' +detail2_name + '/' + detail2_money + '/' + detail2_note
        sqlStr = """INSERT INTO `receipt`(`uid`, `user_id`, `money`, `type`, `receipt_date`, `apply_date`, `title`, `code`, `note`,
                    `detail`) VALUES ('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}')""".format(uid, ids, total_money, type,
                    receipt_date, apply_date, title, code, note, detail)

        sqlInstance.execSql(sqlStr)
        request.response.redirect(self.context.absolute_url() + '/@@receipt_list')


class UpdateReceiptMoney(BrowserView):
    def __call__(self):
        try:
            request = self.request
            money = int(request.get('money'))
            user_id = int(request.get('user_id'))
            uid = self.context.UID()
            sqlInstance = SqlObj()
            sqlStr = """SELECT id FROM receipt_money WHERE user_id = {}""".format(user_id)
            check = sqlInstance.execSql(sqlStr)
            if check:
                sqlStr = """UPDATE receipt_money SET money = {} WHERE user_id = {}""".format(money, user_id)
                sqlInstance.execSql(sqlStr)
            else:
                sqlStr = """INSERT INTO receipt_money(user_id, money, uid) VALUES({}, {}, '{}')""".format(user_id, money, uid)
                sqlInstance.execSql(sqlStr)
            return 'success'
        except  Exception as e:
            print e
            return 'error'


class ReceiptApplyForm(BrowserView):
    template = ViewPageTemplateFile('template/receipt_apply_form.pt')
    def __call__(self):
        request = self.request
        user_id = request.get('user_id')
        sqlInstance = SqlObj()
        sqlStr = """SELECT * FROM reg_course WHERE id = '{}'""".format(user_id)
        self.result = sqlInstance.execSql(sqlStr)
        return self.template()
