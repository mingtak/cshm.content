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
        sqlStr = """SELECT reg_course.*, status
                    FROM reg_course, training_status_code
                    WHERE uid = '{}' and
                          isAlt = 0 and
                          on_training = 1 and
                          reg_course.training_status = training_status_code.id
                    ORDER BY isnull(isReserve),isReserve, isnull(seatNo), seatNo""".format(uid)
        self.admit = sqlInstance.execSql(sqlStr)

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
            sqlStr = """SELECT name,tax_no FROM reg_course WHERE id in {}""".format(tuple(userList))
            result = sqlInstance.execSql(sqlStr)
            data = []
#            data = {}
            # 假定它都選同一家公司
            for item in result:
                name = item[0]
                tax_no = item[1]
#                if data.has_key(tax_no):
#                    data[tax_no] += ',%s' %name
#                else:
#                    data[tax_no] = name
                try:
                    data[0] += ',%s' %name
                except:
                    data.append(name)
                    data.append(tax_no if tax_no else '')
            self.data = data
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
        receipte_date = request.get('receipte_date')
        title = request.get('title')
        code = request.get('code')
        detail1_money = request.get('detail1_money')
        detail2_money = request.get('detail2_money')
        detail1_name = request.get('detail1_name')
        detail2_name = request.get('detail2_name')
        detail1_note = request.get('detail1_note')
        detail2_note = request.get('detail2_note')

        import pdb;pdb.set_trace()
        sqlStr = """INSERT INTO `receipt`(`uid`, `user_id`, `money`, `type`, `receipt_date`, `apply_date`, `title`, `code`, `note`, 
                    `detail`) VALUES ()""".format()

        sqlInstance.execSql(sqlStr)

