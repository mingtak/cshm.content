# -*- coding: utf-8 -*-
from cshm.content import _
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from mingtak.ECBase.browser.views import SqlObj
import json
import datetime
import pdfkit
import wkhtmltopdf
import weasyprint


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
        sqlStr = """SELECT id,user_id FROM receipt WHERE is_cancel = 0"""
        receipt = sqlInstance.execSql(sqlStr)
        idDict = {}
        for item in receipt:
            idList = item[1].split(',')
            for id in idList:
                if id:
                    idDict[id] = item[0]
        self.idDict = idDict
        return self.template()


class AdminReceiptList(BrowserView):
    template = ViewPageTemplateFile("template/admin_receipt_list.pt")
    def __call__(self):
        request = self.request
        sqlInstance = SqlObj()
        uid = self.context.UID()
        sqlStr = """SELECT * FROM receipt WHERE uid = '{}' ORDER BY receipt_date""".format(uid)
        self.result = sqlInstance.execSql(sqlStr)
        return self.template()


class ReceiptCreateView(BrowserView):
    template = ViewPageTemplateFile("template/receipt_create.pt")
    empty_template = ViewPageTemplateFile("template/empty_receipt_create.pt")
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
        else:
            return self.empty_template()


class DoReceiptCreate(BrowserView):
    def __call__(self):
        request = self.request
        # 用來判斷是不是來自空白的create_view
        is_empty = request.get('empty')
        if is_empty:
            uid = ''
            path = ''
        else:
            uid = self.context.UID()
            path = self.context.absolute_url_path()

        sqlInstance = SqlObj()
        user_id = request.get('user_id', 'anonymouse')
        if user_id != 'anonymouse':
            user_id = json.loads(user_id)

        total_money = request.get('total_money')
        type = request.get('type')
        apply_date = request.get('apply_date')
        receipt_date = request.get('receipt_date')
        title = request.get('title')
        tax_no = request.get('tax_no')
        note = request.get('note')
        detail1_money = request.get('detail1_money')
        detail2_money = request.get('detail2_money')
        detail1_name = request.get('detail1_name')
        detail2_name = request.get('detail2_name')
        detail1_note = request.get('detail1_note')
        detail2_note = request.get('detail2_note')
        user_name = api.user.get_current().getUserName()

        # 空白的create_view會抓的到request,非空白頁用context的trainingCenter
        training_center = request.get('training_center')
        if not training_center:
            training_center = self.context.trainingCenter.to_object.title

        if training_center == '台北職訓中心':
            training_center = '北訓'
        elif training_center == '基隆職訓中心':
            training_center = '基訓'
        elif training_center == '桃園職訓中心':
            training_center = '桃訓'
        elif training_center == '中壢職訓中心':
            training_center = '壢訓'
        elif training_center == '台中職訓中心':
            training_center = '中訓'
        elif training_center == '花蓮職訓中心':
            training_center = '花訓'
        elif training_center == '嘉義職訓中心':
            training_center = '嘉訓'
        elif training_center == '高雄職訓中心':
            training_center = '高訓'
        elif training_center == '南科職訓中心':
            training_center = '南訓'

        sqlStr = """SELECT MAX(serial_number) FROM receipt WHERE training_center = '{}'""".format(training_center)
        serial_number = sqlInstance.execSql(sqlStr)[0][0]
        if serial_number:
            serial_number += 1
        else:
            serial_number = 1

        year = datetime.datetime.now().year - 1911

        ids = ''
        if user_id != 'anonymouse':
            for id in user_id:
                ids += '%s,' %id

        detail = detail1_name + '/' + detail1_money + '/' + detail1_note + ',' +detail2_name + '/' + detail2_money + '/' + detail2_note
        sqlStr = """INSERT INTO `receipt`(`uid`, `user_id`, `money`, `type`, `receipt_date`, `apply_date`, `title`, `tax_no`, `note`,
                    `detail`, `undertaker`, training_center, serial_number, year, path) VALUES ('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}',
                    '{}','{}',{}, {}, '{}')""".format(uid, ids, total_money, type, receipt_date, apply_date, title, tax_no, note, detail,
                    user_name, training_center, serial_number, year, path)

        sqlInstance.execSql(sqlStr)
        contextUrl = self.context.absolute_url()
        url = contextUrl + '/@@receipt_create_view' if is_empty else contextUrl + '/@@receipt_list'

        request.response.redirect(url)


class UpdateReceiptMoney(BrowserView):
    def __call__(self):
        try:
            request = self.request
            money = int(request.get('money'))
            user_id = int(request.get('user_id'))
            uid = self.context.UID()
            path = self.context.absolute_url_path()
            sqlInstance = SqlObj()
            sqlStr = """SELECT id FROM receipt_money WHERE user_id = {}""".format(user_id)
            check = sqlInstance.execSql(sqlStr)
            if check:
                sqlStr = """UPDATE receipt_money SET money = {} WHERE user_id = {}""".format(money, user_id)
                sqlInstance.execSql(sqlStr)
            else:
                sqlStr = """INSERT INTO receipt_money(user_id, money, uid, path) VALUES({}, {}, '{}', '{}')
                         """.format(user_id, money, uid, path)
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
        sqlStr = """SELECT reg_course.*,training_status_code.status FROM reg_course,training_status_code WHERE reg_course.id = '{}' AND
                    reg_course.training_status = training_status_code.id""".format(user_id)
        self.result = sqlInstance.execSql(sqlStr)
        return self.template()


class CancelReceipt(BrowserView):
    def __call__(self):
        request = self.request
        receipt_id = request.get('receipt_id')
        if receipt_id:
            sqlInstance = SqlObj()
            sqlStr = """UPDATE receipt SET is_cancel = 1 WHERE id = {}""".format(int(receipt_id))
            sqlInstance.execSql(sqlStr)
            request.response.redirect('%s/@@admin_receipt_list' %self.context.absolute_url())


class ReceiptDetail(BrowserView):
    template = ViewPageTemplateFile("template/receipt_detail.pt")
    def __call__(self):
        request = self.request
        receipt_id = request.get('receipt_id')
        if receipt_id:
            sqlInstance = SqlObj()
            sqlStr = """SELECT * FROM receipt WHERE id = {}""".format(receipt_id)
            self.result = sqlInstance.execSql(sqlStr)
            self.receipt_id = receipt_id
        return self.template()


class PassReceipt(BrowserView):
    def __call__(self):
        request = self.request
        receipt_id = request.get('receipt_id')
        if receipt_id:
            sqlInstance = SqlObj()
            now = datetime.datetime.now().strftime('%Y-%m-%d')
            user_name = api.user.get_current().getUserName()
            time = request.get('time')
            if time == 'one':
                check = 'one_check'
                date = 'one_check_date'
            elif time == 'two':
                check = 'two_check'
                date = 'two_check_date'

            sqlStr = """UPDATE receipt SET {} = 1,{} = '{}',inspector = '{}'  WHERE id = {}
                     """.format(check, date, now, user_name, receipt_id)
            self.result = sqlInstance.execSql(sqlStr)
            request.response.redirect('%s/@@admin_receipt_list' %self.context.absolute_url())


class DownloadReceiptPdf(BrowserView):
    template = ViewPageTemplateFile("template/receipt_pdf_template.pt")
    def __call__(self):
        request = self.request
        receipt_id = request.get('receipt_id')
        sqlInstance = SqlObj()

        sqlStr = """SELECT * FROM receipt WHERE id = {}""".format(receipt_id)
        self.result = sqlInstance.execSql(sqlStr)
        self.trainingCenter = self.context.trainingCenter.to_object.title

        return self.template()


class SearchReceipt(BrowserView):
    template = ViewPageTemplateFile("template/search_receipt.pt")
    search_result = ViewPageTemplateFile("template/search_receipt_result.pt")
    def __call__(self):
        request = self.request
        action = request.get('action')
        self.action = action
        key_word = request.get('key_word')
        if key_word:
            sqlInstance = SqlObj()
            if action == 'between_date':
                start_date = key_word.split(',')[0]
                end_date = key_word.split(',')[1]
                sqlStr = """SELECT * FROM receipt WHERE receipt_date between '{}' AND '{}'""".format(start_date, end_date)
            elif action == 'code':
                keyList = key_word.split(',')
                trainingCenter = keyList[0]
                year = keyList[1]
                sn = keyList[2]
                sqlStr = """SELECT * FROM receipt WHERE 1"""
                if year:
                    sqlStr += ' AND year = %s' %year
                if trainingCenter:
                    sqlStr += " AND training_center LIKE '%s%%%%'" %trainingCenter
                if sn:
                    sqlStr += " AND serial_number = %s" %sn
            else:
                sqlStr = """SELECT * FROM receipt WHERE {} LIKE '{}%%'""".format(action, key_word)
            self.result = sqlInstance.execSql(sqlStr)
            return self.search_result()
        else:
            return self.template()
