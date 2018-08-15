# -*- coding: utf-8 -*-
from cshm.content import _
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from mingtak.ECBase.browser.views import SqlObj
import json


class MaterialView(BrowserView):
    template = ViewPageTemplateFile('template/material_view.pt')
    def __call__(self):
        context = self.context
        portal = api.portal.get()
        uid = context.UID()
        request = self.request
        execSql = SqlObj()

        execStr = """SELECT * FROM material WHERE uid = '%s'""" %uid
        order = execSql.execSql(execStr)
        orderList = []
        totalMaterialList = {}
        if order:
            for item in order:
                tmp = dict(item)
                status = tmp['status']
                apply_time = tmp['apply_time']
                order_id = tmp['id']
                logistics_code = tmp['logistics_code']
                send_date = tmp['send_date']
                action = tmp['action']
                orderList.append([status, apply_time, order_id, logistics_code, send_date, action])

                for detail in tmp['detail'].split('/'):
                    if detail:
                        name = detail.split('*')[0]
                        amount = int(detail.split('*')[1])
                        if action == 'return':
                            amount = amount * -1
                        if totalMaterialList.has_key(name):
                            totalMaterialList[name] += amount
                        else:
                            totalMaterialList[name] =  amount

        self.orderList = orderList
        self.totalMaterialList = totalMaterialList

        materialList = []
        availableMaterial = context.availableMaterial
        for materialUID in availableMaterial:
            content = api.content.get(UID=materialUID)
            title = content.title
            price = content.price
            discountPrice = content.discountPrice
            unit = content.unit
            finalPrice = discountPrice if discountPrice else price

            materialList.append([title, finalPrice, unit])
        self.materialList = materialList
        return self.template()


class ManipulateMaterial(BrowserView):
    def __call__(self):
        request = self.request
        material_detail = request.get('detail')
        send_date = request.get('send_date')
        remark = request.get('remark')
        address = request.get('address')
        action = request.get('action')
        context = self.context

        uid = context.UID()
        course = context.getParentNode().title
        period = context.id
        context_url = context.absolute_url()
        organizer = api.user.get_current().getUserName()

        execSql = SqlObj()
        execStr = """INSERT INTO `material`( `course`, `period`, `uid`, `status`, `send_date`, `logistics_code`, `detail`, `remark`,
                 `organizer`, `address`, `action`) VALUES('{}','{}','{}','{}','{}','{}','{}','{}', '{}', '{}', '{}')
                  """.format(course, period, uid, '處理中(寫死)', send_date, 'TODO', material_detail, remark, organizer, address, action)
        execSql.execSql(execStr)


        request.response.redirect(context_url + '/@@material_view')


class PrintOrder(BrowserView):
    template_shipment = ViewPageTemplateFile('template/print_shipment_order.pt')
    template_material = ViewPageTemplateFile('template/print_material_order.pt')
    template_return = ViewPageTemplateFile('template/print_return_order.pt')
    def __call__(self):
        request = self.request
        order_id = request.get('order_id', '')
        uid = self.context.UID()
        type = request.get('type')
        execSql = SqlObj()
        if type == 'shipment':
            execStr = """SELECT * FROM material WHERE id = '%s'""" %order_id
            self.result = execSql.execSql(execStr)

            return self.template_shipment()

        elif type == 'return':
            execStr = """SELECT * FROM material WHERE id = '%s'""" %order_id
            self.result = execSql.execSql(execStr)

            return self.template_return()

        elif type =='material':
            execStr = """SELECT * FROM material WHERE uid = '%s'""" %uid
            result = execSql.execSql(execStr)
            organizerList = []
            applyTimeList = []
            returnList = {}
            materialList = {}
            availableList = {}
            availableMaterial = self.context.availableMaterial

            for available in availableMaterial:
                content = api.content.get(UID=available)
                title = content.title
                price = content.price
                discountPrice = content.discountPrice
                availableList[title] = discountPrice if discountPrice else price

            for item in result:
                for material in item[8].split('/'):
                    name = material.split('*')[0]
                    if name:
                        amount = int(material.split('*')[1])
                        if item[12] == 'return':
                            if returnList.has_key(name):
                                returnList[name] += amount
                            else:
                                returnList[name] = amount
                            amount = amount * -1

                        if materialList.has_key(name):
                            materialList[name][0] += amount
                        else:
                            price = availableList[name]
                            materialList[name] = [amount, price]
                if item[10] not in organizerList:
                    organizerList.append(item[10])
                if item[5] not in applyTimeList:
                    applyTimeList.append(item[5])

            self.returnList = returnList
            self.materialList = materialList
            self.organizerList = organizerList
            self.applyTimeList = applyTimeList
            return self.template_material()


class PrintPreview(BrowserView):
    template = ViewPageTemplateFile('template/print_preview.pt')
    def __call__(self):
        request = self.request
        order_id = request.get('order_id')
        execSql = SqlObj()
        context = self.context

        execStr = """SELECT * FROM material WHERE id = %s""" %order_id
        self.result = execSql.execSql(execStr)
        availableMaterial = context.availableMaterial

        materialList = {}
        for materialUID in availableMaterial:
            content = api.content.get(UID=materialUID)
            title = content.title
            price = content.price
            discountPrice = content.discountPrice
            unit = content.unit
            finalPrice = discountPrice if discountPrice else price

            materialList[title]= [finalPrice, unit]
        self.materialList = materialList
        return self.template()
