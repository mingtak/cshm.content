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

DATA = {'缺氧危害預防研討會': '研討會',
'ISO45001說明會': '研討會',
'安全衛生教育訓練單位之專責輔導員講習': '研討會',
'職業安全管理師': '職業安全衛生管理人員',
'職業衛生管理師': '職業安全衛生管理人員',
'職業安全衛生管理員': '職業安全衛生管理人員',
'甲種職業安全衛生業務主管': '職業安全衛生管理人員',
'乙種職業安全衛生業務主管': '職業安全衛生管理人員',
'丙種職業安全衛生業務主管': '職業安全衛生管理人員',
'營造業甲種職業安全衛生業務主管': '職業安全衛生管理人員',
'營造業乙種職業安全衛生業務主管': '職業安全衛生管理人員',
'營造業丙種職業安全衛生業務主管': '職業安全衛生管理人員',
'施工安全評估人員': '施工安全評估',
'製程安全評估人員': '製程安全評估',
'擋土支撐作業主管': '營造作業主管',
'模板支撐作業主管': '營造作業主管',
'隧道等挖掘作業主管': '營造作業主管',
'隧道等襯砌作業主管': '營造作業主管',
'施工架組配作業主管': '營造作業主管',
'鋼構組配作業主管': '營造作業主管',
'露天開挖作業主管': '營造作業主管',
'屋頂作業主管': '營造作業主管',
'有機溶劑作業主管': '有害作業主管',
'鉛作業主管': '有害作業主管',
'缺氧作業主管': '有害作業主管',
'特定化學物質作業主管': '有害作業主管',
'粉塵作業主管': '有害作業主管',
'高壓室內作業主管': '有害作業主管',
'防火管理人初訓': '防火管理人',
'防火管理人複訓': '防火管理人',
'急救人員': '急救人員',
'危險物品運送人員專業訓練(初訓)': '危險物品運送人員',
'吊升荷重在三公噸以上之固定式起重機操作人員': '具有危險之機械操作人員',
'吊升荷重在三公噸以上之移動式起重機操作人員': '具有危險之機械操作人員',
'甲級鍋爐操作人員': '具有危險之設備操作人員',
'乙級鍋爐操作人員': '具有危險之設備操作人員',
'丙級鍋爐操作人員': '具有危險之設備操作人員',
'第一種壓力容器操作人員': '具有危險之設備操作人員',
'高壓氣體特定設備操作人員': '具有危險之設備操作人員',
'高壓氣體容器操作人員': '具有危險之設備操作人員',
'高壓氣體製造安全主任': '高壓氣體作業主管',
'高壓氣體供應及消費作業主管': '高壓氣體作業主管',
'高壓氣體製造安全作業主管': '高壓氣體作業主管',
'小型鍋爐操作人員': '特殊作業安全衛生人員',
'荷重在一公噸以上之堆高機操作人員': '特殊作業安全衛生人員',
'吊升荷重在零點五公噸以上未滿三公噸之固定式起重機操作人員': '特殊作業安全衛生人員',
'吊升荷重在零點五公噸以上未滿三公噸之移動式起重機操作人員': '特殊作業安全衛生人員',
'使用起重機具從事吊掛作業人員': '特殊作業安全衛生人員',
'以乙炔熔接裝置或氣體集合裝置從事金屬之熔接、切斷或加熱作業人員': '特殊作業安全衛生人員',
'一般安全衛生教育訓練': '一般安全衛生教育訓練',
'南科一般安全衛生教育訓練－營造業': '一般安全衛生教育訓練',
'南科一般安全衛生教育訓練－非營造業': '一般安全衛生教育訓練',
'營造業工地主任220小時職能訓練': '營造業工地主任',
'室內空氣品質維護管理專責人員': '環境保護人員',
'保稅工廠保稅業務人員': '保稅工廠保稅業務人員',
'乙級檢定考複習班': '職業安全衛生管理人員',
'甲級安全師總複習班': '職業安全衛生管理人員',
'甲級衛生師總複習班': '職業安全衛生管理人員',
'堆高機技術士檢定輔導班': '特殊作業安全衛生人員',
'勞工安全衛生管理員輔導考照班': '其他',
'三公噸以上固定式起重機架空式-地上操作輔導考照班': '其他',
'三公噸以上固定式起重機架空式-機上操作輔導考照班': '其他',
'三公噸以上移動式起重機伸臂可伸縮式輔導考照班': '其他',
'固定式起重機操作技術士檢定輔導班': '其他',
'業務主管結業測驗複習班': '職業安全衛生管理人員',
'勞工健康服務護理人員': '勞工健康服務護理人員',
'人因性危害評估專業人員': '人因性危害評估專業人員',
'火藥爆破作業人員安全衛生訓練': '其他',
'一般高壓氣體類作業主管訓練': '其他',
'職業安全衛生業務主管暨職業安全衛生管理人員在職教育訓練': '在職回訓課程',
'營造業職業安全衛生業務主管暨職業安全衛生管理人員在職教育訓練': '在職回訓課程',
'擋土支撐作業主管在職教育訓練': '在職回訓課程',
'模板支撐作業主管在職教育訓練': '在職回訓課程',
'施工架組配作業主管在職教育訓練': '在職回訓課程',
'隧道等挖掘作業主管在職教育訓練': '在職回訓課程',
'隧道等襯砌作業主管在職教育訓練': '在職回訓課程',
'有機溶劑作業主管在職教育訓練': '在職回訓課程',
'鉛作業主管在職教育訓練': '在職回訓課程',
'粉塵作業主管在職教育訓練': '在職回訓課程',
'缺氧作業主管在職教育訓練': '在職回訓課程',
'特定化學物質作業主管在職教育訓練': '在職回訓課程',
'急救人員在職教育訓練': '在職回訓課程',
'固定式起重機操作人員在職教育訓練': '在職回訓課程',
'移動式起重機操作人員在職教育訓練': '在職回訓課程',
'荷重在一公噸以上之堆高機操作人員在職教育訓練': '在職回訓課程',
'使用起重機具從事吊掛作業人員在職教育訓練': '在職回訓課程',
'鍋爐操作人員在職教育訓練': '在職回訓課程',
'第一種壓力容器操作人員安全衛生在職教育訓練': '在職回訓課程',
'高壓氣體特定設備操作人員安全衛生在職教育訓練': '在職回訓課程',
'有害作業主管在職教育訓練': '在職回訓課程',
'以乙炔熔接裝置或氣體集合裝置從事金屬之熔接、切斷或加熱作業人員安全衛生在職教育訓練': '在職回訓課程',
'營造業法施行前領有建築工程管理甲級或乙級技術士證者回訓課程講習': '在職回訓課程',
'高壓室內作業人員在職教育訓練': '在職回訓課程',
'高壓氣體、室內作業主管在職教育訓練': '在職回訓課程',
'起重機操作及吊掛作業人員安全衛生在職教育訓練': '在職回訓課程',
'具有危險性之機械操作人員在職教育訓練': '在職回訓課程',
'具有危險性之設備操作人員在職教育訓練': '在職回訓課程',
'各級業務主管在職教育訓練': '在職回訓課程',
'高壓氣體作業主管在職教育訓練': '在職回訓課程',
'營造作業主管在職教育訓練': '在職回訓課程',
'一般安全衛生在職教育訓練': '在職回訓課程',
'小型鍋爐操作人員在職教育訓練': '在職回訓課程',
'火藥爆破作業人員在職教育訓練': '在職回訓課程',
'露天開挖作業主管在職教育訓練': '在職回訓課程',
'危險物品運送人員專業訓練(複訓)': '在職回訓課程',
'起重機操作人員安全衛生在職教育訓練': '在職回訓課程',
'鋼構組配作業主管在職教育訓練': '在職回訓課程',
'高壓氣體容器操作人員安全衛生在職教育訓練': '在職回訓課程',
'營造業業務主管人員在職教育訓練': '在職回訓課程',
'施工安全評估人員在職教育訓練': '在職回訓課程',
'危險性之設備操作人員(鍋爐、一壓、小鍋)在職教育訓練': '在職回訓課程',
'製程安全評估人員在職教育訓練': '在職回訓課程',
'人字臂起重桿操作人員安全衛生在職教育訓練': '在職回訓課程',
'勞工健康服務護理人員在職教育訓練': '在職回訓課程',
'具有危險性之設備暨小型鍋爐操作人員在職教育訓練': '在職回訓課程',
'職業安全衛生管理人員在職教育訓練': '在職回訓課程',
'職業安全衛生業務主管在職教育訓練': '在職回訓課程',
'屋頂作業主管在職教育訓練': '在職回訓課程',
'一般安全衛生在職教育訓練－營造業': '在職回訓課程',
'一般安全衛生在職教育訓練－非營造業': '在職回訓課程'}


class ImportSubject(BasicBrowserView):

    #template = ViewPageTemplateFile('template/no_permission_template.pt')
    def __call__(self):
        request = self.request
        portal = api.portal.get()
        alsoProvides(request, IDisableCSRFProtection)

        course = portal['resource']['course_template'].getChildNodes()

        for item in course:
            if item.id == 'a0011':
                import pdb; pdb.set_trace()
            try:
                if item.subject:
                    item.subject.append(DATA[item.title.encode('utf-8')].decode('utf-8'))
                else:
                    item.subject = [DATA[item.title.encode('utf-8')].decode('utf-8')]
                item.reindexObject()
                logger.info('%s: %s/edit' % (item.title, item.absolute_url()))
#                import pdb; pdb.set_trace()
            except:
                pass
#                logger.info('%s: %s' % (item.title, item.absolute_url()))
#                import pdb; pdb.set_trace()
#            break
#                logger.info(item.absolute_url())
#        import pdb; pdb.set_trace()




