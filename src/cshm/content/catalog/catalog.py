# -*- coding: utf-8 -*-
from plone.indexer.decorator import indexer
from zope.interface import Interface
from Products.CMFPlone.utils import safe_unicode
from cshm.content.content.echelon import IEchelon

from mingtak.ECBase.browser.views import SqlObj


@indexer(IEchelon)
def classStatus_indexer(obj):
    return obj.classStatus


@indexer(IEchelon)
def quota_indexer(obj):
    return obj.qutoa


@indexer(IEchelon)
def altCount_indexer(obj):
    return obj.altCount


@indexer(IEchelon)
def studentCount_indexer(obj):
    sqlInstance = SqlObj()
    uid = obj.UID()
    sqlStr = """SELECT COUNT(id) FROM reg_course WHERE uid = '{}'""".format(uid)
    result = sqlInstance.execSql(sqlStr)
    return result[0]['COUNT(id)']

