# -*- coding: utf-8 -*-
from plone.indexer.decorator import indexer
from zope.interface import Interface
from Products.CMFPlone.utils import safe_unicode

from cshm.content.content.echelon import IEchelon

@indexer(IEchelon)
def classStatus_indexer(obj):
    return obj.courseStatus
