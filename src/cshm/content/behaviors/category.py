# -*- coding: utf-8 -*-

from cshm.content import _
from plone import schema
from plone.autoform.interfaces import IFormFieldProvider
from plone.dexterity.interfaces import IDexterityContent
from zope.schema.vocabulary import SimpleTerm
from zope.schema.vocabulary import SimpleVocabulary
from plone.supermodel import model
from zope.component import adapter
from zope.interface import implementer
from zope.interface import provider

category = SimpleVocabulary( 
    [ 
        SimpleTerm(value=u'government', title=_(u'政府單位')), 
        SimpleTerm(value=u'other', title=_(u'其他單位')), 
        SimpleTerm(value=u'foreign', title=_(u'國外機構')), 
    ])
 
@provider(IFormFieldProvider)
class ICategory(model.Schema):  

    category = schema.Choice( 
        title=_(u'Category'), 
        vocabulary=category, 
        required=True, 
    )
