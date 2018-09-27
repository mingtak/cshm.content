# -*- coding: utf-8 -*-
from plone.app.textfield import RichText
from plone.autoform import directives
from plone.dexterity.content import Item
from plone.namedfile import field as namedfile
from plone.supermodel import model
from plone.supermodel.directives import fieldset
from z3c.form.browser.radio import RadioFieldWidget
from z3c.relationfield.schema import RelationList, RelationChoice
from plone.app.vocabularies.catalog import CatalogSource
from zope import schema
from zope.interface import implementer
from cshm.content import _


class IOfficialDoc(model.Schema):
    """ Marker interface and Dexterity Python Schema for Official Document
    """

    title = schema.TextLine(
        title=_(u"Official Document Title"),
        required=True,
    )

    docSN = schema.TextLine(
        title=_(u'Official Document Serial No.'),
        required=True,
    )

    docDate = schema.Date(
        title=_(u"Official Document Date."),
        required=True,
    )

    recipient = schema.TextLine(
        title=_(u'Recipient'),
        required=True,
    )

    detail_1 = schema.TextLine(
        title=_(u'Official Document Detail 1'),
    )

    detail_2 = schema.TextLine(
        title=_(u'Official Document Detail 2'),
    )

    detail_3 = schema.TextLine(
        title=_(u'Official Document Detail 3'),
    )

    detail_4 = schema.TextLine(
        title=_(u'Official Document Detail 4'),
    )

    detail_5 = schema.TextLine(
        title=_(u'Official Document Detail 5'),
    )

    detail_6 = schema.TextLine(
        title=_(u'Official Document Detail 6'),
    )

    detail_7 = schema.TextLine(
        title=_(u'Official Document Detail 7'),
    )

    detail_8 = schema.TextLine(
        title=_(u'Official Document Detail 8'),
    )

    detail_9 = schema.TextLine(
        title=_(u'Official Document Detail 9'),
    )

    detail_10 = schema.TextLine(
        title=_(u'Official Document Detail 10'),
    )




@implementer(IOfficialDoc)
class OfficialDoc(Item):
    """
    """
