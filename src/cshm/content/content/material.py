# -*- coding: utf-8 -*-
from plone.app.textfield import RichText
from plone.autoform import directives
from plone.dexterity.content import Item
from plone.namedfile import field as namedfile
from plone.supermodel import model
from plone.supermodel.directives import fieldset
from z3c.form.browser.radio import RadioFieldWidget
from zope import schema
from zope.interface import implementer
from cshm.content import _


class IMaterial(model.Schema):
    title = schema.TextLine(
        title=_(u"Title"),
        required=True,
    )

    version = schema.TextLine(
        title=_(u"Version"),
        required=False,
    )

    price = schema.Int(
        title=_(u"Price"),
        required=True,
    )

    discountPrice = schema.Int(
        title=_(u"Discount Price"),
        required=False,
    )

    unit = schema.TextLine(
        title=_(u"Unit"),
        required=False,
    )

    code = schema.TextLine(
        title=_(u"Code"),
        required=False,
    )

    cover = namedfile.NamedBlobImage(
        title=_(u'Cover'),
        required=False,
    )

    copyright = namedfile.NamedBlobFile(
        title=_(u'Copyright'),
        required=False,
    )

@implementer(IMaterial)
class Material(Item):
    """
    """
