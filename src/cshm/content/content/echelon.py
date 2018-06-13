# -*- coding: utf-8 -*-
from plone.app.textfield import RichText
from plone.autoform import directives
from plone.dexterity.content import Container
from plone.namedfile import field as namedfile
from plone.supermodel import model
from plone.supermodel.directives import fieldset
from z3c.form.browser.radio import RadioFieldWidget
from z3c.relationfield.schema import RelationList, RelationChoice
from plone.app.vocabularies.catalog import CatalogSource
from zope import schema
from zope.interface import implementer
from cshm.content import _


class IEchelon(model.Schema):
    """ Marker interface and Dexterity Python Schema for Echelon
    """

    duringTime = schema.Choice(
        title=_(u'During Time'),
        vocabulary='cshm.content.ClassTime',
        default=u'notYat',
        required=True
    )

    courseStart = schema.Date(
        title=_(u"Course Start Date"),
        required=False,
    )

    courseEnd = schema.Date(
        title=_(u"Course End Date"),
        required=False,
    )

    classTime = schema.TextLine(
        title=_(u"Class Time, ex. 0900-1800"),
        required=False,
    )

    trainingCenter = RelationChoice(
        title=_(u"Training Center"),
        required=False,
        source=CatalogSource(Type='TrainingCenter')
    )

    quota = schema.Int(
        title=_(u'Class Quota'),
        description=_(u'if 0, is not confirm'),
        default=0,
        required=True,
    )

    altCount = schema.Int(
        title=_(u'Alternate Count'),
        default=100,
        min=0,
        required=True,
    )

    fieldset('Handbook', fields=[
        'courseFee',
        'classStatus',
        'memo',
        'contact',
        'regDeadline',
        'discountProgram',
        'discountInfo_no_open',
        'prepareInfo',
        'courseHours',
        'detailClassTime',
        'submitClassDate',
        'craneType',
    ])

    courseFee = schema.Int(
        title=_(u'Course Fee'),
        required=False,
    )

    classStatus = schema.Choice(
        title=_(u'Course Status'),
        vocabulary='cshm.content.ClassStatus',
        default=u'registerFirst',
        required=False,
    )

    memo = RichText(
        title=_(u'Memo'),
        required=False,
    )

    contact = schema.TextLine(
        title=_(u'Contact'),
        required=False,
    )

    regDeadline = schema.Date(
        title=_(u'Registration Deadline'),
        required=False,
    )

    discountProgram = schema.TextLine(
        title=_(u'Discount Program'),
        required=False,
    )

    discountInfo_no_open = schema.TextLine(
        title=_(u'Discount Information, No Open'),
        required=False,
    )

    prepareInfo = schema.Text(
        title=_(u'Prepare Information'),
        required=False,
    )

    courseHours = schema.Int(
        title=_(u"Course Hours"),
        default=0,
        description=_(u"If 0, asking for phone"),
        required=True,
    )

    detailClassTime = schema.Text(
        title=_(u'Detail Class Time'),
        required=False,
    )

    submitClassDate = schema.Date(
        title=_(u'Submit Class Date'),
        required=False,
    )

    craneType = schema.TextLine(
        title=_(u'Crane Type'),
        required=False,
    )


@implementer(IEchelon)
class Echelon(Container):
    """
    """
