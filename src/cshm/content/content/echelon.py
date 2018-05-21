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
        required=True
    )

    courseStart = schema.Date(
        title=_(u"Course Start Date"),
        required=True,
    )

    courseEnd = schema.Date(
        title=_(u"Course End Date"),
        required=True,
    )

    classTime = schema.TextLine(
        title=_(u"Class Time, ex. 0900-1800"),
        required=True,
    )

    trainingCenter = RelationChoice(
        title=_(u"Training Center"),
        required=True,
        source=CatalogSource(Type='TrainingCenter')
    )

    # directives.widget(level=RadioFieldWidget)
    # level = schema.Choice(
    #     title=_(u'Sponsoring Level'),
    #     vocabulary=LevelVocabulary,
    #     required=True
    # )

    # text = RichText(
    #     title=_(u'Text'),
    #     required=False
    # )

    # url = schema.URI(
    #     title=_(u'Link'),
    #     required=False
    # )

    # fieldset('Images', fields=['logo', 'advertisement'])
    # logo = namedfile.NamedBlobImage(
    #     title=_(u'Logo'),
    #     required=False,
    # )

    # advertisement = namedfile.NamedBlobImage(
    #     title=_(u'Advertisement (Gold-sponsors and above)'),
    #     required=False,
    # )

    # directives.read_permission(notes='cmf.ManagePortal')
    # directives.write_permission(notes='cmf.ManagePortal')
    # notes = RichText(
    #     title=_(u'Secret Notes (only for site-admins)'),
    #     required=False
    # )


@implementer(IEchelon)
class Echelon(Container):
    """
    """
