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
from z3c.relationfield.schema import RelationList, RelationChoice
from plone.app.vocabularies.catalog import CatalogSource
from cshm.content import _


class ISubject(model.Schema):
    """ Marker interface and Dexterity Python Schema for Subject
    """

    hours = schema.Float(
        title=_(u"Subject Hours"),
        required=True,
    )

    startDateTime = schema.Datetime(
        title=_(u"Start time"),
        required=False
    )

    notes = schema.Text(
        title=_(u"Notes"),
        required=False,
    )

    #attachFile = namedfile.NamedBlobFile(
    #    title=_(u'Attach File'),
    #    required=False,
    #)
    isQuiz = schema.Bool(
        title=_(u"Quiz"),
        required=True,
        default=True
    )

    teacher = RelationChoice(
        title=_(u"Teacher"),
        required=False,
        source=CatalogSource(Type='Teacher')
    )

    ignoreSchedule = schema.Bool(
        title=_(u'Ignore Schedule'),
        default=False,
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


@implementer(ISubject)
class Subject(Item):
    """
    """
