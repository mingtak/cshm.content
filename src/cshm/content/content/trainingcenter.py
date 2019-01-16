# -*- coding: utf-8 -*-
from plone.app.textfield import RichText
from plone.autoform import directives
from plone.dexterity.content import Container
from plone.namedfile import field as namedfile
from plone.supermodel import model
from plone.supermodel.directives import fieldset
from z3c.form.browser.radio import RadioFieldWidget
from zope import schema
from zope.interface import implementer
from cshm.content import _


class ITrainingcenter(model.Schema):
    """ Marker interface and Dexterity Python Schema for Trainingcenter
    """

    address = schema.TextLine(
        title=_(u"Address"),
        required=True,
    )

    phone = schema.TextLine(
        title=_(u"Phone Number."),
        required=True,
    )

    fax = schema.TextLine(
        title=_(u"Fax Number."),
        required=True,
    )

    code = schema.TextLine(
        title=_(u"Training Center Code."),
        required=False,
    )

    simpleTitle = schema.TextLine(
        title=_(u"Simple Title"),
        required=False,
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


@implementer(ITrainingcenter)
class Trainingcenter(Container):
    """
    """
