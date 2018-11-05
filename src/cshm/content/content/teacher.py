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


class ITeacher(model.Schema):
    """ Marker interface and Dexterity Python Schema for Teacher
    """

    title = schema.TextLine(
        title=_(u"Teacher Name."),
        required=True,
    )

    activation = schema.Bool(
        title=_(u"Activation"),
        default=True,
        required=True,
    )

    isMember = schema.Bool(
        title=_(u"Is Member"),
        default=False,
        required=True,
    )

    idCardNo = schema.TextLine(
        title=_(u"Id Card Number."),
        required=True,
    )

    teacherSN = schema.TextLine(
        title=_(u"Teacher Serial Number."),
        required=False,
    )

    birthday = schema.Date(
        title=_(u"Birthday"),
        required=False,
    )

    homePhone = schema.TextLine(
        title=_(u"Home Phone Number."),
        required=False,
    )

    cellPhone = schema.TextLine(
        title=_(u"Cell Phone Number."),
        required=False,
    )

    fax = schema.TextLine(
        title=_(u"Fax Number."),
        required=False,
    )

    email = schema.TextLine(
        title=_(u"Email."),
        required=False,
    )

    idCardAddr = schema.TextLine(
        title=_(u"ID Card Address."),
        required=False,
    )

    contactAddr = schema.TextLine(
        title=_(u"Contact Address."),
        required=False,
    )

    edu_1 = schema.TextLine(
        title=_(u"Education 1."),
        required=False,
    )

    dep_1 = schema.TextLine(
        title=_(u"Department 1."),
        required=False,
    )

    degree_1 = schema.TextLine(
        title=_(u"Degree 1."),
        required=False,
    )

    gradYear_1 = schema.Int(
        title=_(u"Graduation Year 1."),
        min=1900,
        max=2100,
        required=False,
    )

    gradMonth_1 = schema.Int(
        title=_(u"Graduation Month 1."),
        min=1,
        max=12,
        required=False,
    )

    edu_2 = schema.TextLine(
        title=_(u"Education 2."),
        required=False,
    )

    dep_2 = schema.TextLine(
        title=_(u"Department 2."),
        required=False,
    )

    degree_2 = schema.TextLine(
        title=_(u"Degree 2."),
        required=False,
    )

    gradYear_2 = schema.Int(
        title=_(u"Graduation Year 2."),
        min=1900,
        max=2100,
        required=False,
    )

    gradMonth_2 = schema.Int(
        title=_(u"Graduation Month 2."),
        min=1,
        max=12,
        required=False,
    )

    edu_3 = schema.TextLine(
        title=_(u"Education 3."),
        required=False,
    )

    dep_3 = schema.TextLine(
        title=_(u"Department 3."),
        required=False,
    )

    degree_3 = schema.TextLine(
        title=_(u"Degree 3."),
        required=False,
    )

    gradYear_3 = schema.Int(
        title=_(u"Graduation Year 3."),
        min=1900,
        max=2100,
        required=False,
    )

    gradMonth_3 = schema.Int(
        title=_(u"Graduation Month 3."),
        min=1,
        max=12,
        required=False,
    )

    serviceUnit = schema.TextLine(
        title=_(u"Service Unit"),
        required=False,
    )

    serviceDep = schema.TextLine(
        title=_(u"Service Department"),
        required=False,
    )

    currentJob = schema.TextLine(
        title=_(u"Current Job"),
        required=False,
    )

    startWorkDate = schema.Date(
        title=_(u"Start Work Date"),
        required=False,
    )

    unitPhone = schema.TextLine(
        title=_(u"Unit Phone"),
        required=False,
    )

    personExp = schema.Text(
        title=_(u"Person Experience"),
        required=False,
    )

    license = schema.Text(
        title=_(u"Person License"),
        required=False,
    )

    unitAddress = schema.TextLine(
        title=_(u"Service Unit Address"),
        required=False,
    )

    creatUser = schema.TextLine(
        title=_(u"Created User"),
        required=False,
    )

    fieldset(_(u'Teach Related'), fields=['teachSubjects', 'teachTrainingCenter'])
    teachSubjects = RelationList(
        title=_(u'Teach Subjects'),
        default=[],
        required=False,
        value_type=RelationChoice(
            title=_(u"Subject"),
            source=CatalogSource(Type='Subject', path='/cshm/resource/corese_template'),
        )
    )

    teachTrainingCenter = RelationList(
        title=_(u'Teach Training Center'),
        default=[],
        required=False,
        value_type=RelationChoice(
            title=_(u"Training Center"),
            source=CatalogSource(Type='TrainingCenter', path='/cshm/resource/training_center'),
        )
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


@implementer(ITeacher)
class Teacher(Item):
    """
    """
