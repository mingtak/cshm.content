# -*- coding: utf-8 -*-
from cshm.content import _
from plone.app.registry.browser.controlpanel import RegistryEditForm
from plone.app.registry.browser.controlpanel import ControlPanelFormWrapper

from zope import schema
from plone.z3cform import layout
from z3c.form import form
from plone.directives import form as Form


class IOffice(Form.Schema):

    docsWorkflows = schema.Text(
        title=_(u"Documents workflows"),
#        description=_(u'User Enter to separate'),
        required=False,
    )

    office_header = schema.Text(
        title=_(u"Office Header"),
        description=_(u'User Enter to separate'),
        required=False,
    )
    count_office_header = schema.Text(
        title=_(u"Count Header"),
        required=False,
    )

    cell_msg_url = schema.TextLine(
        title=_(u'Cell Message Provider URL'),
        description=_(u'format: id,password,url'),
        required=True,
    )

    reg_ok_message = schema.Text(
        title=_(u'Registry OK Auto Message'),
        description=_(u"'name' for Student Name, 'course' for Course Name"),
        required=False,
    )

    reg_finish_alert_message = schema.Text(
        title=_(u'Registry Finish Alert Message'),
        description=_(u"Show in Registry Finish Page."),
        required=True,
    )

    email_template = schema.Text(
        title=_(u'Email Template'),
        description=_(u'One line one record'),
        required=False
    )

    msg_template = schema.Text(
        title=_(u'Msg Template'),
        description=_(u'One line one record'),
        required=False
    )


class OfficeControlPanelForm(RegistryEditForm):
    form.extends(RegistryEditForm)
    schema = IOffice

OfficeControlPanelView = layout.wrap_form(OfficeControlPanelForm, ControlPanelFormWrapper)
OfficeControlPanelView.label = _(u"Office Related Setting")

