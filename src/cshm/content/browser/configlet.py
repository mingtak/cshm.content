# -*- coding: utf-8 -*-
from cshm.content import _
from plone.app.registry.browser.controlpanel import RegistryEditForm
from plone.app.registry.browser.controlpanel import ControlPanelFormWrapper

from zope import schema
from plone.z3cform import layout
from z3c.form import form
from plone.directives import form as Form


class IOffice(Form.Schema):

    office_header = schema.Text(
        title=_(u"Office Header"),
        description=_(u'User Enter to separate'),
        required=False,
    )
    count_office_header = schema.Text(
        title=_(u"Office Header"),
        description=_(u'User Enter to separate'),
        required=False,
    )

class OfficeControlPanelForm(RegistryEditForm):
    form.extends(RegistryEditForm)
    schema = IOffice

OfficeControlPanelView = layout.wrap_form(OfficeControlPanelForm, ControlPanelFormWrapper)
OfficeControlPanelView.label = _(u"Office Related Setting")

