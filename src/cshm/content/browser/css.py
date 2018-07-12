# -*- coding: utf-8 -*-
from cshm.content import _
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from plone.protect.interfaces import IDisableCSRFProtection
from zope.interface import alsoProvides
import logging

logger = logging.getLogger("Backend.css")


class BackendCSS(BrowserView):
    """ Backend.css """

    template = ViewPageTemplateFile("template/backend_css.pt")


    def __call__(self):
        portal = api.portal.get()
        context = self.context
        request = self.request

        alsoProvides(request, IDisableCSRFProtection)

        css = request.form.get('css')
        if css:
            with open('/home/andy/cshm/zeocluster/src/cshm.theme/src/cshm/theme/browser/static/backend.css', 'w') as file:
                file.write(css)
                api.portal.show_message('Override css file finish.', request=request, type='info')
                request.response.redirect('%s/@@backend_css' % portal.absolute_url())

                return
        else:
            with open('/home/andy/cshm/zeocluster/src/cshm.theme/src/cshm/theme/browser/static/backend.css') as file:
                self.css = file.read()
                return self.template()






