# -*- coding: utf-8 -*-
from plone.app.layout.viewlets import common as base
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from zope.component import queryUtility
from plone.i18n.normalizer.interfaces import IIDNormalizer


class PageBanner(base.ViewletBase):
    """  """
    def update(self):
        self.portal = api.portal.get()

        isFrontendView = api.content.get_view(name='is_frontend', context=self.portal, request=self.request)
        self.isFrontend = isFrontendView(self.view)


class CoverSlider(base.ViewletBase):
    """  """
    def update(self):
        self.portal = api.portal.get()

        isFrontendView = api.content.get_view(name='is_frontend', context=self.portal, request=self.request)
        self.isFrontend = isFrontendView(self.view)


class HeaderTools(base.ViewletBase):
    """  """
    def update(self):
        self.portal = api.portal.get()
        self.isAnon = api.user.is_anonymous()

        # TODO: Who can?
        self.isMana = 'Manager' in api.user.get_roles()

        isFrontendView = api.content.get_view(name='is_frontend', context=self.portal, request=self.request)
        self.isFrontend = isFrontendView(self.view)

        if not self.isAnon:
            self.current = api.user.get_current()
