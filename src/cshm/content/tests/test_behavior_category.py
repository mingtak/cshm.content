# -*- coding: utf-8 -*-
from cshm.content.behaviors.category import ICategory
from cshm.content.testing import CSHM_CONTENT_INTEGRATION_TESTING  # noqa
from plone.app.testing import setRoles
from plone.app.testing import TEST_USER_ID
from plone.behavior.interfaces import IBehavior
from zope.component import getUtility

import unittest


class CategoryIntegrationTest(unittest.TestCase):

    layer = CSHM_CONTENT_INTEGRATION_TESTING

    def setUp(self):
        """Custom shared utility setup for tests."""
        self.portal = self.layer['portal']
        setRoles(self.portal, TEST_USER_ID, ['Manager'])

    def test_behavior_category(self):
        behavior = getUtility(IBehavior, 'cshm.content.category')
        self.assertEqual(
            behavior.marker,
            ICategory,
        )
        behavior_name = 'cshm.content.behaviors.category.ICategory'
        behavior = getUtility(IBehavior, behavior_name)
        self.assertEqual(
            behavior.marker,
            ICategory,
        )
