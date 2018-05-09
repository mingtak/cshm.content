# -*- coding: utf-8 -*-
from cshm.content.content.echelon import IEchelon  # NOQA E501
from cshm.content.testing import CSHM_CONTENT_INTEGRATION_TESTING  # noqa
from plone import api
from plone.app.testing import setRoles
from plone.app.testing import TEST_USER_ID
from plone.dexterity.interfaces import IDexterityFTI
from zope.component import createObject
from zope.component import queryUtility

import unittest


try:
    from plone.dexterity.schema import portalTypeToSchemaName
except ImportError:
    # Plone < 5
    from plone.dexterity.utils import portalTypeToSchemaName


class EchelonIntegrationTest(unittest.TestCase):

    layer = CSHM_CONTENT_INTEGRATION_TESTING

    def setUp(self):
        """Custom shared utility setup for tests."""
        self.portal = self.layer['portal']
        setRoles(self.portal, TEST_USER_ID, ['Manager'])

    def test_ct_echelon_schema(self):
        fti = queryUtility(IDexterityFTI, name='Echelon')
        schema = fti.lookupSchema()
        self.assertEqual(IEchelon, schema)

    def test_ct_echelon_fti(self):
        fti = queryUtility(IDexterityFTI, name='Echelon')
        self.assertTrue(fti)

    def test_ct_echelon_factory(self):
        fti = queryUtility(IDexterityFTI, name='Echelon')
        factory = fti.factory
        obj = createObject(factory)

        self.assertTrue(
            IEchelon.providedBy(obj),
            u'IEchelon not provided by {0}!'.format(
                obj,
            ),
        )

    def test_ct_echelon_adding(self):
        setRoles(self.portal, TEST_USER_ID, ['Contributor'])
        obj = api.content.create(
            container=self.portal,
            type='Echelon',
            id='echelon',
        )
        self.assertTrue(
            IEchelon.providedBy(obj),
            u'IEchelon not provided by {0}!'.format(
                obj.id,
            ),
        )
