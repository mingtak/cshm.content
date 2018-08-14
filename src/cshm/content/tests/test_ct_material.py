# -*- coding: utf-8 -*-
from cshm.content.content.material import IMaterial  # NOQA E501
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


class MaterialIntegrationTest(unittest.TestCase):

    layer = CSHM_CONTENT_INTEGRATION_TESTING

    def setUp(self):
        """Custom shared utility setup for tests."""
        self.portal = self.layer['portal']
        setRoles(self.portal, TEST_USER_ID, ['Manager'])

    def test_ct_material_schema(self):
        fti = queryUtility(IDexterityFTI, name='material')
        schema = fti.lookupSchema()
        self.assertEqual(IMaterial, schema)

    def test_ct_material_fti(self):
        fti = queryUtility(IDexterityFTI, name='material')
        self.assertTrue(fti)

    def test_ct_material_factory(self):
        fti = queryUtility(IDexterityFTI, name='material')
        factory = fti.factory
        obj = createObject(factory)

        self.assertTrue(
            IMaterial.providedBy(obj),
            u'IMaterial not provided by {0}!'.format(
                obj,
            ),
        )

    def test_ct_material_adding(self):
        setRoles(self.portal, TEST_USER_ID, ['Contributor'])
        obj = api.content.create(
            container=self.portal,
            type='material',
            id='material',
        )
        self.assertTrue(
            IMaterial.providedBy(obj),
            u'IMaterial not provided by {0}!'.format(
                obj.id,
            ),
        )
