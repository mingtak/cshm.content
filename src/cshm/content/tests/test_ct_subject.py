# -*- coding: utf-8 -*-
from cshm.content.content.subject import ISubject  # NOQA E501
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


class SubjectIntegrationTest(unittest.TestCase):

    layer = CSHM_CONTENT_INTEGRATION_TESTING

    def setUp(self):
        """Custom shared utility setup for tests."""
        self.portal = self.layer['portal']
        setRoles(self.portal, TEST_USER_ID, ['Manager'])

    def test_ct_subject_schema(self):
        fti = queryUtility(IDexterityFTI, name='Subject')
        schema = fti.lookupSchema()
        self.assertEqual(ISubject, schema)

    def test_ct_subject_fti(self):
        fti = queryUtility(IDexterityFTI, name='Subject')
        self.assertTrue(fti)

    def test_ct_subject_factory(self):
        fti = queryUtility(IDexterityFTI, name='Subject')
        factory = fti.factory
        obj = createObject(factory)

        self.assertTrue(
            ISubject.providedBy(obj),
            u'ISubject not provided by {0}!'.format(
                obj,
            ),
        )

    def test_ct_subject_adding(self):
        setRoles(self.portal, TEST_USER_ID, ['Contributor'])
        obj = api.content.create(
            container=self.portal,
            type='Subject',
            id='subject',
        )
        self.assertTrue(
            ISubject.providedBy(obj),
            u'ISubject not provided by {0}!'.format(
                obj.id,
            ),
        )
