# -*- coding: utf-8 -*-

# from plone import api
from cshm.content import _
from plone.dexterity.interfaces import IDexterityContent
from zope.globalrequest import getRequest
from zope.interface import implementer
from zope.schema.interfaces import IVocabularyFactory
from zope.schema.vocabulary import SimpleTerm
from zope.schema.vocabulary import SimpleVocabulary


class VocabItem(object):
    def __init__(self, token, value):
        self.token = token
        self.value = value


@implementer(IVocabularyFactory)
class ClassTime(object):
    """
    """

    def __call__(self, context):
        # Just an example list of content for our vocabulary,
        # this can be any static or dynamic data, a catalog result for example.
        items = [
            VocabItem(u'notYat', _(u'notYat')),
            VocabItem(u'inDay', _(u'inDay')),
            VocabItem(u'inEvening', _(u'inEvening')),
            VocabItem(u'inWeekend', _(u'inWeekend')),
            VocabItem(u'inWeekendEvening', _(u'inWeekendEvening')),
            VocabItem(u'complex', _(u'complex')),
            VocabItem(u'phone', _(u'phone'))
        ]

        # Fix context if you are using the vocabulary in DataGridField.
        # See https://github.com/collective/collective.z3cform.datagridfield/issues/31:  # NOQA: 501
        if not IDexterityContent.providedBy(context):
            req = getRequest()
            context = req.PARENTS[0]

        # create a list of SimpleTerm items:
        terms = []
        for item in items:
            terms.append(
                SimpleTerm(
                    value=item.token,
                    token=str(item.token),
                    title=item.value,
                )
            )
        # Create a SimpleVocabulary from the terms list and return it:
        return SimpleVocabulary(terms)


ClassTimeFactory = ClassTime()
