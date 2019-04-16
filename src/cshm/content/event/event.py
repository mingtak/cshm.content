# -*- coding: utf-8 -*-
from plone import api
from cshm.content import _


def changeId(obj, event):
    portal = api.portal.get()
    if not obj.title.endswith('期'):
        api.portal.show_message(message=_(u'Please ends with "echelon"'), request=obj.REQUEST, type='warn')
        request = obj.REQUEST
        request.response.redirect(request.getURL())
        return
    newId = obj.title.split('期')[0]
    parent = obj.getParentNode()
    if parent.has_key(newId) and parent[newId] != obj:
        api.portal.show_message(message=_(u'Warning!, Has a same name course in this folder, Plsear back page and rename'),
            request=obj.REQUEST, type='error')
        raise
    api.content.rename(obj=obj, new_id=str(newId))
    api.portal.show_message(message=_(u'Rename Already Done.'), request=obj.REQUEST, type='info')
    return

def addRoleObj(event):
    portal = api.portal.get()
    para = portal.REQUEST.form
    if 'selected' in para.get('form.widgets.isTeacher', [False]):
#        import pdb; pdb.set_trace()
        teacherObj = api.content.create(
            type='Teacher',
            container=portal['resource']['mana_teacher'],
            id=para.get('form.username'),
            title=para.get('form.widgets.fullname'),
            email=para.get('form.widgets.email'),
        )

#        portal.REQUEST.response.redirect(teacherObj.absolute_url())
        api.portal.show_message(
            message='%s created, go %s/edit to edit.' % (para.get('form.fullname'), teacherObj.absolute_url()),
            request=portal.REQUEST
        )


def moveObjectsToTop(obj, event):
    """
    Moves Items to the top of its folder
    """
    folder = obj.getParentNode()
    if folder != None and hasattr(folder, 'moveObjectsToTop'):
        folder.moveObjectsToTop(obj.id)

