# -*- coding: utf-8 -*-
from plone import api
from cshm.content import _


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

