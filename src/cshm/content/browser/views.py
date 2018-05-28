# -*- coding: utf-8 -*-
from cshm.content import _
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
import transaction
from Products.CMFPlone.utils import safe_unicode
from mingtak.ECBase.browser.views import SqlObj
import logging
import csv
import datetime
from plone.protect.interfaces import IDisableCSRFProtection
from zope.interface import alsoProvides

class RegCourse(BrowserView):

    template = ViewPageTemplateFile("template/reg_course.pt")

    def __call__(self):
        self.portal = api.portal.get()
        request = self.request
        form = request.form

        if not form.get('studId', False):
            return self.template()

        sqlInstance = SqlObj()
        sqlStr = """INSERT INTO `reg_course`(`cellphone`, `fax`, `tax-no`, `name`, `com-email`, `company-name`, \
                    `tax-title`, `company-address`, `priv-email`, `phone`, `birthday`, `address`, `job-title`, `studId`, `uuid`)
                    VALUES ('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')
        """.format(form.get('cellphone'), form.get('fax'), form.get('tax-no'), form.get('name'), form.get('com-email'),
            form.get('company-name'), form.get('tax-title'), form.get('company-address'), form.get('priv-email'), form.get('phone'),
            form.get('birthday'), form.get('address'), form.get('job-title'), form.get('studId'), self.context.UID())
        sqlInstance.execSql(sqlStr)
        api.portal.show_message(message=_(u"Registry success."), request=request, type='info')
        request.response.redirect(self.portal['courselist'].absolute_url())
        return


class CourseListing(BrowserView):

    template = ViewPageTemplateFile("template/course_listing.pt")

    def __call__(self):
        self.portal = api.portal.get()
        self.echelonBrain = api.content.find(context=self.portal, Type='Echelon')

        return self.template()


class TeacherBasicInfo(BrowserView):

    template = ViewPageTemplateFile("template/teacher_basic_info.pt")

    def __call__(self):
        request = self.request
        alsoProvides(request, IDisableCSRFProtection)
        sqlInstance = SqlObj()
        sqlStr = "SELECT * FROM `teaching_area`"
        sqlResult = sqlInstance.execSql(sqlStr)
        teachingAreaList = {}
        for k,v in sqlResult:
            teachingAreaList.update({int(k): v.encode('utf-8')})
        self.teachingAreaList = teachingAreaList

        formdata = self.request.form
        try:
            if formdata.has_key('name'):

                sqlInstance = SqlObj()
                sqlStr = "SELECT `name`,`ID_number` FROM `teacher_basic_info`"
                sqlResult = sqlInstance.execSql(sqlStr)
                checkExist = False
                for item in sqlResult:
                    if [ formdata['name'].decode('utf8'), formdata['ID_number'] ] == item.values():
                        checkExist =  True
                if not checkExist:
                    mana_teacher = api.content.get(path='/cshm/resource/mana_teacher')
                    obj = api.content.create(type = 'Teacher', title = formdata['name'], container = mana_teacher)
                    
                    sqlInstance = SqlObj()
                    sqlStr = "SELECT * FROM `teacher_basic_info`"
                    sqlResult = sqlInstance.execSql(sqlStr)
                    teacherTbColumn = sqlResult[0].keys()
                    columnStr = ''
                    for i, item in enumerate(teacherTbColumn):
                        if i != 0 : 
                            columnStr += ', '
                        columnStr += '`{}`'.format(item.encode('utf-8'))
                    
                    teacherTbColumn.remove('uid')
                    teacherTbColumn.remove('path')
                    dataStr = ''
                    for i, item in enumerate(teacherTbColumn):
                        if i != 0:
                            dataStr += ', '
                        if item in ['gender', 'birthday', 'working_day', 'isEnabled', 'isBaseEditLock', 'isCourseEdieLock']:
                            dataStr += "{}".format(formdata.get(item, '')) if formdata.get(item, '') != '' else 'NULL'
                        elif item == 'teaching_area':
                            dataStr += "'"
                            for area in formdata.get(item, ''):
                                dataStr += "{},".format(area) if area != '' else ''
                            dataStr += "'"
                        else:
                            dataStr += "'{}'".format(formdata.get(item, ''))
                    dataStr += ", '{}', '{}'".format(obj.UID(), obj.absolute_url_path())
                            
                    sqlinstance = SqlObj()
                    sqlStr = """insert into `teacher_basic_info`({}) values ({})""".format(columnStr, dataStr)
                    sqlInstance.execSql(sqlStr)
                else:
                    api.portal.show_message(message='此姓名及身份證字號已存在'.decode('utf8'), request=self.request, type='error')
        except Exception as e:
            import pdb;pdb.set_trace()
        return self.template()


class CreateTeacher(BrowserView):

    template = ViewPageTemplateFile("template/empty_view.pt")

    def __call__(self):
        import pdb;pdb.set_trace()
        request = self.request
        alsoProvides(request, IDisableCSRFProtection)
        teacherList = []
        errorItem = {}
        with open('/home/kiel/table.csv') as file:
            reader = csv.reader(file)
            for row in reader:
                teacherList.append(row)
            teacherList.pop(0)
            import pdb;pdb.set_trace()
            for teacher in teacherList:
                try:
                    checkExist = False
                    sqlInstance = SqlObj()
                    name = teacher[4] if teacher[4]!='-' and teacher[4]!='*' and teacher[4]!='無' else ''
                    id_num = teacher[3] if teacher[3]!='-' and teacher[3]!='*' and teacher[3]!='無' else ''
                    sqlStr = """SELECT name FROM `teacher_basic_info` WHERE name LIKE '{}' AND `ID_number` LIKE '{}' """.format(name, id_num)
                    sqlResult = sqlInstance.execSql(sqlStr)
                    if sqlResult != []:
                        checkExist = True
                    if not checkExist:
                        teacher[7] = teacher[7].replace('-', '').replace('**', '')
                        teacher[8] = teacher[8].replace('-', '').replace('**', '')
                        teacher[9] = teacher[9].replace('-', '').replace('**', '')

                        graduation  = str(int(teacher[20])+1911) + '-' + teacher[19] if teacher[20]!='' and int(teacher[20])!=0 else''
                        graduation2 = str(int(teacher[24])+1911) + '-' + teacher[25] if teacher[25]!='' and int(teacher[25])!=0 else''
                        graduation3 = str(int(teacher[29])+1911) + '-' + teacher[30] if teacher[29]!='' and int(teacher[29])!=0 else''
                        birthday ='"'+str(datetime.date(int(teacher[7])+1911, int(teacher[8]), int(teacher[9])))+'"' if teacher[8]!='' and int(teacher[8])!=0 else 'NULL'
                        working_day = '"'+str(datetime.datetime.strptime(teacher[34], "%Y年%m月%d日 %H時%M分%S秒").date()) +'"'if teacher[34]!='' else 'NULL'
                        sqlInstance = SqlObj()
                        sqlStr = "SELECT * FROM `teaching_area`"
                        sqlResult = sqlInstance.execSql(sqlStr)
                        areaStr = ''
                        areaNum = ''
                        priority_areaNum = ''
                        for i in range(40,54):
                            areaStr += teacher[i]
                        areaStr += teacher[65]
                        for k, v in sqlResult:
                            if areaStr.find(v.encode('utf-8')) != -1:
                                areaNum += str(k) + ','
                            if teacher[64].find(v.encode('utf-8')) != -1:
                                priority_areaNum = k
                        areaNum = areaNum if areaNum!='' else ''
                        priority_areaNum = priority_areaNum if priority_areaNum!='' else ''
                        
                        gender = int(teacher[5]) if teacher[5] != '' else 'NULL'

                        for i, item in enumerate(teacher):
                            teacher[i] = teacher[i].replace("'", '')
                            if item == '無' or item == '' or item == '-' or item == '**':
                                teacher[i] = ''
                        sqlinstance = SqlObj()
                        
                        mana_teacher = api.content.get(path='/cshm/resource/mana_teacher')
                        obj = api.content.create(type = 'Teacher', title = teacher[4], container = mana_teacher)
                    
                        sqlStr = """insert into `teacher_basic_info`(
                        `teacher_id`, `teacher_num`, `username`, `password`, `id_number`, `name`, `gender`, `birthday`, `telephone`, `cellphone`, `fax`, `email`, `residence_addr`, `mailing_addr`, `education`, `department`, `degree`, `graduation`, `education2`, `department2`, `degree2`, `graduation2`, `education3`, `department3`, `degree3`, `graduation3`, `service_units`, `service_department`, `present_job`, `working_day`, `units_phone`, `units_addr`, `experience`, `license`, `teaching_area`, `priority_area`, `recommender`, `recommended_way`, `teacher_modify`, `union_modify`, `isenabled`, `isbaseeditlock`, `iscourseedielock`, `disable_reason`, `review_status`, `review_note`, `filer`, `uid`, `path`) values 
                        ('{}','{}','{}','{}','{}','{}', {} , {}, '{}','{}',  '{}','{}','{}','{}','{}','{}','{}','{}','{}','{}',
                         '{}','{}','{}','{}','{}','{}','{}','{}','{}', {},   '{}','{}','{}','{}','{}','{}','{}','{}','{}','{}',
                          {},  {},  {}, '{}','{}','{}','{}','{}','{}')""".format(
                        teacher[0],  teacher[6],  teacher[1],  teacher[2],   teacher[3],  teacher[4] , gender,      birthday,     
                        teacher[10], teacher[11], teacher[12], teacher[13],  teacher[14], teacher[15], teacher[16], teacher[17],  
                        teacher[18], graduation,  teacher[21], teacher[23],  teacher[22], graduation2, teacher[26], teacher[28],  
                        teacher[27], graduation3, teacher[31], teacher[32],  teacher[33], working_day, teacher[35], teacher[38], 
                        teacher[36], teacher[37], areaNum, priority_areaNum, teacher[62], teacher[63], teacher[54], teacher[55],  
                        teacher[56], teacher[57], teacher[61], teacher[58], teacher[59], teacher[60],  teacher[39], 
                        obj.UID(),     obj.absolute_url_path() )
                        sqlInstance.execSql(sqlStr)
                    else:
                        print teacher[4] + 'already creade in database~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
                except Exception as e:
                    errorItem.update({teacher[4]: e})
        import pdb;pdb.set_trace()
        errorMessage = ''
        for k, v in errorItem.iteritems():
            errorMessage += '{}:  {}\n'.format(k,v)
        if errorItem == {}:
            errorMessage = '已成功新增所有資料'
        self.errorMessage = errorMessage
        return self.template()
