# -*- coding: utf-8 -*-
from cshm.content import _
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from mingtak.ECBase.browser.views import SqlObj
import json
import datetime
import math
from docxtpl import DocxTemplate, Listing
from Products.CMFPlone.utils import safe_unicode
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_BREAK
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


class Basic(BrowserView):

    def getAllStudent(self):
        uid = self.context.UID()
        sqlInstance = SqlObj()
        sqlStr = """SELECT *, training_status_code.status FROM `reg_course`, `training_status_code`, education_code WHERE uid = '{}' 
                 and isAlt = 0 and seatNo IS NOT null and reg_course.training_status = training_status_code.id AND
                 reg_course.education_id = education_code.id ORDER BY reg_course.seatNo""".format(uid)
        return sqlInstance.execSql(sqlStr)

    def getChineseBirthday(self, birthday):
        if birthday:
            year = birthday.year - 1911
            month = birthday.month
            day = birthday.day
            return '%s/%s/%s' %(year, month, day)
        else:
            return ''

    def downloadFile(self, title):
        self.request.response.setHeader('Content-Type', 'application/docx')
        self.request.response.setHeader(
            'Content-Disposition',
            'attachment; filename="%s.docx"' %title
        )
        with open('/tmp/temp.docx', 'rb') as file:
            docs = file.read()
            return docs


class GernalReport(Basic):
    def __call__(self):
        result = self.getAllStudent()
        context = self.context
        request = self.request
        mode = request.get('mode')
        course = context.getParentNode().Title()
        period = context.id

        if mode == 'horizon':
            filePath = '/home/andy/cshm/zeocluster/src/cshm.content/src/cshm/content/browser/static/horizon.docx'
            title = '%s_%s_橫式' %(course, period)
        elif mode == 'straight':
            filePath = '/home/andy/cshm/zeocluster/src/cshm.content/src/cshm/content/browser/static/straight.docx'
            title = '%s_%s_直式' %(course, period)
        elif mode == 'fire':
            filePath = '/home/andy/cshm/zeocluster/src/cshm.content/src/cshm/content/browser/static/fire.docx'
            title = '%s_%s_防火' %(course, period)
        elif mode == 'issue':
            filePath = '/home/andy/cshm/zeocluster/src/cshm.content/src/cshm/content/browser/static/issue.docx'
            title = '%s_%s_核發清冊' %(course, period)
        elif mode == 'emergency_grade':
            filePath = '/home/andy/cshm/zeocluster/src/cshm.content/src/cshm/content/browser/static/emergency_grade.docx'
            title = '%s_%s_成績冊' %(course, period)

        doc = DocxTemplate(filePath)

        start = self.getChineseBirthday(context.courseStart)
        end = self.getChineseBirthday(context.courseEnd)
        time = '訓練期間：%s~%s' %(start, end)
        trainingCenter = context.trainingCenter.to_object.title

        parameter = {}
        parameter['course'] = safe_unicode(course)
        parameter['period'] = safe_unicode(period)
        parameter['time'] = safe_unicode(time)
        parameter['trainingCenter'] = safe_unicode(trainingCenter)
        temp = []
        for item in result:
            obj = dict(item)
            obj['note'] = ''
            obj['birthday'] = self.getChineseBirthday(obj['birthday'])

            if course == '職業安全管理師' or course == '職業衛生管理師':
                license_date = self.getChineseBirthday(obj['license_date'])
                if obj['license_unit']:
                    obj['note'] += obj['license_unit'] + '\n'
                if license_date:
                    obj['note'] += license_date + '\n'
                if obj['license_code']:
                    obj['note'] += obj['license_code'] + '\n'
            elif re.match('防火', course):
                pass
            else:
                obj['note'] += obj['status'] + '\n'

            temp.append(obj)

        data = []
        count = 0
        length = len(result)
        while length > 0:
            data.append(temp[count * 8: (count + 1) * 8])
            length -= 8
            count += 1
        parameter['data'] = data
        doc.render(parameter)
        doc.save("/tmp/temp.docx")

        return self.downloadFile(title)


class StudentSingInReport(Basic):
    def __call__(self):
        result = self.getAllStudent()
        context = self.context
        document = Document()
        sections = document.sections 
        for section in sections:
            # change orientation to landscape
            section.orientation = WD_ORIENT.LANDSCAPE
        new_width, new_height = section.page_height, section.page_width
        section.page_width = new_width
        section.page_height = new_height

        course = context.getParentNode().Title()
        period = context.id

        title = '中國勞工安全衛生管理學會%s訓練班第%s期學員簽到' %(course, period)
        style1 = document.styles.add_style('style1', 1)
        style1.font.size = Pt(24)
        document.add_paragraph(title.decode(), style = style1)

        count = 1
        table = document.add_table(rows=1, cols=8)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '編號'.decode()
        hdr_cells[1].text = '姓名'.decode()
        hdr_cells[2].text = '上午簽到'.decode()
        hdr_cells[3].text = '下午簽到'.decode()
        hdr_cells[4].text = '編號'.decode()
        hdr_cells[5].text = '姓名'.decode()
        hdr_cells[6].text = '上午簽到'.decode()
        hdr_cells[7].text = '下午簽到'.decode()

        for item in result:
            obj = dict(item)
            if count in [17, 33, 49, 65, 81, 97]:
                table = document.add_table(rows=1, cols=8)
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = '編號'.decode()
                hdr_cells[1].text = '姓名'.decode()
                hdr_cells[2].text = '上午簽到'.decode()
                hdr_cells[3].text = '下午簽到'.decode()
                hdr_cells[4].text = '編號'.decode()
                hdr_cells[5].text = '姓名'.decode()
                hdr_cells[6].text = '上午簽到'.decode()
                hdr_cells[7].text = '下午簽到'.decode()

            row_cells = table.add_row().cells
            row_cells[0].text = str(count)
            row_cells[1].text = obj['name'].decode()
            row_cells[4].text = str(count + 1)
            row_cells[5].text = obj['name'].decode()
            count += 2
            if count in [17, 33, 49, 65, 81, 97]:
                document.add_page_break()

        document.save('/tmp/%s.docx' %title)

        return self.downloadFile(title)

