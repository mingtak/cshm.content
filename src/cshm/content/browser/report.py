# -*- coding: utf-8 -*-
from cshm.content import _
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from mingtak.ECBase.browser.views import SqlObj
import json
import datetime
import math
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_BREAK


class Basic(BrowserView):

    def getAllStudent(self):
        uid = self.context.UID()
        sqlInstance = SqlObj()
        sqlStr = """SELECT *, training_status_code.status FROM `reg_course`, `training_status_code` WHERE uid = '{}' and isAlt = 0 and 
                 seatNo IS NOT null and reg_course.training_status = training_status_code.id""".format(uid)
        return sqlInstance.execSql(sqlStr)

    def getChineseBirthday(self, birthday):
        year = birthday.year - 1911
        month = birthday.month
        day = birthday.day
        return '%s/%s/%s' %(year, month, day)

    def downloadFile(self, title):
        self.request.response.setHeader('Content-Type', 'application/docx')
        self.request.response.setHeader(
            'Content-Disposition',
            'attachment; filename="%s.docx"' %title
        )
        with open('/tmp/%s.docx' %title, 'rb') as file:
            docs = file.read()
            return docs


class StudentListReport(Basic):
    def __call__(self):
        result = self.getAllStudent()
        context = self.context
        document = Document()

        course = context.getParentNode().Title()
        period = context.id

        title = '中國勞工安全衛生管理學會%s訓練班第%s期學員名冊' %(course, period)

        style1 = document.styles.add_style('style1', 1)
        style1.font.size = Pt(24)
	style1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        document.add_paragraph(title.decode(), style = style1)

        start = self.getChineseBirthday(context.courseStart)
        end = self.getChineseBirthday(context.courseEnd)
        time = '訓練期間：%s~%s' %(start, end)
        trainingCenter = context.trainingCenter.to_object.title
        document.add_paragraph('核准開班文號：中華民國107年3月20日北市消預字第10731924800號'.decode())
        document.add_paragraph(time.decode())
        document.add_paragraph(('學科：%s' %trainingCenter).decode())
        document.add_paragraph(('術科：%s' %trainingCenter).decode())

        document.add_page_break()

        table = document.add_table(rows=1, cols=7)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '編號'.decode()
        hdr_cells[1].text = '姓名'.decode()
        hdr_cells[2].text = '生日'.decode()
        hdr_cells[3].text = '身份證字號'.decode()
        hdr_cells[4].text = '服務單位'.decode()
        hdr_cells[5].text = '聯絡地址'.decode()
        hdr_cells[6].text = '備註'.decode()

        count = 1
        number = 0
        for item in result:
            obj = dict(item)
            if count in [9, 17, 25, 33, 41, 49, 57, 65, 73, 81]:
                table = document.add_table(rows=1, cols=7)
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = '編號'.decode()
                hdr_cells[1].text = '姓名'.decode()
                hdr_cells[2].text = '生日'.decode()
                hdr_cells[3].text = '身份證字號'.decode()
                hdr_cells[4].text = '服務單位'.decode()
                hdr_cells[5].text = '聯絡地址'.decode()
                hdr_cells[6].text = '備註'.decode()

            row_cells = table.add_row().cells
            row_cells[0].text = str(count)
            row_cells[1].text = obj['name'].decode()
            row_cells[2].text = self.getChineseBirthday(obj['birthday']).decode()
            row_cells[3].text = obj['studId'].decode()
            row_cells[4].text = obj['company_name'].decode()
            row_cells[5].text = obj['address'].decode()
            row_cells[6].text = obj['status'].decode()

            count += 1
            number += 1

            if count in [9, 17, 25, 33, 41, 49, 57, 65, 73, 81]:
                document.add_page_break()
        document.add_page_break()
        document.save('/tmp/%s.docx' %title)

        return self.downloadFile(title)


class StudentSingReport(Basic):
    def __call__(self):
        result = self.getAllStudent()
        context = self.context
        document = Document()

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


class StudentGradeReport(Basic):
    def __call__(self):
        result = self.getAllStudent()
        context = self.context
        document = Document()

        course = context.getParentNode().Title()
        period = context.id

        title = '中國勞工安全衛生管理學會%s訓練班第%s期學員成績冊' %(course, period)
        style1 = document.styles.add_style('style1', 1)
        style1.font.size = Pt(24)
        document.add_paragraph(title.decode(), style = style1)

        count = 1
        table = document.add_table(rows=1, cols=7)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '編號'.decode()
        hdr_cells[1].text = '姓名'.decode()
        hdr_cells[2].text = '生日'.decode()
        hdr_cells[3].text = '身份證字號'.decode()
        hdr_cells[4].text = '服務單位'.decode()
        hdr_cells[5].text = '分數'.decode()
        hdr_cells[6].text = '備註'.decode()

        for item in result:
            obj = dict(item)
            if count in [9, 17, 25, 33, 41, 49, 57, 65, 73, 81]:
                table = document.add_table(rows=1, cols=7)
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = '編號'.decode()
                hdr_cells[1].text = '姓名'.decode()
                hdr_cells[2].text = '生日'.decode()
                hdr_cells[3].text = '身份證字號'.decode()
                hdr_cells[4].text = '服務單位'.decode()
                hdr_cells[5].text = '分數'.decode()
                hdr_cells[6].text = '備註'.decode()

            row_cells = table.add_row().cells
            row_cells[0].text = str(count)
            row_cells[1].text = obj['name'].decode()
            row_cells[2].text = self.getChineseBirthday(obj['birthday']).decode()
            row_cells[3].text = obj['studId'].decode()
            row_cells[4].text = obj['company_name'].decode()
            row_cells[5].text = ''
            row_cells[6].text = obj['status'].decode()
            count += 1
            if count in [9, 17, 25, 33, 41, 49, 57, 65, 73, 81]:
                document.add_page_break()

        document.save('/tmp/%s.docx' %title)

        return self.downloadFile(title)


class StudentIssueReport(Basic):
    def __call__(self):
        result = self.getAllStudent()
        context = self.context
        document = Document()

        course = context.getParentNode().Title()
        period = context.id

        title = '中國勞工安全衛生管理學會%s訓練班第%s期學員核發清冊' %(course, period)
        style1 = document.styles.add_style('style1', 1)
        style1.font.size = Pt(24)
        document.add_paragraph(title.decode(), style = style1)

        count = 1
        table = document.add_table(rows=1, cols=7)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '編號'.decode()
        hdr_cells[1].text = '姓名'.decode()
        hdr_cells[2].text = '生日'.decode()
        hdr_cells[3].text = '身份證字號'.decode()
        hdr_cells[4].text = '服務單位'.decode()
        hdr_cells[5].text = '分數'.decode()
        hdr_cells[6].text = '備註'.decode()

        for item in result:
            obj = dict(item)
            if count in [9, 17, 25, 33, 41, 49, 57, 65, 73, 81]:
                table = document.add_table(rows=1, cols=7)
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = '編號'.decode()
                hdr_cells[1].text = '姓名'.decode()
                hdr_cells[2].text = '生日'.decode()
                hdr_cells[3].text = '身份證字號'.decode()
                hdr_cells[4].text = '服務單位'.decode()
                hdr_cells[5].text = '分數'.decode()
                hdr_cells[6].text = '備註'.decode()

            row_cells = table.add_row().cells
            row_cells[0].text = str(count)
            row_cells[1].text = obj['name'].decode()
            row_cells[2].text = self.getChineseBirthday(obj['birthday']).decode()
            row_cells[3].text = obj['studId'].decode()
            row_cells[4].text = obj['company_name'].decode()
            row_cells[5].text = ''
            row_cells[6].text = obj['status'].decode()
            count += 1
            if count in [9, 17, 25, 33, 41, 49, 57, 65, 73, 81]:
                document.add_page_break()

        document.save('/tmp/%s.docx' %title)

        return self.downloadFile(title)

