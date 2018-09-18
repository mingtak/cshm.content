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
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
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

        table = document.add_table(rows=1, cols=7, style = 'Table Grid')

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


class StudentSingUpReport(Basic):
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


class EmergencyGradeReport(Basic):
    def __call__(self):
        result = self.getAllStudent()
        context = self.context
        document = Document()

        course = context.getParentNode().Title()
        period = context.id

        title = '中國勞工安全衛生管理學會%s訓練班第%s期學員學、術科成績冊' %(course, period)
        style1 = document.styles.add_style('style1', 1)
        style1.font.size = Pt(24)
        document.add_paragraph(title.decode(), style = style1)
        document.add_paragraph('核准開班文號：中華民國107年8月15日 北市勞職字第1076066693號'.decode())

        start = self.getChineseBirthday(context.courseStart)
        end = self.getChineseBirthday(context.courseEnd)
        time = '訓練期間：%s~%s' %(start, end)

        document.add_paragraph(time.decode())
        document.add_paragraph(('訓練地點：%s' % context.trainingCenter.to_object.title).decode())
        document.add_paragraph('輔導員：'.decode())
        document.add_paragraph('監考人員：'.decode())
        document.add_page_break()

        count = 1
        table = document.add_table(rows=2, cols=11)
        table.rows[0].cells[0].merge(table.rows[1].cells[0]).text = '座號'.decode()
        table.rows[0].cells[1].merge(table.rows[1].cells[1]).text = '姓名'.decode()

        table.rows[0].cells[2].text = '術科'.decode()
        table.rows[1].cells[2].text = '敷料與繃帶'.decode()
        table.rows[0].cells[3].text = '術科'.decode()
        table.rows[1].cells[3].text = '骨骼及肌肉損傷'.decode()
        table.rows[0].cells[4].text = '術科'.decode()
        table.rows[1].cells[4].text = '傷患處理及搬運'.decode()
        table.rows[0].cells[5].text = '術科'.decode()
        table.rows[1].cells[5].text = '中毒、窒息緊急甦醒術'.decode()
        table.rows[0].cells[6].text = '術科'.decode()
        table.rows[1].cells[6].text = ''.decode()
        table.rows[0].cells[7].merge(table.rows[1].cells[7]).text = '總分'.decode()
        table.rows[0].cells[8].merge(table.rows[1].cells[8]).text = '平均'.decode()
        table.rows[0].cells[9].merge(table.rows[1].cells[9]).text = '學科'.decode()
        table.rows[0].cells[10].merge(table.rows[1].cells[10]).text = '備註'.decode()
        for item in result:
            obj = dict(item)
            if count in [11, 21, 31, 41, 51, 60, 70 ,80]:
                table = document.add_table(rows=2, cols=11)
                table.rows[0].cells[0].merge(table.rows[1].cells[0]).text = '座號'.decode()
                table.rows[0].cells[1].merge(table.rows[1].cells[1]).text = '姓名'.decode()
                table.rows[0].cells[2].text = '術科'.decode()
                table.rows[1].cells[2].text = '敷料與繃帶'.decode()
                table.rows[0].cells[3].text = '術科'.decode()
                table.rows[1].cells[3].text = '骨骼及肌肉損傷'.decode()
                table.rows[0].cells[4].text = '術科'.decode()
                table.rows[1].cells[4].text = '傷患處理及搬運'.decode()
                table.rows[0].cells[5].text = '術科'.decode()
                table.rows[1].cells[5].text = '中毒、窒息緊急甦醒術'.decode()
                table.rows[0].cells[6].text = '術科'.decode()
                table.rows[1].cells[6].text = ''.decode()
                table.rows[0].cells[7].merge(table.rows[1].cells[7]).text = '總分'.decode()
                table.rows[0].cells[8].merge(table.rows[1].cells[8]).text = '平均'.decode()
                table.rows[0].cells[9].merge(table.rows[1].cells[9]).text = '學科'.decode()
                table.rows[0].cells[10].merge(table.rows[1].cells[10]).text = '備註'.decode()

            row_cells = table.add_row().cells
            row_cells[0].text = str(obj['seatNo']).decode()
            row_cells[1].text = obj['name'].decode()
            row_cells[2].text = ''
            row_cells[3].text = ''
            row_cells[4].text = ''
            row_cells[5].text = ''
            row_cells[6].text = ''
            row_cells[7].text = ''
            row_cells[8].text = ''
            row_cells[9].text = ''
            row_cells[10].text = ''
            count += 1
            if count in [11, 21, 31, 41, 51, 60, 70 ,80]:
                document.add_page_break()

        document.save('/tmp/%s.docx' %title)

        return self.downloadFile(title)



class StudentSingInReport(Basic):
    def __call__(self):
        result = self.getAllStudent()
        context = self.context
        document = Document()

        course = context.getParentNode().Title()
        period = context.id
        start = self.getChineseBirthday(context.courseStart)
        end = self.getChineseBirthday(context.courseEnd)
        time = '訓練期間：%s~%s' %(start, end)

        document.add_paragraph('中國勞工安全衛生管理學會'.decode())
        document.add_paragraph(('第%s期【%s】訓練班' %(period, course)).decode())
        document.add_paragraph('核准開班文號：'.decode())
        document.add_paragraph(time.decode())
        document.add_paragraph(('訓練地點：%s' % context.trainingCenter.to_object.title).decode())
        document.add_paragraph('輔導員：'.decode())
        document.add_paragraph('監考人員：'.decode())

        count = 1
        table = document.add_table(rows=1, cols=9)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '序號'.decode()
        hdr_cells[1].text = '姓名'.decode()
        hdr_cells[2].text = '生日日期'.decode()
        hdr_cells[3].text = '身份證字號'.decode()
        hdr_cells[4].text = '學歷'.decode()
        hdr_cells[5].text = '服務單位'.decode()
        hdr_cells[6].text = '聯絡地址'.decode()
        hdr_cells[7].text = '電話'.decode()
        hdr_cells[8].text = '備註'.decode()

        for item in result:
            obj = dict(item)
            if count in [6, 11, 16, 21, 26, 31, 36, 41, 46, 51, 56, 61, 66]:
                table = document.add_table(rows=1, cols=9)
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = '序號'.decode()
                hdr_cells[1].text = '姓名'.decode()
                hdr_cells[2].text = '生日日期'.decode()
                hdr_cells[3].text = '身份證字號'.decode()
                hdr_cells[4].text = '學歷'.decode()
                hdr_cells[5].text = '服務單位'.decode()
                hdr_cells[6].text = '聯絡地址'.decode()
                hdr_cells[7].text = '電話'.decode()
                hdr_cells[8].text = '備註'.decode()

            row_cells = table.add_row().cells
            row_cells[0].text = str(count)
            row_cells[1].text = obj['name'].decode()
            row_cells[2].text = self.getChineseBirthday(obj['birthday']).decode()
            row_cells[3].text = obj['studId'].decode()
            row_cells[4].text = obj['degree'].decode()
            row_cells[5].text = obj['company_name'].decode()
            row_cells[6].text = obj['address'].decode()
            row_cells[7].text = obj['phone'].decode()
            note = ''
            if obj['training_hour']:
                note += '%s<br>' %obj['training_hour']
            note += '發證單位TODO<br>'
            note += '發證時間TODO<br>'
            note += '發證字號TODO<br>'

            row_cells[8].text = note.decode()
            count += 1
            if count in [6, 11, 16, 21, 26, 31, 36, 41, 46, 51, 56, 61, 66]:
                document.add_page_break()
        title = '第%s期【%s】訓練班-報班' %(period, course)
        document.save('/tmp/%s.docx' %title)

        return self.downloadFile(title)

