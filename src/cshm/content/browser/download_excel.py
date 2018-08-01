# -*- coding: UTF-8 -*-
from Products.Five.browser import BrowserView
from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile
from plone import api
from plone.protect.auto import safeWrite
from mingtak.ECBase.browser.views import SqlObj
import json
import csv
import datetime
from StringIO import StringIO
import requests
import xlsxwriter
import sys
from io import BytesIO
from urllib import quote
import re
reload(sys)
sys.setdefaultencoding('utf8')


class Basic(BrowserView):
    def checkBrowser(self, course, period):
        request = self.request
        fileName = '%s-%s.xls' %(course, period)
        user_agent = request['HTTP_USER_AGENT']
        try:
            if re.search('MSIE', user_agent) or re.search('Edge', user_agent):
                fileName = quote(fileName)
        except:
            pass

        return fileName


class DownloadSatisfactionExcel(Basic):
    def __call__(self):
        request = self.request
        response = request.response

        count_A = []
        count_B = []
        count_C = []
        count_D = []
        count_E = []
        count_F = []
        envir_data = []
        space_data = []
        each_teacher_data = json.loads(request.get('each_teacher_data'))
        total_anw = []

        for item in request.get('count_A').split('[')[1].split(']')[0].split(','):
            count_A.append(int(item))
        for item in request.get('count_B').split('[')[1].split(']')[0].split(','):
            count_B.append(int(item))
        for item in request.get('count_C').split('[')[1].split(']')[0].split(','):
            count_C.append(int(item))
        for item in request.get('count_D').split('[')[1].split(']')[0].split(','):
            count_D.append(int(item))
        for item in request.get('count_E').split('[')[1].split(']')[0].split(','):
            count_E.append(int(item))
        for item in request.get('count_F').split('[')[1].split(']')[0].split(','):
            count_F.append(int(item))
        for item in request.get('space_data').split('[')[1].split(']')[0].split(','):
            space_data.append(int(item))
        for item in request.get('envir_data').split('[')[1].split(']')[0].split(','):
            envir_data.append(int(item))
        for item in request.get('total_anw').split('[')[1].split(']')[0].split(','):
            total_anw.append(int(item))

        count_F_flag = True if count_F.count(0) != 5 else False
        if count_F_flag:
            posList = ['A64', 'I64']
        else:
            posList = ['I48', 'A64']

        period = request.get('period')
        course = request.get('course')
        count_data = json.loads(request.get('count_data'))
        point_space = request.get('point_space')
        point_envir = request.get('point_envir')
        point_teacher = request.get('point_teacher')
        point_total = request.get('point_total')

        output = StringIO()
        #fileName = '%s-%s.xls' %(course.decode('utf-8'), period)
        workbook = xlsxwriter.Workbook(output)
        worksheet1 = workbook.add_worksheet('Sheet1')
        worksheet2 = workbook.add_worksheet('Sheet2')

        data = [
            ['非常满意', '满意', '尚可', '不满意', '非常不满意'],
            total_anw,
            count_A,
            count_B,
            count_C,
            count_D,
            count_E,
            count_F,
            envir_data,
            space_data,
        ]

        worksheet2.write_column('A1', data[0])
        worksheet2.write_column('B1', data[1])
        worksheet2.write_column('C1', data[2])
        worksheet2.write_column('D1', data[3])
        worksheet2.write_column('E1', data[4])
        worksheet2.write_column('F1', data[5])
        worksheet2.write_column('G1', data[6])
        worksheet2.write_column('H1', data[7])
        worksheet2.write_column('I1', data[8])
        worksheet2.write_column('J1', data[9])

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$5',
            'values':     '=Sheet2!$B$1:$B$5',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '講師總整體滿意度'})
        worksheet1.insert_chart('A1', chart_total)

        chart1 = workbook.add_chart({'type': 'pie'})
        chart1.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$5',
            'values':     '=Sheet2!$C$1:$C$5',
            'data_labels': {'percentage': True},
        })
        chart1.set_title({'name': '教學態度'})
        worksheet1.insert_chart('A16', chart1)

        chart2 = workbook.add_chart({'type': 'pie'})
        chart2.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$5',
            'values':     '=Sheet2!$D$1:$D$5',
            'data_labels': {'percentage': True},
        })
        chart2.set_title({'name': '教學方式能啟發學員'})
        worksheet1.insert_chart('I16', chart2)

        chart3 = workbook.add_chart({'type': 'pie'})
        chart3.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$5',
            'values':     '=Sheet2!$E$1:$E$5',
            'data_labels': {'percentage': True},
        })
        chart3.set_title({'name': '能依課程、教材、內容有進度、系統講授'})
        worksheet1.insert_chart('A32', chart3)

        chart4 = workbook.add_chart({'type': 'pie'})
        chart4.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$5',
            'values':     '=Sheet2!$F$1:$F$5',
            'data_labels': {'percentage': True},
        })
        chart4.set_title({'name': '講授易懂，實務化'})
        worksheet1.insert_chart('I32', chart4)

        chart5 = workbook.add_chart({'type': 'pie'})
        chart5.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$5',
            'values':     '=Sheet2!$G$1:$G$5',
            'data_labels': {'percentage': True},
        })
        chart5.set_title({'name': '上課音量、口音表達適當、清晰'})
        worksheet1.insert_chart('A48', chart5)

        if count_F_flag:
            chart6 = workbook.add_chart({'type': 'pie'})
            chart6.add_series({
                'name':       'Pie sales data',
                'categories': '=Sheet2!$A$1:$A$5',
                'values':     '=Sheet2!$H$1:$H$5',
                'data_labels': {'percentage': True},
            })
            chart6.set_title({'name': '提供技能檢定或考照之建議或協助'})
            worksheet1.insert_chart('I48', chart6)

        chart7 = workbook.add_chart({'type': 'pie'})
        chart7.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$5',
            'values':     '=Sheet2!$I$1:$I$5',
            'data_labels': {'percentage': True},
        })
        chart7.set_title({'name': '學習環境'})
        worksheet1.insert_chart(posList[0], chart7)

        chart8 = workbook.add_chart({'type': 'pie'})
        chart8.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$5',
            'values':     '=Sheet2!$J$1:$J$5',
            'data_labels': {'percentage': True},
        })
        chart8.set_title({'name': '訓練服務'})
        worksheet1.insert_chart(posList[1], chart8)

        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
        })
        merge_format2 = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_color': '#FFFFFF',
            'bg_color': '#000000'
        })

        worksheet1.merge_range('A80:C81', '第%s期' %period, merge_format2)
        worksheet1.merge_range('D80:I81', course, merge_format2)
        worksheet1.merge_range('J80:L81', '訓練班', merge_format2)
        worksheet1.merge_range('A83:L83', '總體權值分數', merge_format2)
        worksheet1.merge_range('A84:L84', point_total, merge_format)
        worksheet1.merge_range('A85:D85', '環境權值分數', merge_format2)
        worksheet1.merge_range('E85:H85', '輔導員權值分數', merge_format2)
        worksheet1.merge_range('I85:L85', '講師整體權值分數', merge_format2)
        worksheet1.merge_range('A86:D86', point_space, merge_format)
        worksheet1.merge_range('E86:H86', point_envir, merge_format)
        worksheet1.merge_range('I86:L86', point_teacher, merge_format)

        worksheet1.merge_range('A87:C87', '項目', merge_format2)
        worksheet1.merge_range('D87:F87', '講師', merge_format2)
        worksheet1.merge_range('G87:I87', '平均權值', merge_format2)
        worksheet1.merge_range('J87:L87', '權值分數', merge_format2)

        count = 1
        row = 88
        for k,v in count_data.items():
            worksheet1.merge_range('A%s:C%s' %(row, row), count, merge_format)
            worksheet1.merge_range('D%s:F%s' %(row, row), k, merge_format)
            worksheet1.merge_range('G%s:I%s' %(row, row), v, merge_format)
            worksheet1.merge_range('J%s:L%s' %(row, row), v * 20, merge_format)
            count += 1
            row += 1

        workbook.close()
        fileName = self.checkBrowser(course, period)

        response.setHeader('Content-Type', 'application/xls')
        response.setHeader('Content-Disposition', 'attachment; filename="%s"' %(fileName))
        return output.getvalue()
#        path = '/home/andy/cshm/zeocluster/%s-%s.xls' %(course.decode('utf-8'), period)
#        with open(path, 'rb') as file:
#            docs = file.read()
#            return docs


class DownloadManagerExcel(Basic):
    def __call__(self):
        request = self.request
        response = request.response
        data = json.loads(request.get('data'))
        course = request.get('course')
        period = request.get('period')

        output = StringIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet1 = workbook.add_worksheet('Sheet1')
        worksheet2 = workbook.add_worksheet('Sheet2')

        worksheet2.write_column('A1', data['2'].keys())
        worksheet2.write_column('B1', data['2'].values())
        worksheet2.write_column('C1', data['3'].keys())
        worksheet2.write_column('D1', data['3'].values())
        worksheet2.write_column('E1', data['4'].keys())
        worksheet2.write_column('F1', data['4'].values())
        worksheet2.write_column('G1', data['5'].keys())
        worksheet2.write_column('H1', data['5'].values())
        worksheet2.write_column('I1', data['6'].keys())
        worksheet2.write_column('J1', data['6'].values())
        worksheet2.write_column('K1', data['7'].keys())
        worksheet2.write_column('L1', data['7'].values())
        worksheet2.write_column('M1', data['8'].keys())
        worksheet2.write_column('N1', data['8'].values())
        worksheet2.write_column('O1', data['9'].keys())
        worksheet2.write_column('P1', data['9'].values())
        worksheet2.write_column('Q1', data['10'].keys())
        worksheet2.write_column('R1', data['10'].values())
        worksheet2.write_column('S1', data['11'].keys())
        worksheet2.write_column('T1', data['11'].values())
        worksheet2.write_column('U1', data['12'].keys())
        worksheet2.write_column('V1', data['12'].values())
        worksheet2.write_column('W1', data['13'].keys())
        worksheet2.write_column('X1', data['13'].values())
        worksheet2.write_column('Y1', data['14'].keys())
        worksheet2.write_column('Z1', data['14'].values())


        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$4',
            'values':     '=Sheet2!$B$1:$B$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '參訓目的'})
        worksheet1.insert_chart('A1', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$C$1:$C$4',
            'values':     '=Sheet2!$D$1:$D$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '年齡'})
        worksheet1.insert_chart('I1', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$E$1:$E$11',
            'values':     '=Sheet2!$F$1:$F$11',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '行業別'})
        worksheet1.insert_chart('A16', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$G$1:$G$5',
            'values':     '=Sheet2!$H$1:$H$5',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '您是如何知道本項訓練課程'})
        worksheet1.insert_chart('I16', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$I$1:$I$4',
            'values':     '=Sheet2!$J$1:$J$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '您選擇本中心得因素(複選)'})
        worksheet1.insert_chart('A31', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$K$1:$K$3',
            'values':     '=Sheet2!$L$1:$L$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '據您所知，職業安全衛生法之中央主管機關為何單位'})
        worksheet1.insert_chart('I31', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$M$1:$M$3',
            'values':     '=Sheet2!$N$1:$N$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '保障工作者健康及安全為下列合法之宗旨'})
        worksheet1.insert_chart('A46', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$O$1:$O$3',
            'values':     '=Sheet2!$P$1:$P$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '何者為符合資格之職業安全衛生管理員'})
        worksheet1.insert_chart('I46', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$Q$1:$Q$3',
            'values':     '=Sheet2!$R$1:$R$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '王先生受雇於OO建設有限公司'})
        worksheet1.insert_chart('A61', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$S$1:$S$3',
            'values':     '=Sheet2!$T$1:$T$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '職業安全衛生法已字103.7.3正式施行'})
        worksheet1.insert_chart('I61', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$U$1:$U$3',
            'values':     '=Sheet2!$V$1:$V$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '僱主對勞工實施必要之安全衛生教育訓練'})
        worksheet1.insert_chart('A76', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$W$1:$W$3',
            'values':     '=Sheet2!$X$1:$X$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '作業中有物體飛落致為害勞工之虞'})
        worksheet1.insert_chart('I76', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$Y$1:$Y$3',
            'values':     '=Sheet2!$Z$1:$Z$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '下列何項為高架作業'})
        worksheet1.insert_chart('A91', chart_total)

        workbook.close()

        fileName = self.checkBrowser(course, period)

        response.setHeader('Content-Type',  'application/x-xlsx')
        response.setHeader('Content-Disposition', 'attachment; filename="%s-%s.xlsx"' %(course, period))
        return output.getvalue()


class DownloadStackerExcel(Basic):
    def __call__(self):
        request = self.request
        response = request.response
        data = json.loads(request.get('data'))
        course = request.get('course')
        period = request.get('period')

        output = StringIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet1 = workbook.add_worksheet('Sheet1')
        worksheet2 = workbook.add_worksheet('Sheet2')

        worksheet2.write_column('A1', data['2'].keys())
        worksheet2.write_column('B1', data['2'].values())
        worksheet2.write_column('C1', data['3'].keys())
        worksheet2.write_column('D1', data['3'].values())
        worksheet2.write_column('E1', data['4'].keys())
        worksheet2.write_column('F1', data['4'].values())
        worksheet2.write_column('G1', data['5'].keys())
        worksheet2.write_column('H1', data['5'].values())
        worksheet2.write_column('I1', data['6'].keys())
        worksheet2.write_column('J1', data['6'].values())
        worksheet2.write_column('K1', data['7'].keys())
        worksheet2.write_column('L1', data['7'].values())
        worksheet2.write_column('M1', data['8'].keys())
        worksheet2.write_column('N1', data['8'].values())
        worksheet2.write_column('O1', data['9'].keys())
        worksheet2.write_column('P1', data['9'].values())

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$4',
            'values':     '=Sheet2!$B$1:$B$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '參訓目的'})
        worksheet1.insert_chart('A1', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$C$1:$C$4',
            'values':     '=Sheet2!$D$1:$D$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '年齡'})
        worksheet1.insert_chart('I1', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$E$1:$E$11',
            'values':     '=Sheet2!$F$1:$F$11',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '行業別'})
        worksheet1.insert_chart('A16', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$G$1:$G$5',
            'values':     '=Sheet2!$H$1:$H$5',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '您是如何知道本項訓練課程'})
        worksheet1.insert_chart('I16', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$I$1:$I$4',
            'values':     '=Sheet2!$J$1:$J$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '您選擇本中心得因素(複選)'})
        worksheet1.insert_chart('A31', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$K$1:$K$4',
            'values':     '=Sheet2!$L$1:$L$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '學歷'})
        worksheet1.insert_chart('I31', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$M$1:$M$3',
            'values':     '=Sheet2!$N$1:$N$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '有無汽車駕駛執照'})
        worksheet1.insert_chart('A46', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$O$1:$O$3',
            'values':     '=Sheet2!$P$1:$P$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '堆高機'})
        worksheet1.insert_chart('I46', chart_total)

        workbook.close()

        fileName = self.checkBrowser(course, period)

        response.setHeader('Content-Type',  'application/x-xlsx')
        response.setHeader('Content-Disposition', 'attachment; filename="%s-%s.xlsx"' %(course, period))
        return output.getvalue()


class DownloadCtypeExcel(Basic):
    def __call__(self):
        request = self.request
        response = request.response
        data = json.loads(request.get('data'))
        course = request.get('course')
        period = request.get('period')

        output = StringIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet1 = workbook.add_worksheet('Sheet1')
        worksheet2 = workbook.add_worksheet('Sheet2')

        worksheet2.write_column('A1', data['2'].keys())
        worksheet2.write_column('B1', data['2'].values())
        worksheet2.write_column('C1', data['3'].keys())
        worksheet2.write_column('D1', data['3'].values())
        worksheet2.write_column('E1', data['4'].keys())
        worksheet2.write_column('F1', data['4'].values())
        worksheet2.write_column('G1', data['5'].keys())
        worksheet2.write_column('H1', data['5'].values())
        worksheet2.write_column('I1', data['6'].keys())
        worksheet2.write_column('J1', data['6'].values())
        worksheet2.write_column('K1', data['7'].keys())
        worksheet2.write_column('L1', data['7'].values())
        worksheet2.write_column('M1', data['8'].keys())
        worksheet2.write_column('N1', data['8'].values())
        worksheet2.write_column('O1', data['9'].keys())
        worksheet2.write_column('P1', data['9'].values())

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$4',
            'values':     '=Sheet2!$B$1:$B$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '參訓目的'})
        worksheet1.insert_chart('A1', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$C$1:$C$4',
            'values':     '=Sheet2!$D$1:$D$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '年齡'})
        worksheet1.insert_chart('I1', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$E$1:$E$11',
            'values':     '=Sheet2!$F$1:$F$11',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '行業別'})
        worksheet1.insert_chart('A16', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$G$1:$G$5',
            'values':     '=Sheet2!$H$1:$H$5',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '您是如何知道本項訓練課程'})
        worksheet1.insert_chart('I16', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$I$1:$I$4',
            'values':     '=Sheet2!$J$1:$J$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '您選擇本中心得因素(複選)'})
        worksheet1.insert_chart('A31', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$K$1:$K$4',
            'values':     '=Sheet2!$L$1:$L$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '職業安全衛生法之中央主管機關為何單位'})
        worksheet1.insert_chart('I31', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$M$1:$M$3',
            'values':     '=Sheet2!$N$1:$N$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '勞動契約係以下列何種目的為正確'})
        worksheet1.insert_chart('A46', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$O$1:$O$3',
            'values':     '=Sheet2!$P$1:$P$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '僱用多少人以下之事業單位'})
        worksheet1.insert_chart('I46', chart_total)

        workbook.close()

        fileName = self.checkBrowser(course, period)

        response.setHeader('Content-Type',  'application/x-xlsx')
        response.setHeader('Content-Disposition', 'attachment; filename="%s-%s.xlsx"' %(course, period))
        return output.getvalue()


class DownloadEmergencyExcel(Basic):
    def __call__(self):
        request = self.request
        response = request.response
        data = json.loads(request.get('data'))
        course = request.get('course')
        period = request.get('period')

        output = StringIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet1 = workbook.add_worksheet('Sheet1')
        worksheet2 = workbook.add_worksheet('Sheet2')

        worksheet2.write_column('A1', data['2'].keys())
        worksheet2.write_column('B1', data['2'].values())
        worksheet2.write_column('C1', data['3'].keys())
        worksheet2.write_column('D1', data['3'].values())
        worksheet2.write_column('E1', data['4'].keys())
        worksheet2.write_column('F1', data['4'].values())
        worksheet2.write_column('G1', data['5'].keys())
        worksheet2.write_column('H1', data['5'].values())
        worksheet2.write_column('I1', data['6'].keys())
        worksheet2.write_column('J1', data['6'].values())
        worksheet2.write_column('K1', data['7'].keys())
        worksheet2.write_column('L1', data['7'].values())
        worksheet2.write_column('M1', data['8'].keys())
        worksheet2.write_column('N1', data['8'].values())
        worksheet2.write_column('O1', data['9'].keys())
        worksheet2.write_column('P1', data['9'].values())
        worksheet2.write_column('Q1', data['10'].keys())
        worksheet2.write_column('R1', data['10'].values())

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$A$1:$A$4',
            'values':     '=Sheet2!$B$1:$B$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '參訓目的'})
        worksheet1.insert_chart('A1', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$C$1:$C$4',
            'values':     '=Sheet2!$D$1:$D$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '年齡'})
        worksheet1.insert_chart('I1', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$E$1:$E$11',
            'values':     '=Sheet2!$F$1:$F$11',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '行業別'})
        worksheet1.insert_chart('A16', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$G$1:$G$5',
            'values':     '=Sheet2!$H$1:$H$5',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '您是如何知道本項訓練課程'})
        worksheet1.insert_chart('I16', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$I$1:$I$4',
            'values':     '=Sheet2!$J$1:$J$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '您選擇本中心得因素(複選)'})
        worksheet1.insert_chart('A31', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$K$1:$K$2',
            'values':     '=Sheet2!$L$1:$L$2',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '是否曾經從事醫護工作'})
        worksheet1.insert_chart('I31', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$M$1:$M$4',
            'values':     '=Sheet2!$N$1:$N$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '預觸電患者急救時應先'})
        worksheet1.insert_chart('A46', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$O$1:$O$4',
            'values':     '=Sheet2!$P$1:$P$4',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '可否移動患者'})
        worksheet1.insert_chart('I46', chart_total)

        chart_total = workbook.add_chart({'type': 'pie'})
        chart_total.add_series({
            'name':       'Pie sales data',
            'categories': '=Sheet2!$Q$1:$Q$3',
            'values':     '=Sheet2!$R$1:$R$3',
            'data_labels': {'percentage': True},
        })
        chart_total.set_title({'name': '應於幾分鐘內施予急救'})
        worksheet1.insert_chart('A61', chart_total)

        workbook.close()

        fileName = self.checkBrowser(course, period)

        response.setHeader('Content-Type',  'application/x-xlsx')
        response.setHeader('Content-Disposition', 'attachment; filename="%s-%s.xlsx"' %(course, period))
        return output.getvalue()

