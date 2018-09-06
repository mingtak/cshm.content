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

    def getUidCourseData(self, uid):
        """抓符合UID的reg_course資料"""
        sqlStr = """SELECT reg_course.*,training_status_code.status,education_code.degree FROM reg_course,training_status_code,
                    education_code WHERE uid = '{}' and reg_course.training_status = training_status_code.id AND
                    reg_course.education_id = education_code.id
                    """.format(uid)
        sqlInstance = SqlObj()

        return sqlInstance.execSql(sqlStr)

    def getCourseCode(self, course):
        courseList = {
            '職業安全管理師教育訓練': '01010',
            '職業衛生管理師教育訓練': '01020',
            '職業安全衛生管理員教育訓練': '01030',
            '甲種職業安全衛生業務主管安全衛生教育訓練': '02010',
            '乙種職業安全衛生業務主管安全衛生教育訓練': '02020',
            '丙種職業安全衛生業務主管安全衛生教育訓練': '02030',
            '吊升荷重在三公噸以上之固定式起重機及吊升荷重在一公噸以上之斯達卡式起重機操作人員安全衛生教育訓練': '03010',
            '吊升荷重在三公噸以上之固定式起重機(架空型─機上操作)操作人員': '03011',
            '吊升荷重在三公噸以上之固定式起重機(架空型─地面操作)操作人員': '03012',
            '吊升荷重在三公噸以上之固定式起重機(伸臂型)操作人員': '03013',
            '吊升荷重在三公噸以上之移動式起重機操作人員安全衛生教育訓練': '03020',
            '吊升荷重在三公噸以上之移動式起重機(伸臂可伸縮式)操作人員': '03021',
            '吊升荷重在三公噸以上之移動式起重機(伸臂不可伸縮式)操作人員': '03022',
            '吊升荷重在三公噸以上之人字臂起重桿操作人員安全衛生教育訓練': '03030',
            '吊籠操作人員安全衛生教育訓練': '03040',
            '(停用)吊升荷重在一公噸以上之斯達卡式起重機操作人員安全衛生教育訓練': '03050',
            '導軌或升降路之高度在二十公尺以上之營建用提升機操作人員安全衛生教育訓練': '03060',
            '鍋爐操作人員安全衛生教育訓練': '04010',
            '甲級鍋爐操作人員安全衛生教育訓練': '04011',
            '乙級鍋爐操作人員安全衛生教育訓練': '04012',
            '丙級鍋爐操作人員安全衛生教育訓練': '04013',
            '第一種壓力容器操作人員安全衛生教育訓練': '04020',
            '高壓氣體特定設備操作人員安全衛生教育訓練': '04030',
            '高壓氣體容器操作人員安全衛生教育訓練': '04040',
            '高壓氣體製造安全主任安全衛生教育訓練': '05040',
            '高壓氣體製造安全作業主管安全衛生教育訓練': '05050',
            '高壓氣體供應及消費作業主管安全衛生教育訓練': '05060',
            '擋土支撐作業主管安全衛生教育訓練': '06010',
            '模板支撐作業主管安全衛生教育訓練': '06020',
            '隧道等挖掘作業主管安全衛生教育訓練': '06030',
            '隧道等襯砌作業主管安全衛生教育訓練': '06040',
            '施工架組配作業主管安全衛生教育訓練': '06050',
            '鋼構組配作業主管安全衛生教育訓練': '06060',
            '露天開挖作業主管安全衛生教育訓練': '06070',
            '(停用)施工架及施工構台組配作業主管安全衛生教育訓練': '06080',
            '屋頂作業主管安全衛生教育訓練': '06090',
            '有機溶劑作業主管安全衛生教育訓練': '07010',
            '鉛作業主管安全衛生教育訓練': '07020',
            '四烷基鉛作業主管安全衛生教育訓練': '07030',
            '缺氧作業主管安全衛生教育訓練': '07040',
            '特定化學物質作業主管安全衛生教育訓練': '07050',
            '粉塵作業主管安全衛生教育訓練': '07060',
            '高壓室內作業主管安全衛生教育訓練': '07070',
            '潛水作業主管安全衛生教育訓練': '07080',
            '小型鍋爐操作人員特殊安全衛生教育訓練': '08010',
            '荷重在一公噸以上之堆高機操作人員特殊安全衛生教育訓練': '08020',
            '吊升荷重在零點五公噸以上未滿三公噸之固定式起重機操作人員特殊安全衛生教育訓練': '08030',
            '吊升荷重在零點五公噸以上未滿三公噸之移動式起重機操作人員特殊安全衛生教育訓練': '08040',
            '吊升荷重在零點五公噸以上未滿三公噸之人字臂起重桿操作人員特殊安全衛生教育訓練': '08050',
            '使用起重機具從事吊掛作業人員特殊安全衛生教育訓練': '08060',
            '以乙炔熔接裝置或氣體集合裝置從事金屬之熔接、切斷或加熱作業人員特殊安全衛生教育訓練': '08080',
            '火藥爆破作業人員特殊安全衛生教育訓練': '08100',
            '胸高直徑七十公分以上之伐木作業人員特殊安全衛生教育訓練': '08120',
            '機械集材運材作業人員特殊安全衛生教育訓練': '08130',
            '高壓室內作業人員特殊安全衛生教育訓練': '08140',
            '潛水作業人員特殊安全衛生教育訓練': '08150',
            '油輪清艙作業人員特殊安全衛生教育訓練': '08160',
            '急救人員安全衛生教育訓練': '10010',
            '甲級化學性因子作業環境監測人員安全衛生教育訓練': '11010',
            '甲級物理性因子作業環境監測人員安全衛生教育訓練': '11020',
            '乙級化學性因子作業環境監測人員教育訓練': '11030',
            '乙級物理性因子作業環境監測人員安全衛生教育訓練': '11040',
            '施工安全評估人員安全衛生教育訓練': '12010',
            '製程安全評估人員安全衛生教育訓練': '12020',
            '(停用)營造業流動勞工安全衛生教育訓練': '13010',
            '營造業甲種職業安全衛生業務主管教育訓練': '14010',
            '營造業乙種職業安全衛生業務主管教育訓練': '14020',
            '營造業丙種職業安全衛生業務主管教育訓練': '14030',
            '勞工健康服務護理人員安全衛生教育訓練': '15010'
        }

        for k,v in courseList.items():
            if re.match(course, k.decode()):
                return v


class RegisterExcelClass(Basic):
    def __call__(self):
        request = self.request
        response = request.response
        context = self.context
        output = StringIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet('Sheet1')
        trainingCenter =  context.trainingCenter.to_object

        trainingCenterCode = trainingCenter.code
        courseCode = self.getCourseCode(context.getParentNode().title)
        trainingCenterAddress = trainingCenter.address
        trainingCenterPhone = trainingCenter.phone
        trainingCenterFax = trainingCenter.fax

        counselor = context.counselor
        classroom = context.classroom
        if classroom == 'first_room':
            classroom = '一'
        elif classroom == 'second_room':
            classroom = '＝'
        elif classroom == 'Taoyuan':
            classroom = '桃園'
        elif classroom == 'Zhongli':
            classroom = '中壢'
        elif classroom == 'Tainan':
            classroom = '台南'
        elif classroom == 'Kaohsiung':
            classroom = '高雄'

        normal_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'top',
        })

        worksheet.merge_range('A1:B1', '訓練單位(請填代碼)：', normal_format)
        worksheet.merge_range('C1:H1', trainingCenterCode, normal_format)
        worksheet.merge_range('A2:B2', '訓練種類：', normal_format)
        worksheet.merge_range('C2:H2', courseCode, normal_format)
        worksheet.merge_range('D2:E2', '期別', normal_format)
        worksheet.write('F2', '第', normal_format)
        worksheet.write('G2', context.id, normal_format)
        worksheet.write('H2', '期', normal_format)

        worksheet.merge_range('A3:B3', '訓練場所地址：', normal_format)
        worksheet.merge_range('C3:H3', trainingCenterAddress, normal_format)
        worksheet.merge_range('A4:B4', '教室名稱：', normal_format)
        worksheet.write('C4', '第', normal_format)
        worksheet.merge_range('D4:G4', classroom, normal_format)
        worksheet.write('H4', '教室', normal_format)

        worksheet.merge_range('A5:B5', '輔導員姓名：', normal_format)
        worksheet.merge_range('C5:H5', counselor, normal_format)
        worksheet.merge_range('A6:B6', '電話：', normal_format)
        worksheet.merge_range('C6:H6', trainingCenterPhone, normal_format)
        worksheet.merge_range('A7:B7', '傳真：', normal_format)
        worksheet.merge_range('C7:H7', trainingCenterFax, normal_format)

        worksheet.merge_range('A8:A9', '日期', normal_format)
        worksheet.merge_range('B8:B9', '星期', normal_format)
        worksheet.merge_range('C8:C9', '時間', normal_format)
        worksheet.merge_range('D8:D9', '課程名稱', normal_format)
        worksheet.merge_range('E8:E9', '時數', normal_format)
        worksheet.merge_range('F8:F9', '講師姓名', normal_format)
        worksheet.merge_range('G8:G9', '講師編號(無講師編號者，敘明符合規定之資格條款)', normal_format)
        worksheet.merge_range('H8:H9', '備註', normal_format)
        worksheet.set_column('A:A', 20)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 30)
        worksheet.set_column('G:G', 80)

        count = 10
        courseList = {}

        for course in context.getChildNodes():
            startDateTime = course.startDateTime
            courseList[startDateTime] = course

        courseList = sorted(courseList.items(), key=lambda x: x[0])
        for item in courseList:
            course = item[1]
            hour = course.hours
            startDateTime = item[0]
            courseTime = '%s~%s' %(startDateTime.strftime('%H%M'), (startDateTime + datetime.timedelta(hours=int(hour))).strftime('%H%M'))
            teacher = course.teacher.to_object

            worksheet.write('A%s' %count, startDateTime.strftime('%Y%m%d'), normal_format)
            worksheet.write('B%s' %count, startDateTime.isoweekday(), normal_format)
            worksheet.write('C%s' %count, courseTime, normal_format)
            worksheet.write('D%s' %count, course.title, normal_format)
            worksheet.write('E%s' %count, hour, normal_format)
            worksheet.write('F%s' %count, teacher.title, normal_format)
            worksheet.write('G%s' %count, teacher.teacherSN if teacher.teacherSN else '符合附表15(2項6款)', normal_format)
            worksheet.write('H%s' %count, '', normal_format)
            count += 1

        workbook.close()

        response.setHeader('Content-Type', 'application/xls')
        response.setHeader('Content-Disposition', 'attachment; filename="class.xlsx"')
        return output.getvalue()


class RegisterExcelTeacher(Basic):
   def __call__(self):
        request = self.request
        response = request.response
        context = self.context
        output = StringIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet('Sheet1')

        trainingCenterCode = context.trainingCenter.to_object.code
        courseCode = self.getCourseCode(context.getParentNode().title)
        normal_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'top',
        })
        worksheet.set_column('A:A', 10)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 15)
        worksheet.set_column('D:D', 40)
        worksheet.set_column('E:E', 35)
        worksheet.set_column('F:F', 25)

        worksheet.merge_range('A1:B1', '訓練單位(請填代碼)：', normal_format)
        worksheet.merge_range('C1:F1', trainingCenterCode, normal_format)
        worksheet.merge_range('A2:B2', '訓練種類：', normal_format)
        worksheet.merge_range('C2:F2', courseCode, normal_format)
        worksheet.write('A3', '期別：', normal_format)
        worksheet.write('B3', '第', normal_format)
        worksheet.merge_range('C3:E3', context.id, normal_format)
        worksheet.write('F3', '期', normal_format)

        worksheet.write('A4', '編號', normal_format)
        worksheet.write('B4', '講師姓名', normal_format)
        worksheet.write('C4', '學歷', normal_format)
        worksheet.write('D4', '經歷', normal_format)
        worksheet.write('E4', '通訊處', normal_format)
        worksheet.write('F4', '電話', normal_format)
        worksheet.set_row(2, 25)
        worksheet.set_row(3, 25)
        worksheet.set_row(4, 25)
        worksheet.set_row(1, 30)
        count = 5
        number = 1
        already = []
        for item in context.getChildNodes():
            teacher = item.teacher.to_object
            title = teacher.title
            if title not in already:
                phone = teacher.cellPhone if teacher.cellPhone else teacher.homePhone
                worksheet.write('A%s' %count, number, normal_format)
                worksheet.write('B%s' %count, title, normal_format)
                worksheet.write('C%s' %count, 'TODO', normal_format)
                worksheet.write('D%s' %count, teacher.personExp, normal_format)
                worksheet.write('E%s' %count, teacher.contactAddr, normal_format)
                worksheet.write('F%s' %count, phone, normal_format)
                worksheet.set_row(count, 70)
                already.append(title)
                number += 1
                count += 1
        workbook.close()

        response.setHeader('Content-Type', 'application/xls')
        response.setHeader('Content-Disposition', 'attachment; filename="teacher.xlsx"')
        return output.getvalue()


class RegisterExcelGraduaction(Basic):
    def __call__(self):
        request = self.request
        response = request.response
        context = self.context

        uid = context.UID()
        data = self.getUidCourseData(uid)
        trainingCenterCode = context.trainingCenter.to_object.code
        courseCode = self.getCourseCode(context.getParentNode().title)
        period = context.id
        courseStart = context.courseStart.strftime('%Y-%m-%d')
        courseEnd = context.courseEnd.strftime('%Y-%m-%d')
        output = StringIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet('Sheet1')

        header_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'center',
            'font_color': '#000000',
            'bg_color': '#d7ffff',
        })
        normal_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'center',
            'font_color': 'black',
        })
        worksheet.set_column('A:A', 20)
        worksheet.set_column('C:C', 15)
        worksheet.set_column('I:I', 30)
        worksheet.set_column('K:L', 15)
        worksheet.set_column('L:L', 15)
        worksheet.set_column('J:J', 15)
        worksheet.set_column('D:D', 15)

        worksheet.write('A1', '訓練單位(請填代碼)：', header_format)
        worksheet.merge_range('B1:E1', trainingCenterCode, normal_format)

        worksheet.write('A2', '訓練種類(請填代碼)：', header_format)
        worksheet.merge_range('B2:E2', courseCode, normal_format)

        worksheet.write('A3', '期別：', header_format)
        worksheet.write('B3', '第', normal_format)
        worksheet.merge_range('C3:D3', period, normal_format)
        worksheet.write('E3', '期', normal_format)

        worksheet.write('A4', '(訓練單位)開班文號：', header_format)
        worksheet.merge_range('B4:E4', 'TODO', normal_format)

        worksheet.write('A5', '文號日期：', header_format)
        worksheet.merge_range('B5:E5', 'TODO', normal_format)

        worksheet.write('A6', '證書編號', header_format)
        worksheet.write('B6', '姓名', header_format)
        worksheet.write('C6', '出生日期', header_format)
        worksheet.write('D6', '身份證字號', header_format)
        worksheet.write('E6', '學歷', header_format)
        worksheet.write('F6', '畢業學校', header_format)
        worksheet.write('G6', '服務單位', header_format)
        worksheet.write('H6', '郵遞區號', header_format)
        worksheet.write('I6', '聯絡地址', header_format)
        worksheet.write('J6', '電話', header_format)
        worksheet.write('K6', '開訓日', header_format)
        worksheet.write('L6', '結訓日', header_format)
        worksheet.write('M6', '備註', header_format)

        count = 7
        for item in data:
            address = item[13] if item[13] else item[9]
            worksheet.write('A%s' %count, '', normal_format)
            worksheet.write('B%s' %count, item[4], normal_format)
            worksheet.write('C%s' %count, item[11], normal_format)
            worksheet.write('D%s' %count, item[14], normal_format)
            worksheet.write('E%s' %count, item[34], normal_format)
            worksheet.write('F%s' %count, item[25], normal_format)
            worksheet.write('G%s' %count, item[6], normal_format)
            worksheet.write('H%s' %count, item[27], normal_format)
            worksheet.write('I%s' %count, address, normal_format)
            worksheet.write('J%s' %count, item[1], normal_format)
            worksheet.write('K%s' %count, courseStart, normal_format)
            worksheet.write('L%s' %count, courseEnd, normal_format)
            worksheet.write('M%s' %count, '', normal_format)

            count += 1

        workbook.close()

        response.setHeader('Content-Type', 'application/xls')
        response.setHeader('Content-Disposition', 'attachment; filename="graduaction.xlsx"')
        return output.getvalue()


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

