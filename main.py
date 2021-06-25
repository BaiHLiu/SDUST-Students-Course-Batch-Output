'''
Description: 批量导出山东科技大学指定学号范围的学生课表
Author: Catop
Date: 2021-06-25 08:42:15
LastEditTime: 2021-06-25 09:52:06
'''

import xlrd
import xlwt
import json
import requests


######################################
#查询接口地址（必须校内网或走代理）
API_END_POINT = "http://192.168.111.167:8081/api/v3/students"
#配置关注词（导出时优先放在前面）
IMPORTANT_COURSES = [
    "体育"
]
#开始时间
START_TIME = "2020-09-10"
#截止时间
END_TIME = "2021-02-01"
#输入文件名
SRC_FILE_NAME = 'templete.xls'
#输出文件名
DEST_FILE_NAME = 'out.xls'
######################################


def read_xls(file_name):
    """读取学生名单"""
    srcFile = xlrd.open_workbook(file_name)
    table = srcFile.sheets()[0]
    students_id_list = table.col_values(0)
    
    return students_id_list

def get_student_info(student_id):
    """查询学生课表"""
    headers = {
        "Connection" : "keep-alive",
        "Accept-Encoding" : "gzip, deflate, br",
        "deviceSn" : "aa3d119f9eda",
        "X-Consumer-Custom-ID" : "sdust"
    }
    params = {
        "start" : START_TIME,
        "end" : END_TIME
    }
    url = API_END_POINT+f"/{student_id}/timetable-items"
    res = requests.get(url=url, headers=headers, params=params)
    
    info_list = {}
    if(res.status_code == 200):
        if not(res.text == '[]'):
            info_list = json.loads(res.text)
    
    #原始数据太多，做一步简化
    lite_list =  []

    for course_info in info_list:
        lite_info = {
        "teacherName" : course_info['teacherName'],
        "roomName" : course_info['roomName'],
        "dayOfWeek" : course_info['dayOfWeek'],
        "courseName" : course_info['courseName']
        }

        lite_list.append(lite_info)

    return lite_list


def sort_student_course(course_lite_list):
    """按要求归类单个学生的课程，老师和课程名称均相同则算一门课"""
    #注意根据要求，上下学期同一门可需要分别统计，而接口返回的就是类似"大学英语（A)(2-1)"，"大学英语（A)(2-2)"，可以直接用
    sorted_course_list = []
    #列表中字符串格式："张三 高等数学（A）（2-2）"

    for course_info in course_lite_list:
        course_str = course_info['teacherName'] + " " + course_info['courseName']
        if not(course_str in sorted_course_list):
            sorted_course_list.append(course_str)

    return sorted_course_list


def work(students_id_list, file_name):
    """导入学生列表，拉取每个人所有课程，保存结果输出xls"""
    des_file = xlwt.Workbook(encoding = 'utf-8')
    worksheet = des_file.add_sheet('Sheet1')

    err_cont = 0
    ok_cont = 0

    for idx,student_id in enumerate(students_id_list):
        try:
            student_course_list = sort_student_course(get_student_info(student_id))
            #默认排序规则
            student_course_list.sort()
        except:
            print(f"第[{idx+1}/{len(students_id_list)}] 拉取失败，学号：{student_id}")
        else:
            print(f"第[{idx+1}/{len(students_id_list)}] 成功!")
            worksheet.write(idx, 0, label=str(student_id))
            
            col_idx = 1

            #匹配关注词课程，放在表格最前面
            for imp_word in IMPORTANT_COURSES:
                imp_courses = [x for i,x in enumerate(student_course_list) if x.find(imp_word) != -1]
                for course_str in imp_courses:
                    worksheet.write(idx, col_idx, label=course_str)
                    col_idx += 1
                    student_course_list.remove(course_str)
            
            #写其他课程
            for course_str in student_course_list:
                worksheet.write(idx, col_idx, label=course_str)
                col_idx += 1

    des_file.save(file_name)

if __name__ == "__main__":
    #read_xls('templete.xls')
    #course_list = get_student_info('202001060912')
    #print(sort_student_course(course_list))
    work(read_xls(SRC_FILE_NAME),DEST_FILE_NAME)