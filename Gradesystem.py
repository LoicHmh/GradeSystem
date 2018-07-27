# -*- coding: utf-8 -*-

import pandas as pd
import os
import numpy as np
import math
import sys

import re
from PyQt5.QtWidgets import QApplication, QCheckBox, QWidget, QToolTip, QPushButton, \
    QMessageBox, QDesktopWidget, QMainWindow, QGridLayout, QRadioButton, QGroupBox, QVBoxLayout, QComboBox, QLabel, \
    QTableWidget, QLineEdit, QAction, QTableWidgetItem, QTextEdit, QFileDialog
from PyQt5.QtGui import QIcon, QFont, QBrush, QTextOption
from PyQt5.QtCore import Qt, QCoreApplication

'''Back part'''


########################################
# Rule for GPA
########################################
def credit_rule(score):
    if score >= 95 and score <= 100:
        return 4.3
    elif score >= 90 and score < 95:
        return 4.0
    elif score >= 85 and score < 90:
        return 3.7
    elif score >= 80 and score < 85:
        return 3.3
    elif score >= 75 and score < 80:
        return 3.0
    elif score >= 70 and score < 75:
        return 2.7
    elif score >= 67 and score < 70:
        return 2.3
    elif score >= 65 and score < 67:
        return 2.0
    elif score >= 62 and score < 65:
        return 1.7
    elif score >= 60 and score < 62:
        return 1.0
    else:
        return 0.0


########################################
# Models
########################################
class Config():
    ''' Some important configures.

        Arguments:
            cal_gpa         # 是否计算gpa， 会自动产生gpa排名
            cal_caa         # 是否计算学积分，会自动产生学积分排名
            sort_by_gpa     # 是否按照GPA排序
            sort_by_caa     # 是否按照学积分排序
            output_path     # 选择输出位置
            data_path       # 选择输入文件夹
            file_name       # 输出文件名

    '''

    def __init__(self, cal_gpa=True, cal_caa=True, sort_by_gpa=False, sort_by_caa=False, sort_by_major=False,
                 sort_by_source=False, student_list_path="../student_list/", output_path="../output/", file_name="test",
                 data_path="../data/"):
        self.cal_gpa = cal_gpa
        self.cal_caa = cal_caa
        self.sort_by_gpa = sort_by_gpa
        self.sort_by_caa = sort_by_caa
        self.sort_by_major = sort_by_major
        self.sort_by_source = sort_by_source
        self.output_path = output_path
        self.data_path = data_path
        self.file_name = file_name
        self.student_list_path = student_list_path
        if not self.file_name.split('.')[-1] == 'xlsx':
            self.file_name += '.xlsx'

    def show(self):
        print("=========configures============")
        print("Calculate GPA:", self.cal_gpa)
        print("Calculate CAA:", self.cal_caa)
        print("Sort by GPA:", self.sort_by_gpa)
        print("Sort by CAA:", self.sort_by_caa)
        print("Sort by major:", self.sort_by_major)
        print("Sort by source:", self.sort_by_source)
        print()
        print("Student list path", self.student_list_path)
        print("Data path:", self.data_path)
        print("Output path:", self.output_path)
        print("===============================")

    def set_data_path(self, data_path):
        self.data_path = data_path

    def set_output_path(self, output_path):
        self.output_path = output_path

    def set_student_list_path(self, student_list_path):
        self.student_list_path = student_list_path

    def set_cal_gpa(self, cal_gpa):
        self.cal_gpa = cal_gpa

    def set_cal_caa(self, cal_caa):
        self.cal_caa = cal_caa

    def set_file_name(self, file_name):
        self.file_name = file_name
        if not self.file_name.split('.')[-1] == 'xlsx':
            self.file_name += '.xlsx'


class Semester:
    ''' A semester contains the year of start, the year of end and its number.

        Arguments:
                        # 以2015-2016-2 为例
            year_start  # 2015
            year_end    # 2016
            number      # 2

    '''

    def __init__(self, year_start, year_end, number):
        self.year_start = year_start
        self.year_end = year_end
        self.number = number
        if not self.varify():
            print("The format of semester is wrong!")

    def varify(self):
        if int(self.year_start) == int(self.year_end) - 1 and 1 <= int(self.number) <= 2:
            return True
        else:
            return False

    def equals(self, other_semester):
        if (self.year_start == other_semester.year_start and
                self.year_end == other_semester.year_end and
                self.number == other_semester.number):
            return True
        else:
            return False

    def to_str(self):
        return str(self.year_start) + '-' + str(self.year_end) + '-' + str(self.number)


class Student:
    ''' Student infomation.

        Attributes：
            student_id:     # 学号， "51526****".
            student_name:   # 中文姓名，"张三".
            class_id:       # 班级号，"F1526002".
            grade_data:     # 学期成绩信息的列表， class Grades_data 的 list
    '''

    def __init__(self, student_name, student_id, student_year, class_id, major, source):
        self.student_id = student_id
        self.student_name = student_name
        self.student_year = student_year
        self.class_id = class_id
        self.major = major
        self.source = source
        self.msg = []

        self.grades_data = []
        np.seterr(divide='ignore', invalid='ignore')

    def clean_msg(self):
        self.msg = []

    def get_msg(self):
        return self.msg

    def add_msg(self, m):
        if m not in self.msg:
            self.msg.append(m)

    def get_student_name(self):
        return self.student_name

    def get_student_id(self):
        return self.student_id

    def get_class_id(self):
        return self.class_id

    def get_grades_data(self):
        return self.grades_data

    def get_major(self):
        return self.major

    def get_source(self):
        return self.source

    def get_student_year(self):
        return self.student_year

    def add_grades_data(self, grades_data):
        '''To add a grades_data.'''
        self.grades_data.append(grades_data)

    def find_grades_data(self, semester):
        '''To find the grades_data for a given semester.'''
        for gd in self.grades_data:
            if gd.get_semester().equals(semester):
                return gd
        return None

    def show(self):
        '''To show the student's all grades information.'''
        print(self.student_id, self.student_name, self.class_id)
        for gd in self.grades_data:
            gd.show()

    def calculate_gpa(self, semesters, return_credit=False):
        '''To calculate the student's GPA'''
        gpa_list, credit_list = [], []
        for semester in semesters:
            grades_data = self.find_grades_data(semester)
            if grades_data is not None:
                gpa, credit = grades_data.calculate_gpa(return_credit=True)
                gpa_list.append(gpa)
                credit_list.append(credit)
            else:
                print("没有找到 ", self.class_id, self.student_name, semester.to_str(), "学期成绩 ")
                self.add_msg("没有找到 {} {} {} 学期成绩".format(self.class_id, self.student_name, semester.to_str()))
        gpa = np.asarray(gpa_list)
        credit = np.asarray(credit_list)
        gpa_average = np.dot(gpa.T, credit) / credit.sum()
        if return_credit:
            return gpa_average, credit.sum()
        else:
            return gpa_average

    def calculate_caa(self, semesters, return_credit=False):
        '''To calculate the student's cumulative academic average.'''
        caa_list, credit_list = [], []
        for semester in semesters:
            grades_data = self.find_grades_data(semester)
            if grades_data is not None:
                caa, credit = grades_data.calculate_caa(return_credit=True)
                caa_list.append(caa)
                credit_list.append(credit)
            else:
                print("没有找到 ", self.class_id, self.student_name, semester.to_str(), "学期成绩 ")
                self.add_msg("没有找到 {} {} {} 学期成绩".format(self.class_id, self.student_name, semester.to_str()))
        caa = np.asarray(caa_list)
        credit = np.asarray(credit_list)
        caa_average = np.dot(caa.T, credit) / credit.sum()
        if return_credit:
            return caa_average, credit.sum()
        else:
            return caa_average


class Grades_data:
    '''Grade information for one semester.

        Attributes：
            semester    #学期， class Semester 的实例
            grades      #成绩列表， class Grade 的实例的list

    '''

    def __init__(self, semester, grades):
        self.semester = semester
        self.grades = grades

    def get_semester(self):
        return self.semester

    def get_grades(self):
        return self.grades

    def show(self):
        print(self.semester.to_str())
        for g in self.grades:
            g.show()

    def calculate_gpa(self, return_credit=False):
        scores = []
        credits = []
        for grade in self.grades:
            score = grade.get_student_grade()
            if score == 'P':
                continue
            score = credit_rule(float(score))
            scores.append(score)
            credits.append(float(grade.get_course_credit()))
        scores = np.asarray(scores)
        credits = np.asarray(credits)
        gpa = np.dot(scores.T, credits) / credits.sum()
        if return_credit:
            return gpa, credits.sum()
        else:
            return gpa

    def calculate_caa(self, return_credit=False):
        scores = []
        credits = []
        for grade in self.grades:
            score = grade.get_student_grade()
            if score == 'P':
                continue
            scores.append(float(score))
            credits.append(float(grade.get_course_credit()))
        scores = np.asarray(scores)
        credits = np.asarray(credits)
        acc = np.dot(scores.T, credits) / credits.sum()
        if return_credit:
            return acc, credits.sum()
        else:
            return acc


class Grade:
    ''' Grade for one course

    Arguments:
        course_name     # 课程名称
        course_credit   # 课程绩点
        student_grade   # 学生成绩

    '''

    def __init__(self, course_name, course_credit, student_grade):
        self.course_name = course_name
        self.course_credit = course_credit
        self.student_grade = student_grade

    def get_course_name(self):
        return self.course_name

    def get_course_credit(self):
        return self.course_credit

    def get_student_grade(self):
        return self.student_grade

    def show(self):
        print(self.course_name, self.course_credit, self.student_grade)


########################################
# Controler
########################################
class Controler:
    '''
    控制类
    '''

    def __init__(self, config):
        self.msg = []
        self.config = config
        self.data_path = self.config.data_path
        self.file_list = os.listdir(self.data_path)

        self.studentyear_semester_dic = {}
        self.data_list = []
        self.student_dic = {}

        self.read_files()
        self.init_student_dic()
        self.update()

    def read_files(self):
        # file_name exemple: F1526002-2015-2016-1.xls
        for file_name in self.file_list:
            # 检查file_name
            suffix = file_name.split(".")[-1]
            if suffix != "xls":
                print("无法读取", file_name, "文件!")
                self.add_msg("无法读取" + file_name + "文件!")
                self.file_list.remove(file_name)
                continue
            # 从班级号中识别学生的年级
            student_year = "20" + file_name[1:3]
            # 从file_name中识别学期
            fn_split_list = file_name.split('.')[0].split('-')
            semester = Semester(fn_split_list[1], fn_split_list[2], fn_split_list[3])
            semester = semester.to_str()
            if student_year in self.studentyear_semester_dic:
                if semester not in self.studentyear_semester_dic[student_year]:
                    self.studentyear_semester_dic[student_year].append(semester)
            else:
                self.studentyear_semester_dic[student_year] = [semester]
            self.data_list.append(self.load_xls(file_name))

    def clean_msg(self):
        self.msg = []

    def get_msg(self):
        return self.msg

    def add_msg(self, m):
        if m not in self.msg:
            self.msg.append(m)

    def check_data(self, student_year, semesters):
        flag = True
        msg = []
        if student_year in self.studentyear_semester_dic:
            for semester in semesters:
                if semester.to_str() not in self.studentyear_semester_dic[student_year]:
                    msg.append("缺少{}级学生在{}学期的成绩！".format(student_year, semester.to_str()))
                    flag = False
        else:
            flag = False
            msg.append("缺少{}级学生的成绩！".format(student_year))
        return flag, msg

    # 初始化学生信息
    def init_student_dic(self):
        student_list_dir = "../student_list"
        file_name_list = os.listdir(student_list_dir)
        student_info_dic = {}
        for file_name in file_name_list:
            file_path = os.path.join(student_list_dir, file_name)
            excel = pd.read_excel(file_path, sheet_name="录取结果")
            student_year = file_name[0:5]
            # print("#debug", "excel columns", excel.columns)
            # print("#debug", "excel index", excel.index)
            for i in range(len(excel["姓名"])):
                student_name = excel["姓名"][i]
                student_id = excel["学号"][i]
                class_id = excel["班级"][i]
                major = excel["录取专业"][i]
                source = excel["招生来源"][i]

                self.student_dic[student_name] = Student(student_name, student_id, student_year, class_id,
                                                         major, source)

    def load_xls(self, file_name):
        # TODO xls
        file_path = os.path.join(self.data_path, file_name)
        try:
            print(file_name)
            data = pd.read_html(file_path, encoding='utf-8')
            # 改columns名称
            data = data[0]
            columns = data[0:1].values[0]
            data = data[1:].values
            ans = pd.DataFrame(data, columns=columns)
            return ans
        except ValueError as e:
            print(file_name, "is wrong!")
            print("file_list:", len(self.file_list))
            # self.file_list.remove(file_name)
            print("file_list_apres:", len(self.file_list))
            ans = pd.read_excel(file_path, encoding='utf-8')
            print("ceshi \n", ans)
            return ans

    # 读取学生成绩信息并转换为数据结构
    def update(self):
        # 遍历所有xls文件
        print("len", len(self.data_list), len(self.file_list))
        for sheet, file_name in zip(self.data_list, self.file_list):
            print(file_name)
            # xls文件名上有 班级，学期信息
            fn = file_name.split('.')[0].split('-')
            _, year_start, year_end, number = fn
            semester = Semester(year_start, year_end, number)
            # 将sheet转换为numpy形式
            sheet_np = np.asarray(sheet)
            # sheet第一行有课程名称信息
            index = sheet_np[0]
            # 对sheet从第二行开始按行遍历

            for i in range(len(sheet["学号"])):
                # 获取学号、姓名、班级信息
                student_id = sheet["学号"][i]  # sheet[i, 0]
                student_name = sheet["姓名"][i]  # sheet[i, 1]
                class_id = sheet["班号"][i]  # sheet[i, 2]

                # 对这一行的每一列从第四列开始遍历，步长为2，因为有学分列
                grades = []
                for j in range(3, sheet_np.shape[1], 2):
                    # 如果为nan，说明数据为空，不读入
                    if sheet_np[i, j] is not np.nan:
                        # 新建一个Grade类型的数据，课程名，学分，该学生的成绩
                        grade = Grade(index[j], sheet_np[i, j + 1], sheet_np[i, j])
                        grades.append(grade)
                # 新建一个Grades_data类型的数据，学期，成绩list
                grades_data = Grades_data(semester, grades)

                # 检查这个学生是否已经存在
                if student_name not in self.student_dic:
                    print("缺少" + class_id + student_name + "的基本信息！")
                    self.add_msg("缺少" + class_id + student_name + "的基本信息！")
                    continue
                # 将这一学期的成绩添加到这个学生的数据中
                self.student_dic[student_name].add_grades_data(grades_data)

    # 返回学生信息字典{student_name, student}, 支持选择班级
    def get_student_dic(self, student_year, class_id=None, student_id=None):
        res_dic_raw = {}
        for student_name, student in self.student_dic.items():
            # print("#debug student_year", student.get_student_year(), "grade", student_year)
            if student.get_student_year()[:-1] == student_year[:-1]:
                res_dic_raw[student_name] = student

        res_dic = {}
        if class_id is not None:
            for student_name, student in res_dic_raw.items():
                if student.get_class_id() == class_id:
                    res_dic[student_name] = student
            return res_dic
        elif student_id is not None:
            for student_name, student in res_dic_raw.items():
                if student.get_student_id() == student_id:
                    res_dic[student_name] = student
            return res_dic
        else:
            return res_dic_raw

    # 得到某个同学的成绩数据
    # 返回 Grades_data 的 list
    def list_grades(self, student_name, show=False):
        if student_name in self.student_dic:
            student = self.student_dic[student_name]
            if show:
                student.show()
            return student.get_grades_data()
        else:
            print("没有找到 ", student_name)
            self.add_msg("没有找到 {}".format(student_name))
            return None

    # 输出Excel表格
    def write_excel(self, students, semesters, config, save=False):
        print("students", students)
        print("semesters", semesters)
        self.config = config
        # 输出路径
        if save:
            if not os.path.exists(self.config.output_path):
                os.mkdir(self.config.output_path)

            writer = pd.ExcelWriter(os.path.join(self.config.output_path, self.config.file_name))
        # 表头
        index = ["姓名", "学号", "班级", "专业", "招生来源"]
        if self.config.cal_gpa:
            index.append("GPA")
            index.append("GPA总学分")
        if self.config.cal_caa:
            index.append("学积分")
            index.append("学积分总学分")

        if self.config.sort_by_major and self.config.sort_by_source:
            group_dic = {}
            for student in students:
                if student.major + '-' + student.source not in group_dic:
                    group_dic[student.major + '-' + student.source] = [student]
                else:
                    group_dic[student.major + '-' + student.source].append(student)

            dfs = []
            for group_name, group in group_dic.items():
                print("1")
                lines = self.get_content(group, semesters)
                df = self.get_dataframe(lines, index)
                if save:
                    df.to_excel(writer, group_name)
                dfs.append(df)
            if save:
                writer.save()
            return dfs, self.get_msg()

        elif self.config.sort_by_major and not self.config.sort_by_source:
            group_dic = {}
            for student in students:
                if student.major not in group_dic:
                    group_dic[student.major] = [student]
                else:
                    group_dic[student.major].append(student)

            dfs = []
            for group_name, group in group_dic.items():
                print("2")
                lines = self.get_content(group, semesters)
                df = self.get_dataframe(lines, index)
                if save:
                    df.to_excel(writer, group_name)
                dfs.append(df)
            if save:
                writer.save()
            return dfs, self.get_msg()

        elif not self.config.sort_by_major and self.config.sort_by_source:
            group_dic = {}
            for student in students:
                if student.source not in group_dic:
                    group_dic[student.source] = [student]
                else:
                    group_dic[student.source].append(student)

            dfs = []
            for group_name, group in group_dic.items():
                print("3")
                lines = self.get_content(group, semesters)
                df = self.get_dataframe(lines, index)
                if save:
                    df.to_excel(writer, group_name)
                dfs.append(df)
            if save:
                writer.save()
            return dfs, self.get_msg()

        else:
            print("4", students, semesters)
            lines = self.get_content(students, semesters)
            df = self.get_dataframe(lines, index)
            if save:
                df.to_excel(writer, "sheet1")
                writer.save()
            return df, self.get_msg()

    def get_dataframe(self, lines, index):
        # 转换为pandas.Dataframe, 并排序
        print("lines:", lines)
        print("index", index)
        df = pd.DataFrame(data=lines, columns=index)
        if self.config.sort_by_gpa:
            df = df.sort_values(by=['GPA'], ascending=False)
        elif self.config.sort_by_caa:
            df = df.sort_values(by=['学积分'], ascending=False)
        else:
            df = df.sort_values(by=['学号'], ascending=True)

        print("==========config==============")
        self.config.show()
        # 排名数据
        if self.config.cal_gpa or self.config.cal_caa:
            rank = df.rank(axis=0, method='min', numeric_only=True, na_option='keep', ascending=False, pct=False)

            if self.config.cal_gpa:
                print(df)
                rank_gpa = [-1 for _ in range(len(rank['GPA']))]
                for x in range(len(rank['GPA'])):
                    if not math.isnan(rank['GPA'][x]):
                        rank_gpa[x] = int(rank['GPA'][x])
                rank_gpa = {'rank_gpa': rank_gpa}
                rank_gpa = pd.DataFrame(rank_gpa)
                df['GPA排名'] = rank_gpa
            if self.config.cal_caa:
                rank_caa = [-1 for _ in range(len(rank['学积分']))]
                for x in range(len(rank['学积分'])):
                    if not math.isnan(rank['学积分'][x]):
                        rank_caa[x] = int(rank['学积分'][x])
                rank_caa = {'rank_caa': rank_caa}
                rank_caa = pd.DataFrame(rank_caa)
                df['学积分排名'] = rank_caa
        print(df)
        return df

    def get_content(self, students, semesters):
        lines = []
        for student in students:
            line = []
            line.append(student.get_student_name())
            line.append(student.get_student_id())
            line.append(student.get_class_id())

            line.append(student.get_major())
            line.append(student.get_source())

            if self.config.cal_gpa:
                # student.clean_msg()
                gpa, credit_gpa = student.calculate_gpa(semesters, return_credit=True)
                line.append(gpa)
                line.append(credit_gpa)
            if self.config.cal_caa:
                # student.clean_msg()
                caa, credit_caa = student.calculate_caa(semesters, return_credit=True)
                line.append(caa)
                line.append(credit_gpa)

            student_msg = student.get_msg()
            if student_msg != []:
                for temp_msg in student_msg:
                    self.add_msg(temp_msg)
            lines.append(line)
        return lines

    def show(self):
        for s, g in self.student_dic.items():
            g.show()

    def show_data(self):
        print(self.data_list)

    def one_semester(self):
        return self.data_list[0]


'''UI part'''


class UI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.config = Config()
        self.resize(1000, 800)
        self.center()
        self.setWindowTitle("上海交大-巴黎高科卓越工程师学院学积分管理系统")

        font = QFont("song", 10)
        QToolTip.setFont(font)
        self.setToolTip('欢迎使用学积分管理系统')
        self.setWindowIcon(QIcon('icon.png'))

        menubar = self.menuBar()
        optionMenu = menubar.addMenu("选项")
        # studentFileOption = QAction("打开学生信息存放位置", self)
        # optionMenu.addAction(studentFileOption)
        # inputFileOption = QAction("打开成绩文件存放位置", self)
        # optionMenu.addAction(inputFileOption)
        toExcelOption = QAction("导出当前查询结果至Excel文件", self)
        optionMenu.addAction(toExcelOption)
        # outputFileOption = QAction("打开导出文件存放位置", self)
        # optionMenu.addAction(outputFileOption)
        optionMenu.triggered[QAction].connect(self.processTrigger)

        aboutMenu = menubar.addMenu("关于")
        howOption = QAction("使用说明", self)
        aboutMenu.addAction(howOption)
        contactOption = QAction("联系我们", self)
        aboutMenu.addAction(contactOption)
        aboutMenu.triggered[QAction].connect(self.aboutTrigger)

        self.createFileLayout()
        self.createSearchLayout()
        self.createPeriodLayout()
        self.createResultLayout()
        self.createMessageLayout()

        windowLayout = QVBoxLayout()
        windowLayout.addWidget(self.fileGroupBox)
        windowLayout.addWidget(self.searchGroupBox)
        windowLayout.addWidget(self.periodGroupBox)
        windowLayout.addWidget(self.resultGroupBox)
        windowLayout.addWidget(self.msgGroupBox)

        widget = QWidget()
        widget.setLayout(windowLayout)
        self.setCentralWidget(widget)
        self.show()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def processTrigger(self, process):
        if process.text() == "打开成绩文件存放位置":
            os.system("start explorer ..\\data")
        elif process.text() == "打开学生信息存放位置":
            os.system("start explorer ..\\student_list")
        elif process.text() == "导出当前查询结果至Excel文件":
            filePath, ok = QFileDialog.getSaveFileName(self,
                                                       "文件保存",
                                                       "../output/",
                                                       "Excel Files (*.xlsx)")
            tempPath = filePath.split('/')
            fileName = tempPath[-1]
            output_path = ""
            for i in range(len(tempPath) - 1):
                output_path += (tempPath[i] + '/')
            self.config.set_output_path(output_path)
            self.config.set_file_name(fileName)
            self.controler.write_excel(self.students.values(), self.semesters, self.config, save=True)

    def aboutTrigger(self, process):
        if process.text() == "使用说明":
            self.window = aboutWindow(["本软件使用过程应遵循以下步骤：",
                                       "1.打开“选项——打开学生信息存放位置”，放入学生信息文件。该文件应为excel格式，并将有招生来源和专业的sheet命名为录取结果。",
                                       "2.打开“选项——打开成绩文件存放位置”，放入成绩文件。文件应为excel格式，命名格式为“班级号-起始年份-结束年份-学期数”，如“F1526001-2015-2016-2”。",
                                       "3.在主界面选择相关选项后，点击查询即可查看结果。查询结果可导出为excel文件，可以打开“选项——导出当前查询结果至Excel文件”选择存放位置并命名文件。"])
            self.window.show()
        elif process.text() == "联系我们":
            self.window = aboutWindow(["软件使用过程中如有任何问题请联系开发者：",
                                       "郑勤锴 2363471925@qq.com",
                                       "胡敏浩 hmh0205@qq.com"])
            self.window.show()

    def createFileLayout(self):
        self.fileGroupBox = QGroupBox("信息存放位置")
        gridlayout = QGridLayout()

        self.lb_student_list = QLabel("学生信息存放位置：")
        self.studentEdit = QLineEdit()
        self.studentEdit.setReadOnly(True)
        self.studentEdit.setPlaceholderText(os.path.realpath(self.config.student_list_path))

        self.lb_note = QLabel("成绩信息存放位置：")
        self.noteEdit = QLineEdit()
        self.noteEdit.setReadOnly(True)
        self.noteEdit.setPlaceholderText(os.path.realpath(self.config.data_path))

        self.btn_open_student_list = QPushButton("打开", self)
        self.btn_open_student_list.setFixedWidth(100)
        self.btn_open_student_list.setToolTip("打开存放学生信息文件夹")
        self.btn_open_student_list.resize(self.btn_open_student_list.sizeHint())
        self.btn_open_student_list.clicked.connect(
            lambda: self.openfileClicked(os.path.realpath(self.config.student_list_path)))

        self.btn_edit_student_list = QPushButton("更改", self)
        self.btn_edit_student_list.setFixedWidth(100)
        self.btn_edit_student_list.setToolTip("更改存放学生信息文件夹")
        self.btn_edit_student_list.resize(self.btn_edit_student_list.sizeHint())
        self.btn_edit_student_list.clicked.connect(lambda: self.editfileClicked("student_list"))

        self.btn_open_note = QPushButton("打开", self)
        self.btn_open_note.setFixedWidth(100)
        self.btn_open_note.setToolTip("打开存放成绩信息文件夹")
        self.btn_open_note.resize(self.btn_open_note.sizeHint())
        self.btn_open_note.clicked.connect(lambda: self.openfileClicked(os.path.realpath(self.config.data_path)))

        self.btn_edit_note = QPushButton("更改", self)
        self.btn_edit_note.setFixedWidth(100)
        self.btn_edit_note.setToolTip("更改存放成绩信息文件夹")
        self.btn_edit_note.resize(self.btn_edit_note.sizeHint())
        self.btn_edit_note.clicked.connect(lambda: self.editfileClicked("note"))

        gridlayout.addWidget(self.lb_student_list, 0, 0)
        gridlayout.addWidget(self.studentEdit, 0, 1)
        gridlayout.addWidget(self.btn_open_student_list, 0, 2)
        gridlayout.addWidget(self.btn_edit_student_list, 0, 3)
        gridlayout.addWidget(self.lb_note, 1, 0)
        gridlayout.addWidget(self.noteEdit, 1, 1)
        gridlayout.addWidget(self.btn_open_note, 1, 2)
        gridlayout.addWidget(self.btn_edit_note, 1, 3)

        self.fileGroupBox.setLayout(gridlayout)

    def createSearchLayout(self):
        self.searchGroupBox = QGroupBox("基础查询")
        gridlayout = QGridLayout()

        self.lb_grade = QLabel("年级：")
        self.lb_grade.setFixedWidth(50)
        self.cb_grade = QComboBox()
        self.cb_grade.setObjectName("年级")
        for grade in range(2012, 2030):
            self.cb_grade.addItem(str(grade))

        self.lb_major = QLabel("专业：")
        self.lb_major.setFixedWidth(50)
        self.cb_major = QComboBox()
        self.cb_major.setObjectName("专业")
        self.cb_major.addItems(["不限专业", "IE信息工程", "ME机械工程", "EPE能源与动力工程"])

        self.lb_year = QLabel("学年：")
        self.lb_year.setFixedWidth(50)
        self.cb_year = QComboBox()
        self.cb_year.setObjectName("学年")
        for year in range(2012, 2030):
            self.cb_year.addItem(str(year) + "-" + str(year + 1))

        self.lb_semester = QLabel("学期：")
        self.lb_semester.setFixedWidth(50)
        self.cb_semester = QComboBox()
        self.cb_semester.setObjectName("学期")
        self.cb_semester.addItems(["第一学期", "第二学期"])

        self.lb_rank = QLabel("排名类型: ")
        self.cb_rank = QComboBox()
        self.cb_rank.setObjectName("排名类型")
        self.cb_rank.addItems(["学号排名", "GPA排名", "学积分排名"])

        self.btn_search = QPushButton("查询", self)
        self.btn_search.setFixedWidth(100)
        self.btn_search.setToolTip("选择好各选项后点击查询成绩")
        self.btn_search.resize(self.btn_search.sizeHint())
        self.btn_search.clicked.connect(self.searchClicked)

        self.lb_type = QLabel("招生来源: ")
        self.cb_type = QComboBox()
        self.cb_type.setObjectName("招生来源")
        self.cb_type.addItems(["不限来源", "法语", "工科试验班类（中外合作办学）"])

        self.lb_number = QLabel("学号：")
        self.numberEdit = QLineEdit()
        self.numberEdit.setPlaceholderText("请输入学号，如为空则默认查询全年级")
        self.numberEdit.setFixedWidth(300)

        gridlayout.addWidget(self.lb_grade, 0, 0)
        gridlayout.addWidget(self.cb_grade, 0, 1)
        gridlayout.addWidget(self.lb_major, 0, 2)
        gridlayout.addWidget(self.cb_major, 0, 3)
        gridlayout.addWidget(self.lb_year, 0, 4)
        gridlayout.addWidget(self.cb_year, 0, 5)
        gridlayout.addWidget(self.lb_semester, 0, 6)
        gridlayout.addWidget(self.cb_semester, 0, 7)
        gridlayout.addWidget(self.btn_search, 0, 8)
        gridlayout.addWidget(self.lb_type, 1, 0)
        gridlayout.addWidget(self.cb_type, 1, 1)
        gridlayout.addWidget(self.lb_rank, 1, 2)
        gridlayout.addWidget(self.cb_rank, 1, 3)
        gridlayout.addWidget(self.lb_number, 1, 4)
        gridlayout.addWidget(self.numberEdit, 1, 5)

        self.searchGroupBox.setLayout(gridlayout)

    def createPeriodLayout(self):
        self.periodGroupBox = QGroupBox("高级查询")
        gridlayout = QGridLayout()

        self.lb_year1 = QLabel("起始学年：")
        self.lb_year1.setFixedWidth(80)
        self.cb_year1 = QComboBox()
        self.cb_year1.setObjectName("起始学年")
        for year in range(2012, 2030):
            self.cb_year1.addItem(str(year) + "-" + str(year + 1))

        self.lb_semester1 = QLabel("学期：")
        self.lb_semester1.setFixedWidth(50)
        self.cb_semester1 = QComboBox()
        self.cb_semester1.setObjectName("学期")
        self.cb_semester1.addItems(["第一学期", "第二学期"])

        self.lb_year2 = QLabel("结束学年：")
        self.lb_year2.setFixedWidth(80)
        self.cb_year2 = QComboBox()
        self.cb_year2.setObjectName("结束学年")
        for year in range(2012, 2030):
            self.cb_year2.addItem(str(year) + "-" + str(year + 1))

        self.lb_semester2 = QLabel("学期：")
        self.lb_semester2.setFixedWidth(50)
        self.cb_semester2 = QComboBox()
        self.cb_semester2.setObjectName("学期")
        self.cb_semester2.addItems(["第一学期", "第二学期"])

        self.btn_search = QPushButton("查询", self)
        self.btn_search.setFixedWidth(100)
        self.btn_search.setToolTip('选择好时间段后点击查询成绩')
        self.btn_search.resize(self.btn_search.sizeHint())
        self.btn_search.clicked.connect(self.periodClicked)

        gridlayout.addWidget(self.lb_year1, 0, 0)
        gridlayout.addWidget(self.cb_year1, 0, 1)
        gridlayout.addWidget(self.lb_semester1, 0, 2)
        gridlayout.addWidget(self.cb_semester1, 0, 3)
        gridlayout.addWidget(self.lb_year2, 0, 4)
        gridlayout.addWidget(self.cb_year2, 0, 5)
        gridlayout.addWidget(self.lb_semester2, 0, 6)
        gridlayout.addWidget(self.cb_semester2, 0, 7)
        gridlayout.addWidget(self.btn_search, 0, 10)

        self.periodGroupBox.setLayout(gridlayout)

    def createResultLayout(self):
        self.resultGroupBox = QGroupBox("查询结果")
        gridlayout = QGridLayout()
        self.table = QTableWidget()
        self.table.setColumnCount(11)
        self.table.setRowCount(100)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tableClass = ["学号", "姓名", "班级", "招生来源", "专业", "GPA", "GPA总学分", "GPA排名",
                           "学积分", "学积分总学分", "学积分排名"]
        self.table.setHorizontalHeaderLabels(self.tableClass)
        for index in range(self.table.columnCount()):
            item = self.table.horizontalHeaderItem(index)
            item.setFont(QFont("song", 12, QFont.Bold))
            item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)

        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()

        gridlayout.addWidget(self.table)
        self.resultGroupBox.setLayout(gridlayout)

    def createMessageLayout(self):
        self.msgGroupBox = QGroupBox("查询信息")
        self.msgGroupBox.setMaximumHeight(200)
        gridlayout = QGridLayout()
        self.msgBox = QTextEdit()
        self.msgBox.setReadOnly(True)
        gridlayout.addWidget(self.msgBox)
        self.msgGroupBox.setLayout(gridlayout)

    def openfileClicked(self, file_path):
        os.system("start explorer " + file_path)

    def editfileClicked(self, path_type):
        if path_type == 'student_list':
            path = QFileDialog.getExistingDirectory(self, "更改学生信息文件夹")
            self.config.set_student_list_path(path)
            self.studentEdit.setPlaceholderText(os.path.realpath(self.config.student_list_path))
        elif path_type == 'note':
            path = QFileDialog.getExistingDirectory(self, "更改成绩信息文件夹")
            self.config.set_data_path(path)
            self.noteEdit.setPlaceholderText(os.path.realpath(self.config.data_path))

    def searchClicked(self):
        self.config = Config()
        self.controler = Controler(self.config)
        self.table.clearContents()
        grade = self.cb_grade.currentText() + "级"

        if self.numberEdit.text() != "":
            self.students = self.controler.get_student_dic(grade, student_id=int(self.numberEdit.text()))
        else:
            if self.cb_rank.currentText() == 'GPA排名':
                self.config.sort_by_gpa = True

            elif self.cb_rank.currentText() == '学积分排名':
                self.config.sort_by_caa = True

            if not self.cb_type.currentText() == '不限来源':
                self.config.sort_by_source = True

            if not self.cb_major.currentText() == '不限专业':
                self.config.sort_by_major = True

            self.students = self.controler.get_student_dic(grade)

        year = self.cb_year.currentText().split("-")
        if self.cb_semester.currentText() == "第一学期":
            semester = '1'
        else:
            semester = '2'
        semester1 = Semester(year[0], year[1], semester)
        self.semesters = [semester1]
        self.config.set_file_name(grade + year[0] + '-' + year[1] + '-' + semester)
        flag, warning = self.controler.check_data(grade[:-1], self.semesters)
        if not flag:
            self.showMsg(warning)
        else:
            data, msg = self.controler.write_excel(self.students.values(), self.semesters, self.config)
            self.showResult(data)
            self.showMsg(msg)

    def periodClicked(self):
        self.config = Config()
        self.controler = Controler(self.config)
        self.table.clearContents()
        grade = self.cb_grade.currentText() + "级"

        if self.numberEdit.text() != "":
            self.students = self.controler.get_student_dic(grade, student_id=int(self.numberEdit.text()))
        else:
            if self.cb_rank.currentText() == 'GPA排名':
                self.config.sort_by_gpa = True

            elif self.cb_rank.currentText() == '学积分排名':
                self.config.sort_by_caa = True

            if not self.cb_type.currentText() == '不限来源':
                self.config.sort_by_source = True

            if not self.cb_major.currentText() == '不限专业':
                self.config.sort_by_major = True

            self.students = self.controler.get_student_dic(grade)

        year1 = self.cb_year1.currentText().split("-")
        if self.cb_semester1.currentText() == "第一学期":
            st1 = '1'
        else:
            st1 = '2'
        year2 = self.cb_year2.currentText().split("-")
        if self.cb_semester2.currentText() == "第一学期":
            st2 = '1'
        else:
            st2 = '2'
        semester1 = Semester(year1[0], year1[1], st1)
        semester2 = Semester(year2[0], year2[1], st2)

        self.semesters = [semester1]
        if str(semester1.number) == '1':
            self.semesters.append(Semester(year1[0], year1[1], '2'))
        for y in range(int(year1[0]) + 1, int(year2[0])):
            self.semesters.append(Semester(y, y + 1, '1'))
            self.semesters.append(Semester(y, y + 1, '2'))
        if str(semester2.number) == '2':
            self.semesters.append(Semester(year2[0], year2[1], '1'))
        self.semesters.append(semester2)

        self.config.set_file_name(
            grade + year1[0] + '_' + year1[1] + '_' + st1 + '-' + year2[0] + '_' + year2[1] + '_' + st2)
        flag, warning = self.controler.check_data(grade[:-1], self.semesters)
        if not flag:
            self.showMsg(warning)
        else:
            data, msg = self.controler.write_excel(self.students.values(), self.semesters, self.config)
            self.showResult(data)
            self.showMsg(msg)

    def showResult(self, data):
        if len(data) == 6:
            for temp_data in data:
                if self.infoMatch(temp_data["招生来源"][0], self.cb_type.currentText(), 'source') \
                        and self.infoMatch(temp_data['专业'][0], self.cb_major.currentText(), 'major'):
                    data = temp_data
                    break
        elif len(data) == 3:
            for temp_data in data:
                if self.infoMatch(temp_data['专业'][0], self.cb_major.currentText(), 'major'):
                    data = temp_data
                    break
        elif len(data) == 2:
            for temp_data in data:
                if self.infoMatch(temp_data["招生来源"][0], self.cb_type.currentText(), 'source'):
                    data = temp_data
                    break

        self.table.clearContents()
        for col in range(data.shape[1]):
            templist = np.array(data[self.tableClass[col]]).tolist()
            for row in range(data.shape[0]):
                newItem = QTableWidgetItem(str(templist[row]))
                newItem.setFont(QFont("song", 12))
                newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                self.table.setItem(row, col, newItem)

        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()

    def infoMatch(self, string_1, string_2, type):
        if type == 'source':
            search_list = ['工科', '法语']
        elif type == 'major':
            search_list = ['IE', 'ME', 'EPE']
        else:
            return False
        for element in search_list:
            if (re.search(element, string_1) is not None) and (re.search(element, string_2) is not None):
                return True
        return False

    def showMsg(self, msg):
        self.msgBox.clear()
        for msgs in msg:
            if type(msgs).__name__ == 'str':
                self.msgBox.append(msgs)

    def rankType(self, btn):
        if btn.text() == "GPA排名":
            if btn.isChecked():
                self.rank = 'gpa'
        elif btn.text() == "学积分排名":
            if btn.isChecked():
                self.rank = 'caa'

    def closeEvent(self, QCloseEvent):
        reply = QMessageBox.question(self, "Message",
                                     "是否确认退出?", QMessageBox.Yes |
                                     QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            QCloseEvent.accept()
        else:
            QCloseEvent.ignore()


class aboutWindow(QWidget):
    def __init__(self, Info):
        super().__init__()
        self.init_window(Info)

    def init_window(self, Info):
        self.setWindowTitle('关于')
        self.resize(500, 400)
        self.aboutInfo = QTextEdit()
        self.aboutInfo.setReadOnly(True)
        for info in Info:
            self.aboutInfo.append(info)
        windowLayout = QVBoxLayout(self)
        windowLayout.addWidget(self.aboutInfo)
        self.setLayout(windowLayout)


if __name__ == '__main__':
    if not os.path.exists("..\\data"):
        os.makedirs("..\\data")
    if not os.path.exists("..\\output"):
        os.makedirs("..\\output")
    app = QApplication(sys.argv)
    ex = UI()
    sys.exit(app.exec_())