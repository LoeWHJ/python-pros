# dirname = os.path.dirname(__file__)
# plugin_path = os.path.join(dirname,'plugins','platforms')
# os.environ['QT_QPA_PLATFORM_PLUGIN_PATH']=plugin_path
# pyinstaller -F  search.py  --noconsole

from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QApplication, QLineEdit, QMainWindow, QPushButton,  QPlainTextEdit
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QRegExpValidator, QIntValidator, QDoubleValidator
from PyQt5.QtCore import QRegExp
# from PySide2.QtWidgets import QApplication, QLineEdit, QMainWindow, QPushButton,  QPlainTextEdit
# from PySide2.QtUiTools import QUiLoader
# from PySide2 import QtCore, QtGui
# from PySide2.QtGui import QRegExpValidator, QIntValidator, QDoubleValidator
# from PySide2.QtCore import QRegExp
import xlrd
import re
import time
import pypinyin
from pypinyin import Style
# from threading import Thread
# import synonyms
# from gensim.models.word2vec import Word2Vec,KeyedVectors
# import gensim
from fuzzywuzzy import fuzz
import jieba

# 全排列
from copy import deepcopy
def permutations(arr, position, end, h):
    if position == end:
        #    if position == end-1:
        h.append(deepcopy(arr))

    else:
        for index in range(position, end):
            arr[index], arr[position] = arr[position], arr[index]

            permutations(arr, position + 1, end,h)
            arr[index], arr[position] = arr[position], arr[index]

# 全组合
from copy import deepcopy
def per(arr, start, num,res,h):
    if num == 0:
        h.append(deepcopy(res))
        return
    if len(arr)-start<num:
        return
    else:
        res.append(arr[start])
        per(arr, start+1, num-1,res,h)
        temp=res.pop()
        per(arr, start+1, num,res,h)

from itertools import combinations
def combine(temp_list, n):
    '''根据n获得列表中的所有可能组合（n个元素为一组）'''
    temp_list2 = []
    for c in combinations(temp_list, n):
        temp_list2.append(c)
    return temp_list2


class Stats:
    def __init__(self):
        self.book = xlrd.open_workbook("jxmp.xlsx")
        # self.book = xlrd.open_workbook("jxmp_100.xlsx")
        print(f"包含表单数量 {self.book.nsheets}")
        print(f"表单的名分别为: {self.book.sheet_names()}")
        self.sheet = self.book.sheet_by_index(0)
        print(f"表单名：{self.sheet.name} ")
        print(f"表单索引：{self.sheet.number}")
        print(f"表单行数：{self.sheet.nrows}")
        print(f"表单列数：{self.sheet.ncols}")
        print(f"返回类型：{type(self.sheet.row_values(rowx=0))}")
        print(f"返回类型：{self.sheet.row_values(rowx=0)}")
        # 从文件中加载UI定义
        # 从 UI 定义中动态 创建一个相应的窗口对象
        # 注意：里面的控件对象也成为窗口对象的属性了
        # 比如 self.ui.button , self.ui.textEdit
        # self.ui = QUiLoader().load('gui.ui')
        self.ui = loadUi('gui.ui')

        self.ui.lcdNumber.setStyleSheet('background-color:#cbf7ff')
        self.ui.setWindowTitle('数据检索')  # 设置该窗口的名称

        self.model = QtGui.QStandardItemModel(self.sheet.nrows, self.sheet.ncols)
        # self.model.setHorizontalHeaderLabels(self.sheet.row_values(rowx=0))
        for row in range(self.sheet.nrows):
            for col in range(self.sheet.ncols):
                item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=row, colx=col)))
                self.model.setItem(row, col, item)
        print("111111111111111")
        self.model_result = QtGui.QStandardItemModel(0, self.sheet.ncols+1)
        # self.model_result = QtGui.QStandardItemModel(0, self.sheet.ncols)
        self.model_result.setHorizontalHeaderLabels(self.sheet.row_values(rowx=0))

        print("2222222222222222")
        # self.ui.tableView_result.setModel(self.model_result)
        # 实例化整型校验器，并设置范围0~65536
        portValidator = QIntValidator(0, 65536)
        # 设置 正则表达式，显示输入0.0.0.0~255.255.255.255
        regExp = QRegExp('^((2[0-4]\d|25[0-5]|\d?\d|1\d{2})\.){3}(2[0-4]\d|25[0-5]|[01]?\d\d?)$')
        # 实例化自定义校验器
        ipValidator = QRegExpValidator(regExp)
        # 实例化浮点校验器，并设置范围-360~360，精度为小数点两位
        doubleValidator = QDoubleValidator(0, 999999, 2)
        # 为文本输入框设置对应的校验器
        self.ui.lineEdit_price.setValidator(doubleValidator)
        self.ui.lineEdit_low.setValidator(doubleValidator)
        self.ui.lineEdit_high.setValidator(doubleValidator)

        print("333333333333333333")

        self.ui.Button_price.setEnabled(False)
        self.ui.Button_name.setEnabled(False)
        self.ui.Button_para.setEnabled(False)
        self.ui.Button_m_name.setEnabled(False)
        self.ui.Button_m_para.setEnabled(False)
        self.ui.Button_low_high.setEnabled(False)
        self.ui.pushButton_all.setEnabled(False)
        self.ui.pushButton.setEnabled(False)
        self.ui.pushButton_2.setEnabled(False)

        print("4444444444444444")
        # self.ui.lineEdit_m_name.editingFinished.connect(self.show_text)
        self.ui.lineEdit_price.textChanged.connect(self.show_text)
        self.ui.lineEdit_name.textChanged.connect(self.show_text)
        self.ui.lineEdit_para.textChanged.connect(self.show_text)
        self.ui.lineEdit_m_name.textChanged.connect(self.show_text)
        self.ui.lineEdit_m_para.textChanged.connect(self.show_text)
        self.ui.lineEdit_low.textChanged.connect(self.show_text)
        self.ui.lineEdit_high.textChanged.connect(self.show_text)
        print("555555555555555555555555")
        self.ui.pushButton_open.clicked.connect(self.open_excel)
        self.ui.Button_price.clicked.connect(self.price_search)
        self.ui.Button_name.clicked.connect(self.name_search)
        self.ui.Button_para.clicked.connect(self.para_search)
        self.ui.Button_m_name.clicked.connect(self.name_search_m)
        self.ui.pushButton.clicked.connect(self.m_context_name)
        self.ui.pushButton_2.clicked.connect(self.m_context_para)
        self.ui.Button_m_para.clicked.connect(self.para_search_m)
        self.ui.Button_low_high.clicked.connect(self.price_search_range)
        # self.ui.pushButton_all.clicked.connect(self.frzzy)
        self.ui.pushButton_all.clicked.connect(self.all_search)
        print("666666666666666666666")

        self.pinyin_name = []
        self.pinyin_para = []
        self.pinyin_price = []
        for i in range(1, self.sheet.nrows):
            self.pinyin_name.append(pypinyin.slug(str(self.sheet.cell_value(rowx=i, colx=3)),separator=''))
            self.pinyin_para.append(pypinyin.slug(str(self.sheet.cell_value(rowx=i, colx=5)),separator=''))
            self.pinyin_price.append(self.sheet.cell_value(rowx=i, colx=7))
        print("777777777777")


    # def pinyin_trans(self):
    #     for i in range(1, self.sheet.nrows):
    #         self.pinyin_name.append(pypinyin.slug(str(self.sheet.cell_value(rowx=i, colx=3)),separator=''))
    #         self.pinyin_para.append(pypinyin.slug(str(self.sheet.cell_value(rowx=i, colx=5)),separator=''))
    #         self.pinyin_price.append(self.sheet.cell_value(rowx=i, colx=7))
        # print(self.pinyin_name)
        # print(self.pinyin_para)
        # print(self.pinyin_price)

    def show_text(self):
        print(f"输入的价格是：{self.ui.lineEdit_price.text()}")
        print(f"输入的名称是：{self.ui.lineEdit_name.text()}")
        print(f"输入的性能参数是：{self.ui.lineEdit_para.text()}")
        print(f"输入的模糊名称是：{self.ui.lineEdit_m_name.text()}")
        # pattern = '.*'.join(self.ui.lineEdit_m_name.text())
        # print(pattern)
        print(f"输入的模糊性能参数是：{self.ui.lineEdit_m_para.text()}")
        print(f"输入的最低价格是：{self.ui.lineEdit_low.text()}")
        print(f"输入的最高价格是：{self.ui.lineEdit_high.text()}")

    def time_strasp(self):
        value = time.time()  # float
        value = int(round(value * 1000))
        return value

    def open_excel(self):
        self.ui.tableView_source.setModel(self.model)
        self.ui.tableView_result.setModel(self.model_result)

        self.ui.Button_price.setEnabled(True)
        self.ui.Button_name.setEnabled(True)
        self.ui.Button_para.setEnabled(True)
        self.ui.Button_m_name.setEnabled(True)
        self.ui.Button_m_para.setEnabled(True)
        self.ui.Button_low_high.setEnabled(True)
        self.ui.pushButton_all.setEnabled(True)
        self.ui.pushButton.setEnabled(True)
        self.ui.pushButton_2.setEnabled(True)
        print("successful")

        # for i in range(1, self.sheet.nrows):
        #     self.pinyin_name.append(pypinyin.slug(str(self.sheet.cell_value(rowx=i, colx=3)),separator=''))
        #     self.pinyin_para.append(pypinyin.slug(str(self.sheet.cell_value(rowx=i, colx=5)),separator=''))
        #     self.pinyin_price.append(self.sheet.cell_value(rowx=i, colx=7))
        # # print(self.pinyin_name)
        # # print(self.pinyin_para)
        # # print(self.pinyin_price)

    def price_search(self):
        # datetime = QDateTime.currentDateTime().toString()
        # print('tineeee:',datetime)
        # print(f'时间: {QDateTime.currentDateTime()[5]}')
        # self.ui.lineEdit_price.setText(4499)
        # print(f"输入的价格是：{float(self.ui.lineEdit_price.text())}")
        # 价格精确
        self.model_result.removeRows(0, self.model_result.rowCount())  # 清空结果表
        value1 = self.time_strasp()  # float
        print(f'第一次计时：{value1}')
        n = 0
        for i in range(1, self.sheet.nrows):
            if self.sheet.cell_value(rowx=i, colx=7) == '':
                continue
            if self.ui.lineEdit_price.text() == '':
                break
            if float(self.sheet.cell_value(rowx=i, colx=7)) == float(self.ui.lineEdit_price.text()):
                n += 1
                print(f"内容是: {self.sheet.row_values(rowx=i)}")
                print(f"内容是: {self.sheet.cell_value(rowx=i, colx=7)}")
                for col in range(11):
                    item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                    self.model_result.setItem(n, col, item)
        print(f"一共检索出: {n} 条记录")
        value2 = self.time_strasp()  # float
        print(f'第二次计时：{value2}')
        print(f'时间差：{value2-value1}')
        self.ui.lcdNumber.display(float(value2-value1))
        self.ui.lineEdit_result.setText(f"一共检索出: {n} 条记录")

    def price_search_range(self):
        # print(f"输入的最低价格是：{int(self.ui.lineEdit_low.text())}")
        # print(f"输入的最高价格是：{self.ui.lineEdit_high.text()}")
        # print(f"高价格类型是：{type(float(self.ui.lineEdit_high.text()))}")
        # print(f"价格类型是：{type(float(self.sheet.cell_value(rowx=2, colx=7)))}")
        # 价格范围
        self.model_result.removeRows(0, self.model_result.rowCount())  # 清空结果表
        value1 = self.time_strasp()  # float
        n = 0
        for i in range(1, self.sheet.nrows):
            if self.sheet.cell_value(rowx=i, colx=7) == '':
                continue
            if self.ui.lineEdit_high.text() == '' and self.ui.lineEdit_low.text() == '':
                break
            if self.ui.lineEdit_high.text() == '':
                if float(self.ui.lineEdit_low.text()) <= float(self.sheet.cell_value(rowx=i, colx=7)):
                    n += 1
                    print(f"内容是: {self.sheet.row_values(rowx=i)}")
                    print(f"内容是: {self.sheet.cell_value(rowx=i, colx=7)}")
                    for col in range(11):
                        item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                        self.model_result.setItem(n, col, item)
            if self.ui.lineEdit_low.text() == '':
                if float(self.sheet.cell_value(rowx=i, colx=7)) <= float(self.ui.lineEdit_high.text()):
                    n += 1
                    print(f"内容是: {self.sheet.row_values(rowx=i)}")
                    print(f"内容是: {self.sheet.cell_value(rowx=i, colx=7)}")
                    for col in range(11):
                        item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                        self.model_result.setItem(n, col, item)
            if self.ui.lineEdit_high.text() != '' and self.ui.lineEdit_low.text() != '':
                if float(self.ui.lineEdit_low.text()) <= float(self.sheet.cell_value(rowx=i, colx=7)) <= float(
                        self.ui.lineEdit_high.text()):
                    n += 1
                    print(f"内容是: {self.sheet.row_values(rowx=i)}")
                    print(f"内容是: {self.sheet.cell_value(rowx=i, colx=7)}")
                    for col in range(11):
                        item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                        self.model_result.setItem(n, col, item)
        print(f"一共检索出: {n} 条记录")
        value2 = self.time_strasp()  # float
        print(f'第二次计时：{value2}')
        print(f'时间差：{value2 - value1}')
        self.ui.lcdNumber.display(float(value2 - value1))
        self.ui.lineEdit_result.setText(f"一共检索出: {n} 条记录")

    def name_search(self):
        # print(f"输入的名称是：{self.ui.lineEdit_name.text()}")
        # 名称精确
        self.model_result.removeRows(0, self.model_result.rowCount())
        value1 = self.time_strasp()
        n = 0
        for i in range(1, self.sheet.nrows):
            if self.sheet.cell_value(rowx=i, colx=3) == '':
                continue
            if self.ui.lineEdit_name.text() == '':
                break
            if self.ui.lineEdit_name.text() in self.sheet.cell_value(rowx=i, colx=3):
                n += 1
                print(f"内容是: {self.sheet.row_values(rowx=i)}")
                print(f"内容是: {self.sheet.cell_value(rowx=i, colx=3)}")
                for col in range(11):
                    item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                    self.model_result.setItem(n, col, item)
        print(f"一共检索出: {n} 条记录")
        value2 = self.time_strasp()  # float
        print(f'第二次计时：{value2}')
        print(f'时间差：{value2 - value1}')
        self.ui.lcdNumber.display(float(value2 - value1))
        self.ui.lineEdit_result.setText(f"一共检索出: {n} 条记录")

    def para_search(self):
        # print(f"输入的性能参数是：{self.ui.lineEdit_para.text()}")
        # 性能精确
        self.model_result.removeRows(0, self.model_result.rowCount())
        value1 = self.time_strasp()  # float
        n = 0
        for i in range(1, self.sheet.nrows):
            if self.sheet.cell_value(rowx=i, colx=5) == '':
                continue
            if self.ui.lineEdit_para.text() == '':
                break
            if self.ui.lineEdit_para.text() in self.sheet.cell_value(rowx=i, colx=5):
                n += 1
                print(f"内容是: {self.sheet.row_values(rowx=i)}")
                print(f"内容是: {self.sheet.cell_value(rowx=i, colx=5)}")
                for col in range(11):
                    item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                    self.model_result.setItem(n, col, item)
        print(f"一共检索出: {n} 条记录")
        value2 = self.time_strasp()  # float
        print(f'第二次计时：{value2}')
        print(f'时间差：{value2 - value1}')
        self.ui.lcdNumber.display(float(value2 - value1))
        self.ui.lineEdit_result.setText(f"一共检索出: {n} 条记录")

    def name_search_m(self):
        # keyword_processor = KeywordProcessor()  # keyword_processor.add_keyword(, )
        # keyword_processor.add_keyword('Big Apple', 'New York')
        # keyword_processor.add_keyword('Bay Area')
        # keywords_found = keyword_processor.extract_keywords('I love Big Apple and Bay Area.')

        # print(f"输入的名称参数是：{type(self.ui.lineEdit_m_name.text())}")
        # 名称模糊
        self.model_result.removeRows(0, self.model_result.rowCount())
        value1 = self.time_strasp()  # float

        str_list = re.split(' ', self.ui.lineEdit_m_name.text())
        h = []
        permutations(str_list, 0, len(str_list), h)
        print(h)
        pattern_list = []
        # for dex in h:
        for dex in range(len(h)):
            # pattern = '.*'.join(dex)
            pattern = '.*'.join(h[dex])
            pattern_list.append(pattern)
            print("pattern", pattern)

        n = 0
        for i in range(1, self.sheet.nrows):
            if self.sheet.cell_value(rowx=i, colx=3) == '':
                continue
            if self.ui.lineEdit_m_name.text() == '':
                break
            for pattern in pattern_list:
                # pattern = pattern_list[j]
                matchObj = re.search(pattern, self.sheet.cell_value(rowx=i, colx=3), re.M | re.I)
                if matchObj:
                    # if re.sub(' ','',self.ui.lineEdit_m_name.text()) in re.sub(' ','',self.sheet.cell_value(rowx=i, colx=3)):
                    n += 1
                    # print(f"内容是: {self.sheet.row_values(rowx=i)}")
                    print(f"内容是: {re.sub(' ', '', self.sheet.cell_value(rowx=i, colx=3))}")
                    print(f"查到的内容是: {matchObj.group()}")
                    # for col in range(11):
                    for col in range(12):
                        if col == 11:
                            item = QtGui.QStandardItem(str(pattern.split("*")))
                            self.model_result.setItem(n, col, item)
                        else:
                            item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                            self.model_result.setItem(n, col, item)
                    break
            # 提取name输入，compile 函数用于编译正则表达式
            # str_list = re.split(' ',self.ui.lineEdit_m_name.text())
            # print(str_list)
            # pattern = '.*'.join(str_list)
            # print(pattern)
            # regex = re.compile(pattern)
            # matchObj = re.search(pattern,self.sheet.cell_value(rowx=i, colx=3),re.M|re.I)
            # # print(matchObj)
            # if matchObj:
            # # if re.sub(' ','',self.ui.lineEdit_m_name.text()) in re.sub(' ','',self.sheet.cell_value(rowx=i, colx=3)):
            #     n += 1
            #     # print(f"内容是: {self.sheet.row_values(rowx=i)}")
            #     print(f"内容是: {re.sub(' ','',self.sheet.cell_value(rowx=i, colx=3))}")
            #     print(f"查到的内容是: { matchObj.group()}")
            #     # for col in range(11):
            #     for col in range(12):
            #         if col == 11:
            #             item = QtGui.QStandardItem(str(str_list))
            #             self.model_result.setItem(n, col, item)
            #         else:
            #             item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
            #             self.model_result.setItem(n, col, item)
        print(f"一共检索出: {n} 条记录")
        value2 = self.time_strasp()  # float
        print(f'第二次计时：{value2}')
        print(f'时间差：{value2 - value1}')
        self.ui.lcdNumber.display(float(value2 - value1))
        self.ui.lineEdit_result.setText(f"一共检索出: {n} 条记录")

    def para_search_m(self):
        # print(f"输入的性能参数是：{self.ui.lineEdit_m_para.text()}")
        # print(f"输入的性能参数是：{type(self.ui.lineEdit_m_para.text())}")
        # 性能模糊
        self.model_result.removeRows(0, self.model_result.rowCount())
        value1 = self.time_strasp()  # float

        str_list = re.split(' ', self.ui.lineEdit_m_para.text())
        h = []
        permutations(str_list, 0, len(str_list), h)
        print(h)
        pattern_list = []
        for dex in h:
        # for dex in range(len(h)):
            pattern = '.*'.join(dex)
            # pattern = '.*'.join(h[dex])
            pattern_list.append(pattern)
            print("pattern", pattern)

        n = 0
        for i in range(1, self.sheet.nrows):
            if self.sheet.cell_value(rowx=i, colx=5) == '':
                continue
            if self.ui.lineEdit_m_para.text() == '':
                break
            for pattern in pattern_list:
            # str_list = re.split(' ',self.ui.lineEdit_m_para.text())
            # print(str_list)
            # pattern = '.*'.join(str_list)
            # print(pattern)
            # regex = re.compile(pattern)
                matchObj = re.search(pattern,self.sheet.cell_value(rowx=i, colx=5),re.M|re.I)
                # print(matchObj)
                if matchObj:
                # if fuzz.token_set_ratio(self.sheet.cell_value(rowx=i, colx=5), str(self.ui.lineEdit_m_para.text())) == 100:
                    n += 1
                    print(f"内容是: {self.sheet.row_values(rowx=i)}")
                    print(f"内容是: {self.sheet.cell_value(rowx=i, colx=5)}")
                    for col in range(12):
                        if col == 11:
                            item = QtGui.QStandardItem(str(pattern.split("*")))
                            self.model_result.setItem(n, col, item)
                        else:
                            item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                            self.model_result.setItem(n, col, item)
        print(f"一共检索出: {n} 条记录")
        value2 = self.time_strasp()  # float
        print(f'第二次计时：{value2}')
        print(f'时间差：{value2 - value1}')
        self.ui.lcdNumber.display(float(value2 - value1))
        self.ui.lineEdit_result.setText(f"一共检索出: {n} 条记录")

    def all_search(self):
        # print(f"输入的名称参数是：{type(self.ui.lineEdit_m_name.text())}")
        # 名称模糊
        self.model_result.removeRows(0, self.model_result.rowCount())
        value1 = self.time_strasp()  # float

        n = 0
        if self.ui.lineEdit_price.text() == '' and self.ui.lineEdit_name.text() != '' and self.ui.lineEdit_para.text() != '':
            for i in range(1, self.sheet.nrows):
                if self.sheet.cell_value(rowx=i, colx=5) == '':
                    continue
                if self.sheet.cell_value(rowx=i, colx=3) == '':
                    continue
                if self.ui.lineEdit_name.text() in self.sheet.cell_value(rowx=i,colx=3) and self.ui.lineEdit_para.text() in self.sheet.cell_value(rowx=i, colx=5):
                    n += 1
                    for col in range(11):
                        item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                        self.model_result.setItem(n, col, item)

        if self.ui.lineEdit_price.text() != '' and self.ui.lineEdit_name.text() == '' and self.ui.lineEdit_para.text() != '':
            for i in range(1, self.sheet.nrows):
                if self.sheet.cell_value(rowx=i, colx=7) == '':
                    continue
                if self.sheet.cell_value(rowx=i, colx=5) == '':
                    continue
                if float(self.sheet.cell_value(rowx=i, colx=7)) == float(self.ui.lineEdit_price.text()) and self.ui.lineEdit_para.text() in self.sheet.cell_value(rowx=i, colx=5):
                    n += 1
                    for col in range(11):
                        item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                        self.model_result.setItem(n, col, item)

        if self.ui.lineEdit_price.text() != '' and self.ui.lineEdit_name.text() != '' and self.ui.lineEdit_para.text() == '':
            for i in range(1, self.sheet.nrows):
                if self.sheet.cell_value(rowx=i, colx=7) == '':
                    continue
                if self.sheet.cell_value(rowx=i, colx=3) == '':
                    continue
                if float(self.sheet.cell_value(rowx=i, colx=7)) == float(self.ui.lineEdit_price.text()) and self.ui.lineEdit_name.text() in self.sheet.cell_value(rowx=i, colx=3):
                    n += 1
                    for col in range(11):
                        item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                        self.model_result.setItem(n, col, item)

        if self.ui.lineEdit_price.text() != '' and self.ui.lineEdit_name.text() != '' and self.ui.lineEdit_para.text() != '':
            for i in range(1, self.sheet.nrows):
                if self.sheet.cell_value(rowx=i, colx=7) == '':
                    continue
                if self.sheet.cell_value(rowx=i, colx=5) == '':
                    continue
                if self.sheet.cell_value(rowx=i, colx=3) == '':
                    continue
                if float(self.sheet.cell_value(rowx=i, colx=7)) == float(self.ui.lineEdit_price.text()) and self.ui.lineEdit_name.text() in self.sheet.cell_value(rowx=i, colx=3) and self.ui.lineEdit_para.text() in self.sheet.cell_value(rowx=i, colx=5):
                    n += 1
                    for col in range(11):
                        item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                        self.model_result.setItem(n, col, item)
        if self.ui.lineEdit_price.text() == '' and self.ui.lineEdit_name.text() == '' and self.ui.lineEdit_para.text() != '':
            for i in range(1, self.sheet.nrows):
                if self.sheet.cell_value(rowx=i, colx=5) == '':
                    continue
                if self.ui.lineEdit_para.text() in self.sheet.cell_value(rowx=i, colx=5):
                    n += 1
                    for col in range(11):
                        item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                        self.model_result.setItem(n, col, item)

        if self.ui.lineEdit_price.text() == '' and self.ui.lineEdit_name.text() != '' and self.ui.lineEdit_para.text() == '':
            for i in range(1, self.sheet.nrows):
                if self.sheet.cell_value(rowx=i, colx=3) == '':
                    continue
                if self.ui.lineEdit_name.text() in self.sheet.cell_value(
                        rowx=i, colx=3):
                    n += 1
                    for col in range(11):
                        item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                        self.model_result.setItem(n, col, item)

        if self.ui.lineEdit_price.text() != '' and self.ui.lineEdit_name.text() == '' and self.ui.lineEdit_para.text() == '':
            for i in range(1, self.sheet.nrows):
                if self.sheet.cell_value(rowx=i, colx=7) == '':
                    continue
                if float(self.sheet.cell_value(rowx=i, colx=7)) == float(
                        self.ui.lineEdit_price.text()):
                    n += 1
                    for col in range(11):
                        item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                        self.model_result.setItem(n, col, item)

        print(f"一共检索出: {n} 条记录")
        value2 = self.time_strasp()  # float
        print(f'第二次计时：{value2}')
        print(f'时间差：{value2 - value1}')
        self.ui.lcdNumber.display(float(value2 - value1))
        self.ui.lineEdit_result.setText(f"一共检索出: {n} 条记录")

    def m_context_name(self):
        self.model_result.removeRows(0, self.model_result.rowCount())
        value1 = self.time_strasp()  # float


        # print(pypinyin.slug(str(self.ui.lineEdit_m_name.text()),style=Style.FIRST_LETTER))
        str_list = re.split(' ', pypinyin.slug(str(self.ui.lineEdit_m_name.text()), separator=''))
        # str_list = re.split(' ', pypinyin.slug(str(self.ui.lineEdit_m_name.text()), style=Style.FIRST_LETTER))

        h = []
        resul = []
        for i in range(1, len(str_list) + 1):
            per(str_list, 0, i, resul, h)
        # print(h)

        h2 = []
        for data in h:
            permutations(data, 0, len(data), h2)
        h2.reverse()
        # print(h2)

        pattern_list = []
        for dex in h2:
        # for dex in h:
            # for dex in range(len(h)):
            pattern = '.*'.join(dex)
            # pattern = '.*'.join(h[dex])
            pattern_list.append(pattern)
            # print("pattern", pattern)

        n = 0
        # for i in range(1, self.sheet.nrows):
        # for name_list in self.pinyin_name:
        for i,name_list in enumerate(self.pinyin_name):
            # print(pypinyin.slug(str(self.sheet.cell_value(rowx=i, colx=3)), style=Style.FIRST_LETTER))
            # if self.sheet.cell_value(rowx=i, colx=3) == '':
            if name_list == '':
                continue
            if self.ui.lineEdit_m_name.text() == '':
                break
            for pattern in pattern_list:
                # pattern = pattern_list[j]
                # matchObj = re.search(pattern, self.sheet.cell_value(rowx=i, colx=3), re.M | re.I)
                # matchObj = re.search(pattern, pypinyin.slug(str(self.sheet.cell_value(rowx=i, colx=3)), separator=''), re.M | re.I)
                matchObj = re.search(pattern, name_list, re.M | re.I)
                # matchObj = re.search(pattern, pypinyin.slug(str(self.sheet.cell_value(rowx=i, colx=3)), style=Style.FIRST_LETTER), re.M | re.I)
                if matchObj:
                    # if re.sub(' ','',self.ui.lineEdit_m_name.text()) in re.sub(' ','',self.sheet.cell_value(rowx=i, colx=3)):
                    n += 1
                    # print(f"内容是: {self.sheet.row_values(rowx=i)}")
                    # print(f"内容是: {re.sub(' ', '', self.sheet.cell_value(rowx=i, colx=3))}")
                    # print(f"查到的内容是: {matchObj.group()}")
                    # for col in range(11):
                    for col in range(12):
                        if col == 11:
                            item = QtGui.QStandardItem(str(pattern.split("*")))
                            self.model_result.setItem(n, col, item)
                        else:
                            item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                            self.model_result.setItem(n, col, item)
                    break
        print(f"一共检索出: {n} 条记录")
        value2 = self.time_strasp()  # float
        print(f'第二次计时：{value2}')
        print(f'时间差：{value2 - value1}')
        self.ui.lcdNumber.display(float(value2 - value1))
        self.ui.lineEdit_result.setText(f"一共检索出: {n} 条记录")

    def m_context_para(self):
        self.model_result.removeRows(0, self.model_result.rowCount())
        value1 = self.time_strasp()  # float

        # str_list = re.split(' ', self.ui.lineEdit_m_para.text())
        str_list = re.split(' ', pypinyin.slug(str(self.ui.lineEdit_m_para.text()), separator=''))

        h = []
        # for i in range(len(str_list)):
        #     if combine(str_list, i) != [()]:
        #         h.extend(combine(str_list, i))
        # start = 0
        resul = []
        for i in range(1, len(str_list) + 1):
            per(str_list, 0, i, resul, h)
        # print(h)

        h2 = []
        for data in h:
            permutations(data, 0, len(data), h2)
        h2.reverse()
        # print(h2)

        pattern_list = []
        for dex in h2:
            # for dex in range(len(h)):
            pattern = '.*'.join(dex)
            # pattern = '.*'.join(h[dex])
            pattern_list.append(pattern)
            # print("pattern", pattern)

        n = 0
        for i,para_list in enumerate(self.pinyin_para):
        # for i in range(1, self.sheet.nrows):
        #     if self.sheet.cell_value(rowx=i, colx=5) == '':
            if para_list == '':
                continue
            if self.ui.lineEdit_m_para.text() == '':
                break
            for pattern in pattern_list:
                # pattern = pattern_list[j]
                # matchObj = re.search(pattern, self.sheet.cell_value(rowx=i, colx=5), re.M | re.I)
                matchObj = re.search(pattern, para_list, re.M | re.I)
                if matchObj:
                    # if re.sub(' ','',self.ui.lineEdit_m_name.text()) in re.sub(' ','',self.sheet.cell_value(rowx=i, colx=3)):
                    n += 1
                    # print(f"内容是: {self.sheet.row_values(rowx=i)}")
                    print(f"内容是: {re.sub(' ', '', self.sheet.cell_value(rowx=i, colx=5))}")
                    print(f"查到的内容是: {matchObj.group()}")
                    # for col in range(11):
                    for col in range(12):
                        if col == 11:
                            item = QtGui.QStandardItem(str(pattern.split("*")))
                            self.model_result.setItem(n, col, item)
                        else:
                            item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                            self.model_result.setItem(n, col, item)
                    break
        print(f"一共检索出: {n} 条记录")
        value2 = self.time_strasp()  # float
        print(f'第二次计时：{value2}')
        print(f'时间差：{value2 - value1}')
        self.ui.lcdNumber.display(float(value2 - value1))
        self.ui.lineEdit_result.setText(f"一共检索出: {n} 条记录")

    def frzzy(self):
        # print(f"输入的名称是：{self.ui.lineEdit_name.text()}")
        # 名称精确
        print(222222)
        self.model_result.removeRows(0, self.model_result.rowCount())
        value1 = self.time_strasp()
        n = 0
        for i in range(1, self.sheet.nrows):
            if self.sheet.cell_value(rowx=i, colx=3) == '':
                continue
            if self.ui.lineEdit_m_name.text() == '':
                break
            # print(self.sheet.cell_value(rowx=i, colx=3))
            # print(4444444)
            s = fuzz.partial_ratio(pypinyin.slug(self.sheet.cell_value(rowx=i, colx=3), separator=''),
                                   self.ui.lineEdit_m_name.text(), separator='')
            # words = list(jieba.cut(self.sheet.cell_value(rowx=i, colx=3)))
            # print(words)
            # for i in words:
                # print(pypinyin.slug(i, separator=''))
                # print(pypinyin.slug(self.ui.lineEdit_m_name.text(), separator=''))

                # s = fuzz.partial_ratio(pypinyin.slug(i, separator=''),
                #                self.ui.lineEdit_m_name.text(), separator='')

            print(s)
            if s > 49:
                # print(s)
                n += 1
                for col in range(11):
                    item = QtGui.QStandardItem(str(self.sheet.cell_value(rowx=i, colx=col)))
                    self.model_result.setItem(n, col, item)

        print(f"一共检索出: {n} 条记录")
        value2 = self.time_strasp()  # float
        print(f'第二次计时：{value2}')
        print(f'时间差：{value2 - value1}')
        self.ui.lcdNumber.display(float(value2 - value1))
        self.ui.lineEdit_result.setText(f"一共检索出: {n} 条记录")

if __name__ == "__main__":
    # synlst = synonyms.display('笔记本')
    # print(synonyms.compare("电脑", "笔记本"))
    # print(synlst)
    # wordVec = gensim.models.KeyedVectors.load_word2vec_format("sohu.news.word2vec.bin", binary=True)
    # sim = wordVec.most_similar('computer',topn=10)# 取得最相似的前十个单词
    app = QApplication([])

    # 启动界面  创建QSplashScreen对象实例
    splash = QtWidgets.QSplashScreen(QtGui.QPixmap("./logo.png"))
    # 显示画面
    splash.show()
    stats = Stats()

    print("888888888888888888")
    stats.ui.show()
    splash.finish(stats.ui)
    app.exec_()

