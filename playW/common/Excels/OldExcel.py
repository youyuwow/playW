# -*- coding: utf-8 -*-
"""
@Time ： 2018-12-21 21:07
@Auth ： Mr. William 1052949192
@Company ：特斯汀学院 @testingedu.com.cn
@Function ：读写Excel文件
"""
import os, xlrd, xlwt
from xlutils.copy import copy

class Reader:
    """
        powered by Mr Will
           at 2018-12-21
        用来读取Excel文件内容
    """

    def __init__(self):
        # 整个excel工作簿缓存
        self.workbook = None
        # 当前工作sheet
        self.sheet = None
        # 当前sheet的行数
        self.rows = 0
        # 当前读取到的行数
        self.r = 0

    # 打开excel
    def open_excel(self, srcfile):
        # 如果打开的文件不存在，就报错
        if not os.path.isfile(srcfile):
            print("%s not exist!" % (srcfile))
            return

        # 设置读取excel使用utf8编码
        xlrd.Book.encoding = "utf8"
        # 读取excel内容到缓存workbook
        self.workbook = xlrd.open_workbook(filename=srcfile)
        # 选取第一个sheet页面
        self.sheet = self.workbook.sheet_by_index(0)
        # 设置rows为当前sheet的行数
        self.rows = self.sheet.nrows
        # 设置默认读取为第一行
        self.r = 0
        return

    # 获取sheet页面
    def get_sheets(self):
        # 获取所有sheet的名字，并返回为一个列表
        sheets = self.workbook.sheet_names()
        # print(sheets)
        return sheets

    # 切换sheet页面
    def set_sheet(self, name):
        # 通过sheet名字，切换sheet页面
        self.sheet = self.workbook.sheet_by_name(name)
        self.rows = self.sheet.nrows
        self.r = 0
        return

    # 读取当前sheet全部行
    def readline(self):
        lines = []
        # 如果当前还没到最后一行，则往下读取一行
        while self.r < self.rows:
            row1 = None
            # 读取第r行的内容
            row = self.sheet.row_values(self.r)
            # 设置下一次读取r的下一行
            self.r = self.r + 1
            # 辅助遍历行里面的列
            i = 0
            row1 = row
            # 把读取的数据都变为字符串
            for strs in row:
                row1[i] = str(strs)
                i = i + 1
            lines.append(row1)

        return lines


class Writer:
    """
        powered by Mr Will
           at 2018-12-21
        用来复制写入Excel文件
    """

    def __init__(self):
        # 读取需要复制的excel
        self.workbook = None
        # 拷贝的工作空间
        self.wb = None
        # 当前工作的sheet页
        self.sheet = None
        # 记录生成的文件，用来保存
        self.df = None
        # 记录写入的行
        self.row = 0
        # 记录写入的列
        self.clo = 0

    # 复制并打开excel
    def copy_open(self, srcfile, dstfile):
        # 判断要复制的文件是否存在
        if not os.path.isfile(srcfile):
            print(srcfile + " not exist!")
            return

        # 判断要新建的文档是否存在，存在则提示
        if os.path.isfile(dstfile):
            print(dstfile + " file already exist!")

        # 记录要保存的文件
        self.df = dstfile
        # 读取excel到缓存
        # formatting_info带格式的复制
        self.workbook = xlrd.open_workbook(filename=srcfile, formatting_info=True)
        # 拷贝，也在内存里面
        self.wb = copy(self.workbook)
        return

    # 获取sheet页面
    def get_sheets(self):
        # 获取所有sheet的名字，并返回为一个列表
        sheets = self.workbook.sheet_names()
        # print(sheets)
        return sheets

    # 切换sheet页面
    def set_sheet(self, name):
        # 通过sheet名字，切换sheet页面
        self.sheet = self.wb.get_sheet(name)
        return

    # 写入指定单元格，保留原格式
    def write(self, r, c, value, color=None):
        """
        :param r: 行
        :param c: 列
        :param value: 要写入的字符串
        :param color: 0，黑色；1，白色；2，红色；3，绿色；4，蓝色；5，黄色
        :return:
        """

        # 获取要写入的单元格
        def _getCell(sheet, r, c):
            """ HACK: Extract the internal xlwt cell representation. """
            # 获取行
            row = sheet._Worksheet__rows.get(r)
            if not row:
                return None

            # 获取单元格
            cell = row._Row__cells.get(c)
            return cell

        # 获取要写入的单元格
        cell = _getCell(self.sheet, r, c)
        # 格式，把单元格原来的格式保存下来
        if cell:
            idx = cell.xf_idx

        if color is None:
            self.sheet.write(r, c, value)
            if cell:
                # 获取要写入的单元格
                ncell = _getCell(self.sheet, r, c)
                if ncell:
                    # 设置写入后格式和写入前一样
                    ncell.xf_idx = idx

        else:
            style = xlwt.XFStyle()
            font = xlwt.Font()  # 创建字体
            font.name = 'Arial'
            font.bold = True  # 黑体
            # font.underline = True  # 下划线
            # font.italic = True  # 斜体字
            font.colour_index = color  # 颜色为红色
            style.font = font

            # borders.left = xlwt.Borders.THIN
            # NO_LINE： 官方代码中NO_LINE所表示的值为0，没有边框
            # THIN： 官方代码中THIN所表示的值为1，边框为实线
            borders = xlwt.Borders()
            # 0 为黑色，与color参数一致
            borders.left = xlwt.Borders.THIN
            borders.right = xlwt.Borders.THIN
            borders.top = xlwt.Borders.THIN
            borders.bottom = xlwt.Borders.THIN
            style.borders = borders
            # 写入值
            self.sheet.write(r, c, value, style)

        return

    # 保存
    def save_close(self):
        # 保存复制后的文件到硬盘
        self.wb.save(self.df)
        return



if __name__ == '__main__':
    reader = Reader()
    # 打开一个excel
    reader.open_excel('../../lib/cases/电商项目用例.xls')
    # 获取所有sheet
    sheets = reader.get_sheets()
    for sheet in sheets:
        # 设置读取的sheet页面
        reader.set_sheet(sheet)
        # 读取当前sheet的所有行
        lines = reader.readline()
        print(lines)

    writer = Writer()
    writer.copy_open('../../lib/cases/电商项目用例.xls','../../lib/cases/result-电商项目用例.xls')

    # 获取所有sheet
    sheets = writer.get_sheets()
    writer.set_sheet(sheets[0])
    # 在原格式上写入内容
    writer.write(1,1,'Will')
    # 更改颜色写入
    writer.write(1,2,'hello',color=2)
    writer.save_close()

