# -*- coding: utf-8 -*-
"""
@Time ： 2020/11/18 21:22
@Auth ： Mr. William 1052949192
@Company ：特斯汀学院 @testingedu.com.cn
@Function ：兼容不同的excel版本
"""
import os
from common.Excels import NewExcel
from common.Excels import OldExcel


def get_reader(srcfile='') -> NewExcel.Reader:
    """
    获取读取excel的对象
    :param srcfile: excel文件路径
    :return: 读取excel的对象
    """
    reader = None

    # 如果打开的文件不存在，就报错
    if not os.path.isfile(srcfile):
        print("%s not exist!" % (srcfile))
        return reader

    if srcfile.endswith('.xls'):
        reader = OldExcel.Reader()
        reader.open_excel(srcfile)
        return reader

    if srcfile.endswith('.xlsx'):
        reader = NewExcel.Reader()
        reader.open_excel(srcfile)
        return reader


def get_writer(srcfile, dstfile) -> NewExcel.Writer:
    """
    获取写入excel的对象
    :param srcfile: excel远文件路径
    :param dstfile: 写入后excel文件保存的路径
    :return: 写入excel的对象
    """
    writer = None

    # 如果打开的文件不存在，就报错
    if not os.path.isfile(srcfile):
        print("%s not exist!" % (srcfile))
        return writer

    if srcfile.endswith('.xls'):
        writer = OldExcel.Writer()
        writer.copy_open(srcfile, dstfile)
        return writer

    if srcfile.endswith('.xlsx'):
        writer = NewExcel.Writer()
        writer.copy_open(srcfile, dstfile)
        return writer


# 调试
if __name__ == '__main__':
    reader = get_reader('../lib/cases/电商项目用例.xlsx')
    sheetname = reader.get_sheets()
    print(sheetname)
    for sheet in sheetname:
        # 设置当前读取的sheet页面
        reader.set_sheet(sheet)
        lines = reader.readline()
        print(lines)

        print()
        break

    writer = get_writer('../lib/cases/电商项目用例.xlsx', '../lib/cases/result-电商项目用例.xlsx')
    sheetname = writer.get_sheets()
    writer.set_sheet(sheetname[0])
    writer.write(1, 1, 'William', 3)
    writer.save_close()
