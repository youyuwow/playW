# -*- coding: utf-8 -*-
"""
@Time ： 2022/3/30 21:36
@Auth ： Mr. ZX 2929184523
@Company ：特斯汀学院 @testingedu.com.cn
@Function ：通过修改文件名，复用test_web.py运行所有的用例
"""
import os

import pytest

from common.Excel import get_reader, get_writer
from common.log import logger


class DDT:

    def __init__(self):
        """初始化一些实例变量"""
        from playwright.sync_api import sync_playwright

        play = sync_playwright().start()
        browser = play.chromium.launch(headless=False)
        page = browser.new_page()

        self.web = page

        # 记录项目模块、分组名字
        self.feature = ''
        self.story = ''
        # 记录重命名的序号
        self.story_idx = 0
        # 记录项目模块的序号
        self.feature_idx = 0

        # 记录一个模块的用例
        self.cases = []
        # 记录运行的类型
        self.type = 'web'

        # 写入Excel结果
        self.writer = None

    def __run_pytest_case(self):
        logger.info(str(self.story_idx))
        # 通过更改文件名，去运行数据驱动
        os.rename('./ddt/test_%s_%d.py' % (self.type, self.story_idx - 1,),
                  './ddt/test_%s_%d.py' % (self.type, self.story_idx,))
        pytest.main(['-s', './ddt/test_%s_%d.py' % (self.type, self.story_idx,), '--alluredir', 'result'])

    def run_web_case(self, filepath='./lib/cases/电商项目用例.xlsx'):
        self.type = 'web'
        reader = get_reader(filepath)
        sheetname = reader.get_sheets()
        logger.info(sheetname)
        for sheet in sheetname:
            # 设置当前读取的sheet页面
            reader.set_sheet(sheet)
            # 设置项目模块名字
            self.feature = sheet
            self.feature_idx += 1
            lines = reader.readline()

            case = []
            # 表头不需要
            for i in range(1, len(lines)):
                line = lines[i]
                # 第一个单元格有内容，说明是一个模块
                if len(line[0]) > 0:
                    # 如果模块用例不为空，说明上一个模块统计完成
                    # 你需要把上一个用例添加到cases里面去，并且执行整个模块用例
                    if case:
                        self.cases.append(case)
                        logger.debug(self.cases)
                        logger.debug('执行用例')
                        self.__run_pytest_case()

                    self.cases = []
                    case = []
                    # 记录模块名字
                    self.story = line[0]
                    logger.debug(line)
                    self.story_idx += 1
                # 第二个单元格有内容，说明是用例
                elif len(line[1]) > 0:
                    # 如果case不空，就说明上一组用例统计完成
                    # 我们把用例放到模块用例cases里面去
                    if case:
                        self.cases.append(case)

                    # 用一个列表准备存放一组用例的所有数据
                    case = []
                    case.append(line)
                else:
                    # 记录一组用例的数据
                    case.append(line)

            # 一个sheet统计完成，把最后一个用例添加到模块里面
            # 并且执行最后一个sheet
            if case:
                self.cases.append(case)
                logger.debug(self.cases)
                logger.debug('执行用例')
                self.__run_pytest_case()

            # 注意置空，否则，每一个sheet，最后一个分组可能多次执行
            self.cases = []
            case = []

        # 所有用例跑完之后，把文件名还原
        os.rename('./ddt/test_web_%d.py' % (self.story_idx,), './ddt/test_web_0.py')


ddt = DDT()
