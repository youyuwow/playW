
import io
import os
import sys

from ddt.excel_ddt import ddt

if __name__ == '__main__':
    # 每一次运行前，删除结果和报告（保证结果的准确性）
    os.system('rd /s/q result')
    os.system('rd /s/q report')

    ddt.run_web_case('./lib/cases/电商项目用例.xlsx')
    os.system('allure generate result -o report --clean')
