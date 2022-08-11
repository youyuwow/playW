import os
import time
import traceback

import allure
import pytest

from ddt.excel_ddt import ddt


@allure.feature('#' + str(ddt.feature_idx) + ' ' + ddt.feature)
class Test_Web:

    @allure.step
    def run_step(self, func, params):
        # 有参数就传
        if params:
            return func(*params)
        else:
            # 没有就不传
            return func()

    @allure.story('#' + str(ddt.story_idx) + ' ' + ddt.story)
    @pytest.mark.parametrize('cases', ddt.cases)
    def test_case(self, cases):
        """测试用例"""
        time.sleep(1)
        allure.dynamic.title(cases[0][1])
        cases = cases[1:]
        try:
            for case in cases:
                func = getattr(ddt.web, case[3])
                # 参数处理
                params = case[4:]
                # 截取非空参数
                params = params[:params.index('')]
                with allure.step(case[2]):
                    self.run_step(func, params)
            time.sleep(0.3)
            # 成功后截图
            allure.attach(ddt.web.screenshot(), '成功截图', allure.attachment_type.PNG)
        except Exception as e:
            # 失败截图
            time.sleep(0.3)
            allure.attach(ddt.web.screenshot(), '失败截图', allure.attachment_type.PNG)
            pytest.fail(str(traceback.format_exc()))

