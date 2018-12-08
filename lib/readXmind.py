import sys
from xmindparser import xmind_to_dict
from .writeExcel import *
class ReadXmindList(object):
    def __init__(self, filename):
        self.filename = filename #xmind文件路径
        self.content, self.canvas_name, self.excel_title = self.__get_dic_content(self.filename)

    def __get_dic_content(self, filename):
        """
        #从Xmind读取字典形式的数据
        :param filename: xmind文件路径
        :return:
        """
        if not os.path.exists(filename):
            print("[ERROR] 文件不存在")
            sys.exit(-1)
        out = xmind_to_dict(filename)
        dic_content = out[0]
        canvas_name = dic_content.get('title')  # 获取画布名称
        cavas_values = dic_content.get('topic')
        if cavas_values:
            excel_title = cavas_values.get('title')  # 获取模块名称
        content = [dic_content.get('topic')]
        return content, canvas_name, excel_title

    def __format_list(self, testcase_list):
        """
        格式化为excell需要的列表
        :param testcase_list: 需要处理的列表
        :return:
        """
        new_testcase = []
        step_list = []
        expected_list = []
        tag = False
        step_index = 1  # 用例步骤编号
        expected_index = 1  # 预期结果编号
        testcase = [item for item in testcase_list]
        for item in testcase:
            item_index = testcase.index(item)
            if item != '预期结果' and not tag:
                if item_index <= 4:
                    new_testcase.append(item)
                if item_index >= 5:
                    if step_index != 1:
                        step_list.append('\n' + str(step_index) + '.' + item)
                    else:
                        step_list.append(str(step_index) + '.' + item)
                    step_index += 1
            if item == '预期结果':
                tag = True
                continue
            if tag:
                if expected_index != 1:
                    expected_list.append('\n' + str(expected_index) + '.' + item)
                else:
                    expected_list.append(str(expected_index) + '.' + item)
                expected_index += 1
        if step_list:
            new_testcase.insert(5, step_list)
        return new_testcase, expected_list

    def get_list_content(self, content, testcase_list, write_excel):
        """
        #从Xmind文件中读取数据，保存至excel文件中
        :param testcase_list: 储存处理后的列表信息
        :param content:  xmind读取的内容
        :param write_excel: WriteExcel 实例对象
        :return:
        """
        for dic_val in content:
            val = list(dic_val.values())
            if val:
                testcase_list.append(val[0])
                if len(val) == 2:
                    self.get_list_content(val[1], testcase_list, write_excel)
                else:
                    new_testcase = self.__format_list(testcase_list)  # 格式化为excell需要的数据
                    # print(new_testcase)
                    write_excel.write_testcase_excel(new_testcase)  # 写入测试用例
                    write_excel.write_outline_excel(new_testcase)   # 写入测试大纲
                    write_excel.write_testscope_wooksheek(new_testcase)  # 写入测试范围
                    testcase_list.pop()
        if testcase_list:
            testcase_list.pop()