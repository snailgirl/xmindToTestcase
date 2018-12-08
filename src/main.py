#coding:utf8
# import subprocess

# 安装依赖包
# import os
# file_path = os.path.dirname(os.path.dirname(__file__))
# print(file_path)
# setup_path = os.path.join(file_path, 'lib', 'setup.py')
# subprocess.Popen('python3 ' + setup_path + ' install', shell=True)

from lib.readXmind import *
from  lib.writeExcel import *
import logging
def get_xmind_content(xmind_file,output_file):
    #生成测试用例
    read_xmind = ReadXmindList(xmind_file)
    # excel_title=read_xmind.excel_title #获取模块标题
    write_excel = WriteExcel(output_file)
    testcase_list = []
    read_xmind.get_list_content(read_xmind.content, testcase_list, write_excel )#写入excel
    write_excel.write_analysis_wooksheek()#写入测试分析excel
    write_excel.save_excel() #保存excel
    logging.info("Generate Xmind file successfully: {}".format(output_file))
