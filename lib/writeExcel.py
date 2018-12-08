
import os,xlwt,datetime
# from conf.settings import filename_path

class WriteExcel():
    style= xlwt.easyxf('pattern: pattern solid, fore_colour 0x31; font: bold on;alignment:HORZ CENTER;'
                         'borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')
    style_nocenter = xlwt.easyxf('pattern: pattern solid, fore_colour White;'
                         'borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')#未居中无背景颜色0x34
    style_center = xlwt.easyxf('pattern: pattern solid, fore_colour White;alignment:HORZ CENTER;'
                                 'borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')  # 无背景颜色居中
    def __init__(self, output_file):
        self.testcase_filename = output_file  # 生成用例的目录
        self.wookbook=self.__init_excel()
        self.testcase_wooksheek=self.__init_testcase_wooksheek() #测试用例
        self.testscope_wooksheek = self.__init_testscope_wooksheek()  # 测试范围
        self.outline_wooksheek=self.__init_outline_wooksheek()  #测试大纲
        self.analysis_wooksheek=self.__init_analysis_wooksheek()  #测试分析
        self.__row = 1  # 测试用例、测试大纲excel行数
        self.__temp_list=[] #测试范围临时存储列表
        self.__testscope_row=1 ## 测试范围excel行数

    def __init_excel(self):
        """
        初始化excel
        :return:
        """
        f = open(self.testcase_filename, 'w')
        f.close()
        wookbook = xlwt.Workbook()  # 创建工作簿
        return wookbook

    def __init_outline_wooksheek(self):
        """
        初始化测试大纲sheet
        :return:
        """
        outline_wooksheek = self.wookbook.add_sheet('测试大纲', cell_overwrite_ok='True')  # 测试大纲
        for i in range(14):
            outline_wooksheek.col(i).width = (13 * 367)
        outline_wooksheek.write(0, 0, '需求编号', self.style)
        outline_wooksheek.write(0, 1, '功能模块', self.style)
        outline_wooksheek.write(0, 2, '功能名称', self.style)
        outline_wooksheek.write(0, 3, '子功能名称', self.style)
        outline_wooksheek.write(0, 4, '功能点',self.style)
        outline_wooksheek.write(0, 5, '用例类型', self.style)
        outline_wooksheek.write(0, 6, '检查点', self.style)
        outline_wooksheek.write(0, 7, '用例设计', self.style)
        outline_wooksheek.write(0, 8, '预期结果', self.style)
        outline_wooksheek.write(0, 9, '类别', self.style)
        outline_wooksheek.write(0, 10, '责任人', self.style)
        outline_wooksheek.write(0, 11, '状态', self.style)
        outline_wooksheek.write(0, 12, '更新日期', self.style)
        outline_wooksheek.write(0, 13, '用例编号', self.style)
        self.save_excel()
        return  outline_wooksheek

    def __init_testcase_wooksheek(self):
        """
        初始化测试用例sheet
        :return:
        """
        testcase_wooksheek = self.wookbook.add_sheet('测试用例', cell_overwrite_ok='True')  # 测试用例
        for i in range(12):
            testcase_wooksheek.col(i).width = (15 * 367)
        testcase_wooksheek.write(0, 0, '用例目录', self.style)
        testcase_wooksheek.write(0, 1, '用例名称', self.style)
        testcase_wooksheek.write(0, 2, '前置条件', self.style)
        testcase_wooksheek.write(0, 3, '用例步骤', self.style)
        testcase_wooksheek.write(0, 4, '预期结果', self.style)
        testcase_wooksheek.write(0, 5, '用例类型', self.style)
        testcase_wooksheek.write(0, 6, '用例状态', self.style)
        testcase_wooksheek.write(0, 7, '用例等级', self.style)
        testcase_wooksheek.write(0, 8, '需求ID', self.style)
        testcase_wooksheek.write(0, 9, '创建人', self.style)
        testcase_wooksheek.write(0, 10, '测试结果', self.style)
        testcase_wooksheek.write(0, 11, '是否开发自测', self.style)
        self.save_excel()
        return testcase_wooksheek

    def __init_testscope_wooksheek(self):
        """
        初始化测试范围
        :return:
        """
        testscope_wooksheek = self.wookbook.add_sheet('测试范围', cell_overwrite_ok='True')  # 测试范围
        for i in range(8):
            testscope_wooksheek.col(i).width = (13 * 367)
        testscope_wooksheek.write(0, 0, '序号', self.style)
        testscope_wooksheek.write(0, 1, '功能模块', self.style)
        testscope_wooksheek.write(0, 2, '功能名称', self.style)
        testscope_wooksheek.write(0, 3, '子功能名称', self.style)
        testscope_wooksheek.write(0, 4, '角色', self.style)
        testscope_wooksheek.write(0, 5, '责任人', self.style)
        testscope_wooksheek.write(0, 6, '更新日期', self.style)
        testscope_wooksheek.write(0, 7, '备注', self.style)
        self.save_excel()
        return testscope_wooksheek

    def __init_analysis_wooksheek(self):
        analysis_wooksheek = self.wookbook.add_sheet('测试分析', cell_overwrite_ok='True')  # 测试范围
        # testcase_wooksheek.col(0).width = 256
        for i in range(10):
            analysis_wooksheek.col(i).width = (10 * 367)
            analysis_wooksheek.write_merge(1, 1, 1,9,'测试覆盖范围及执行结果', self.style)
        analysis_wooksheek.write(2, 1, '项目', self.style_nocenter)
        analysis_wooksheek.write_merge(2, 2, 2,3,'', self.style_nocenter)
        analysis_wooksheek.write(2, 4,'需求编号', self.style_nocenter)
        analysis_wooksheek.write_merge(2, 2,5,7,'', self.style_nocenter)
        analysis_wooksheek.write(2, 8, '产品', self.style_nocenter)
        analysis_wooksheek.write(2, 9, '',self.style_nocenter)
        analysis_wooksheek.write(3, 1, '投产日期', self.style_nocenter)
        analysis_wooksheek.write_merge(3, 3, 2,3,'', self.style_nocenter)
        analysis_wooksheek.write(3, 4,'迭代编号', self.style_nocenter)
        analysis_wooksheek.write_merge(3, 3,5,7,'', self.style_nocenter)
        analysis_wooksheek.write(3, 8, '测试周期', self.style_nocenter)
        analysis_wooksheek.write(3, 9, '', self.style_nocenter)
        analysis_wooksheek.write(4, 1, '测试人员', self.style_nocenter)
        analysis_wooksheek.write_merge(4, 4,2,9, '', self.style_nocenter)
        analysis_wooksheek.write(5, 1, '环境',self.style_nocenter)
        analysis_wooksheek.write_merge(5, 5,2,9, 'SIT/UAT/生产',self.style_nocenter)
        analysis_wooksheek.write(6, 1, '模块', self.style)
        analysis_wooksheek.write(6, 2, 'Total', self.style)
        analysis_wooksheek.write(6, 3, 'Pass', self.style)
        analysis_wooksheek.write(6, 4, 'Fail', self.style)
        analysis_wooksheek.write(6, 5, 'Block', self.style)
        analysis_wooksheek.write(6, 6, 'NA', self.style)
        analysis_wooksheek.write(6, 7, 'Not Run', self.style)
        analysis_wooksheek.write(6, 8, 'Run Rate', self.style)
        analysis_wooksheek.write(6, 9, 'Pass Ratee', self.style)
        self.save_excel()
        return analysis_wooksheek

    def write_outline_excel(self,new_testcase):
        """
        写入测试大纲excel
        :param new_testcase:写入的列表信息
        :return:
        """
        style = xlwt.easyxf('borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')
        for i in range(14):
            self.outline_wooksheek.write(self.__row,i,"",style)
        col=1
        for item in new_testcase[0]:
            if col==5:
                self.outline_wooksheek.write(self.__row,col,'功能',style)
                col+=1
                self.outline_wooksheek.write(self.__row,col,item,style)
            else:
                self.outline_wooksheek.write(self.__row,col,item,style)
            col += 1
        if new_testcase[1]:
            self.outline_wooksheek.write(self.__row, 8, new_testcase[1],style)
        self.outline_wooksheek.write(self.__row, 5, '功能',style)
        self.outline_wooksheek.write(self.__row, 11, 'Not Run',style)
        date_time=datetime.date.today()
        self.outline_wooksheek.write(self.__row, 12,str(date_time),style)
        self.__row+=1

    def write_testcase_excel(self,new_testcase):
        """
        写入测试用例excel
        :param new_testcase:写入的列表信息
        :return:
        """
        style = xlwt.easyxf('borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')
        for i in range(12):
            self.testcase_wooksheek.write(self.__row, i, "", style)
        if len(new_testcase[0])>=1:
            self.testcase_wooksheek.write(self.__row, 0, new_testcase[0][0], style)
        if len(new_testcase[0])>=2:
            self.testcase_wooksheek.write(self.__row, 0, new_testcase[0][0] + '-' +new_testcase[0][1], style)
            self.testcase_wooksheek.write(self.__row, 2, '1.进入【'+new_testcase[0][1]+'】界面；', style)
        if len(new_testcase[0])>=3:
            self.testcase_wooksheek.write(self.__row, 1, new_testcase[0][2], style)
            self.testcase_wooksheek.write(self.__row, 2, '1.进入【'+new_testcase[0][1]+'-'+new_testcase[0][2]+'】界面；', style)
        if len(new_testcase[0]) >= 5:
            self.testcase_wooksheek.write(self.__row, 1, new_testcase[0][3] + '-' + new_testcase[0][4], style)
        if len(new_testcase[0]) >= 6:
            self.testcase_wooksheek.write(self.__row, 3, new_testcase[0][5], style)
        if new_testcase[1]:
            self.testcase_wooksheek.write(self.__row, 4, new_testcase[1], style)
        self.testcase_wooksheek.write(self.__row, 5, '功能测试', style)
        self.testcase_wooksheek.write(self.__row, 6, '正常', style)
        self.testcase_wooksheek.write(self.__row, 7, '中', style)
        self.testcase_wooksheek.write(self.__row, 10, 'Not Run', style)
        self.testcase_wooksheek.write(self.__row, 11, '否', style)
        # self.__row+=1

    def write_testscope_wooksheek(self,new_testcase):
        """
        写入测试范围sheet
        :param new_testcase:
        :return:
        """
        style = xlwt.easyxf('borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')
        item_list=new_testcase[0][0:3]
        col=1
        if item_list not in self.__temp_list:
            self.__temp_list.append(item_list)
            for i in range(8):
                self.testscope_wooksheek.write(self.__testscope_row, i, "", style)
            for item in item_list:
                self.testscope_wooksheek.write(self.__testscope_row,col,item,style)
                col+=1
            self.__testscope_row+=1

    def write_analysis_wooksheek(self):
        row=7
        temp_list=[]
        for item in self.__temp_list:
            if len(item)>=2:
                if item[1] not in temp_list:
                    temp_list.append(item[1])
                    self.analysis_wooksheek.write(row,1,item[1],self.style_center)
                    self.analysis_wooksheek.write(row,2,0,self.style_center)
                    self.analysis_wooksheek.write(row,3,0,self.style_center)
                    self.analysis_wooksheek.write(row,4,0,self.style_center)
                    self.analysis_wooksheek.write(row,5,0,self.style_center)
                    self.analysis_wooksheek.write(row,6,0,self.style_center)
                    self.analysis_wooksheek.write(row,7,0,self.style_center)
                    self.analysis_wooksheek.write(row,8,'0.00%',self.style_center)
                    self.analysis_wooksheek.write(row,9,'0.00%',self.style_center)
                    row += 1
            else:
                self.analysis_wooksheek.write(row, 1, '', self.style_center)
                self.analysis_wooksheek.write(row, 2,0,self.style_center)
                self.analysis_wooksheek.write(row, 3,0,self.style_center)
                self.analysis_wooksheek.write(row, 4,0,self.style_center)
                self.analysis_wooksheek.write(row, 5,0,self.style_center)
                self.analysis_wooksheek.write(row, 6,0,self.style_center)
                self.analysis_wooksheek.write(row, 7,0,self.style_center)
                self.analysis_wooksheek.write(row, 8,'0.00%',self.style_center)
                self.analysis_wooksheek.write(row, 9,'0.00%',self.style_center)
                row += 1
        self.analysis_wooksheek.write(row, 1, '总计', self.style_center)
        self.analysis_wooksheek.write(row, 2, xlwt.Formula("SUM(C8:C"+str(row+1)+")"),self.style_center)
        self.analysis_wooksheek.write(row, 3, xlwt.Formula("SUM(D8:D"+str(row+1)+")"),self.style_center)
        self.analysis_wooksheek.write(row, 4, xlwt.Formula("SUM(E8:E"+str(row+1)+")"),self.style_center)
        self.analysis_wooksheek.write(row, 5, xlwt.Formula("SUM(F8:F"+str(row+1)+")"),self.style_center)
        self.analysis_wooksheek.write(row, 6, xlwt.Formula("SUM(G8:G"+str(row+1)+")"),self.style_center)
        self.analysis_wooksheek.write(row, 7, xlwt.Formula("SUM(H8:H"+str(row+1)+")"),self.style_center)
        self.analysis_wooksheek.write(row, 8,'0.00%',self.style_center)
        self.analysis_wooksheek.write(row, 9,'0.00%',self.style_center)
        row += 2
        self.analysis_wooksheek.write(row, 1, '说明:')
        row+=1
        self.analysis_wooksheek.write(row, 1, 'Pass-验证通过 Fail-验证未通过 Block-阻塞 NA-本期不涉及 Not Run-尚未执行')
        row+=1
        self.analysis_wooksheek.write(row, 1, 'Run Rate=(Pass+Fail+Block)/(Total-NA)')
        self.analysis_wooksheek.write(row+1, 1, 'Pass Rate=Pass/(Total-NA)')

    def save_excel(self):
        """
        保存excel
        :return:
        """
        self.wookbook.save(self.testcase_filename)