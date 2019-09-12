# coding=utf-8
# Start Date: 2015/11/24
# Intro.: This project is to develop the new version of "PyACT for checklist" with the excel UI,
# based on the calculation core A0.97 Beta_(150810)

# 2015/11/24~2015/11/27:
#   Build the Excel UI, and modify the original "PyACT fo r checklist" code to fit the check topology requirement
# 2015/11/30~2015/12/1:
#   Modify the extraction function, add new result for every component and crossover
# 2015/12/2~2015/12/3:
#   Test diff and se extraction result and debug
# 2015/12/4:
#   Use VBA to specify the absolute path of template file
#   Release A1.0 ver.
# 2015/12/7:
#   Add pin-mapping database for BASSO, Z6, and Z8
# 2015/12/8:
#   Add pin-mapping database for BASSO (TI Chip)
#   Add functions:Clear Result, Show Result
#   Improve "Detect Start & End Component"
# 2015/12/9:
#   Add pin-mapping for Tesla
#   Modify "Detect Start & End Component", export all relation items
# 2015/12/14:
#   Upgrade main function to extract topology: use sch name and sch pin to determine the start seg ind of one net
#   Add "Cross Over" section in topology format
# 2015/12/16: A1.2 Release
# 2015/12/17:
#   Add "Segment Mismatch" & "Skew to Bundle" & "Skew to Target" mismatch supported
#   Add ignore syntax, user can ignore the specified check of specified segment
# 2015/12/26:
#   Add Pin in-out relation for NMOS (3-Pin): Source(S) to Drain(D)
# 2015/12/28:
#   Take XTAL as the termination component
#   Exclude test point
# 2016/1/4:
#   Speed up the writing process
#   Add new syntax: Skew to Target
#   A1.40 Release
#   Modify Skew to Target Error
# 2016/1/6:
#   A1.41 Release
# 2016/1/7:
#   Modify "Connect Component" bug
# 2016/1/8:
#   A1.42 Release
# 2016/1/15:
#   Add "GenerateSummary" function
#   A1.45 Release
# 2016/1/15:
#   Add "LoadStackup" & "ClearLayerList" for "Setting" Sheet
# 2016/1/18:
#   Add "LoadTXList" for "Setting" Sheet
# 2016/1/20:
#   Add RESA Pin-in-out mapping for Volga Project
# 2016/1/21:
#   Support polynomial calculation in Check Topology function and ignore syntax
# 2016/1/22:
#   Add the checking function for file integrity of Allegro Report File
#   A1.48 Release
# 2016/1/26~2016/1/27:
#   Simplified the use process by loading the brd directly
#   Add auto-organize feature for Check Topology
# 2016/1/28:
#   A1.50 Release
# 2016/1/29:
#   Fix bug of "Check Topology" function (Error location)
#   A1.51 Release
# 2016/1/29:
#   Add Group Mismatch function
# 2016/2/24: Improve "Load Simple Topology" function
# 2016/3/8:
#   A1.52 Release
# 2016/3/8:
# 1. Add new method for comp_device class to detect the comp layer
# 2. Speed up the LoadAllegroFile function
# 3. Add extracta_brd_data for specified layout data
# 2016/3/9:
# 1. Use extracta_brd_data to add new info of SCH_detect and net_detect function (layer and via)
# 2. Modify the writing style of connected_SCH_Pin_dict in net_detect, add SCH Layer to judge correct connection
# 2016/3/11:
# 1. Remove Component Pin Report
# 2. Solve the possible error of pin connection
# 2016/3/14:
# Modify the writing process of ACT sheet, including the exception part
# 2016/3/15:
# Simplified the content of getpinnumio
# A1.60 Release
# 2016/3/18:
# Support external reading method of pin_mapping
# 2016/3/24:
# Enhance the auto-identify function for diff. pair
# 2016/3/25:
# x86/x64 support
# 2016/3/29:
# 1.Exclude hand error routing case
# 2.Add LOther_CAP Cal. for AMD Platform
# 2016/3/30:
# Add LOther_Via Cal. for AMD Platform
# 2016/3/31:
# A1.62 Release
# 2016/4/5:
# 1.Add Type C CONN support
# 2.Modify the possible error of "Check Topology"
# 2016/4/7:
# Modify warining bug of Check Topology-Layer Change
# 2016/4/8:
# Add specified location of display where segment Mismatch show fail
# 2016/4/9:
# Add UpdateReferenceWorkbook function to update the macro path of excel
# 2016/4/11:
# A1.64 Release
# 2016/4/10:
# 1.Fix the bug of Summary function
# 2.Add fail number and fail rate auto-calculation in summary sheet
# 2016/4/15:
# Add new syntax of CheckTopology for change_layer and cross_over
# 2016/4/20:
# Fix the bug of UpdateReferenceWorkbook
# 2016/4/29:
# Support Micro Via
# 2016/5/3:
# Support multi-Group Mismatch in the same table
# 2016/5/5:
# Support multi-Skew to Target in the same table
# Support multi-ignore syntax for Skew to Target or Group Mismatch(EX:$SK2T1, $SK2T2, $GM1, $GM2,...)
# Add import function of topology table "Load Default Topology Format: Single-Ended" & "Load Default Topology Format: Differential"
# 2016/5/9
# Modify the gui and reduce the function button
# 2016/5/31
# Support Allegro 17.2
# 2016/6/30
# Include the dis-continuity of comp. into the topology check function
# 2016/7/7
# A1.70 Release
# 2016/7/12
# Add new user-defined option in "SymbolList" sheet by "User-Defined Termination {format:SCH Name-Pin}", which can override the default setup
# 2016/7/14
# Support show len feature in Simple Topology
# Add new Signal Type option in setting sheet
# 2016/7/15
# A1.71 Release
# 2016/9/24
# Update for python 3
# 2016/9/25: Add new template fucmm-ccp-ke-rf@mail.foxconn.comnction
# 2016/9/26: Batch Check all Topology

########################################################
# Gorgeous

# 2018.3.11 完善了checklist，使之可运行
# 2018.3.16 发现位于NetDetect()函数中的一个错误：没有筛选出所有明显的差分对。已改正。
# 2018.3.22 改进topology表格形式。
# 2018.3.22 修复四层板中间为NA时，checklist无法显示数据的情况。后发现，这是一种特殊的走线方式，大多存在于DDR中，
# 一般会用Inter工具进行管控，不能通过修改代码进行忽略。解决方法：在setting中修改叠层信息后再进行管控测试。
# 2018.3.31 添加了关于DDR TAB NUMBER管控的模块
# 2018.4.4 添加了计算DDR中DQ_TO_DQ, DQ_TO_DQS, CMD_TO_CLK, CTL_TO_CLK的数据计算及管控模块
# 2018.4.4 添加了管控DIMM_TO_DIMM length的模块
# 2018.4.9 添加了Check Tab, CALC_DQ, CALC_CLK, DIMM_TO_DIMM四个功能按钮的User Guide
# 2018.4.9 更新了NetDelete()中区别差分对的代码，使代码减少40s的运行时间
# 2018.4.10 发现了topology_extract2中初始信号后接三极管与电容无限循环的bug，已解决
# 2018.4.11 发现了topology_extract1中以信号线中间芯片作为起始芯片不能读出走线信息的bug，已解决
# 2018.4.12 发现了topology_extract1中同一条信号线连接起始芯片的不同pin脚，只能读出一条走线信息的bug,已解决
# 2018.4.12 发现了topology_extract1中同一条信号线中间芯片作为起始芯片连接两个pin脚不能读出走线规格的bug，
# 已解决（但是不确定是否正确）
# 2018.4.13 发现了topology_extract2中间信号后接三极管与电容无限循环的bug，已解决
# 2018.4.13 添加了dll_group_length_matching的功能并添加了相应的User Guide
# 2018.4.20 解决了checklist的性能问题：将函数topology_list_format_simplified()中一段代码改成了立即执行函数表达式的形式，
# 使代码速度提高了10的二次方的数量级。结论多用立即执行函数表达式去取代循环判断。
# 2018.4.24 发现signal_net_list = list(set(All_Net_List) ^ set(non_signal_net_list))的方法要优于立即执行函数表达式，
# 运行速度会降低10倍，为Whitely平台制作了特定的DDR表格形式
# 2018.5.2 改进了check topology的功能，将CALC_DQ, CALC_CLK等按钮合并进了check topology按钮里面
# 2018.5.3 发现了segment mismatch功能的bug，进行了代码重写，已解决
# 2018.5.4 发现线太多时GetSetting()函数运行较慢的bug，测试发现是for循环中没有写break，导致每个循环都循环了全部的时间，已改正
# 进而思考其他函数体内有没有类似的情况，已改正
# 2018.5.4 考虑到试用版的checklist工具下的report不应该对外开放，因此添加了加密解密功能模块
# 2018.5.4 改进了NetDetect()函数中运行效率慢的bug，经所有立即执行函数表达式都修改成了集合逻辑操作，减少了代码的运行时间
# 2018.5.22 改进了ACT_se与ACT_diff表中关于分段show出规则的问题，改成将每段信号线及其相邻的信号线全部show出，将每条信号线的
# 最后一个芯片当做终止芯片，缺点：数据过多，可能造成内存占用过多运行速率较慢的问题。
# 2018.5.24 发现22号修改代码不完善，进而完善了代码
# 2018.5.25 解决了segment mismatch功能不完全的问题
# 2018.5.25 发现了同一pin脚连接两个电阻，一个接地一个走线，不能同时读出两条线的bug，已解决

# 2018.5.29 廖俊涛提出的GenerateSummary()中，关于Template也会误读的现象进行了改进 ——KiKi
# 2018.5.30 何鹏提出的anaconda的关于差分对判断标准不规范的bug,已解决 ——KiKi
# 2018.6.1 张珂提出的信号线存在判断fail失误的问题，已解决
# 2018.6.1-2018.6.20 简易版的checklist的制作
# 2018.6.18 - 2018.6.22 check tab功能的改进，添加了加上tab length后的mismatch ——KiKi
# 2018.6.20 - 2018.6.25 checklist评分模块的制作

# 2018.6.27 - 2018.6.28 check tab功能 update，关于DO_TO_DQ之类的mismatch计算逻辑的改变 ——KiKi
# 2018.7.2 checklist评分模块的更新，加入了DFE评分
# 2018.7.3 更新了checklist评分模块（去重，去负分）
# 2018.8.13 添加PCIE对stub的管控 ——KiKi

# 2019.04.14 對余鹏程提出的Netlist界面user define的信號線無法添加到實際走線的bug，進行了改進 ——Kiki
# 2019.04.15 何鵬提出關於單根信號線分叉不完整，兩個元器件共pin時，只會隨機顯示其中一個的問題，已解決
# 2019.05.06 陈敏提出的关于线路回流的问题，已解决
# 2019.05.20 在制作choose_test_point工具时，发现从 Diffpair Gap Report 中获取的差分对信号线组并不完整，因此作出修改 ——Gorgeous
# 2019.05.23 为了防止连线坐标点的误差，将坐标点的坐标值取一位小数 ——Gorgeous
# 2019.05.27 由于无法获得正确的电源线信息，无法准确分类出电源线与单根线，所以将所有带有\dV字样的信号线全部归为单根线 ——Gorgeous
# 2019.06.10 将Netlist中的User Defined single net无法添加到single net中的bug进行了修复 ——Gorgeous
# 2019.06.12 添加了可能一根线的一段同时连接3个pin的情况，修改了net_detect中对于pin中点和线端点范围的值，
# 修改了topology_format中的表达方式，将三个芯片依次作为终点显示 ——Gorgeous
# 2019.07.09 更新Summary表单，删除重复的错误显示 --kiki
# 2019.07.15 批量更新Topology --kiki
# 2019.08.06 添加代码 net_list_temp.append(None) next_pin_list_temp.append(None) ——Daniel
# 2019.08.10 添代码  topology_out_list.append(topology_list[0])， or next_net[idxx_] == start_net_name，debug ACT Topology 缺少短分支的问题 -- Daniel
# 2019.08.26 添加topology数据中cross分叉点标示，update check topology中check cross over功能 --Daniel
# 2019.09.11 更改了两元器件三元器件共Pin的数据出问题的bug ——Gorgeous
# 2019.09.12 update批量更新功能 ——Kiki


from xlwings import Book
import xlwings as xw
import re, sys, os, time, datetime, copy
from PySide import QtGui, QtCore
import time
# from Crypto.Cipher import AES
from binascii import b2a_hex, a2b_hex
from Tkinter import _flatten
import sys
from xlwings import Book
import xlwings as xw
import re, sys, os, time, datetime, copy
from PySide import QtGui, QtCore
import time
# from Crypto.Cipher import AES
from binascii import b2a_hex, a2b_hex
from Tkinter import _flatten
import sys


# 加密代码
# class MyCrypt():
#     def __init__(self, key):
#         self.key = key
#         self.mode = AES.MODE_CBC
#
#     def myencrypt(self, text):
#         length = 16
#         count = len(text)
#         # # print(count)
#
#         if count < length:
#             add = length - count
#             text = text + ('\0' * add)
#         elif count > length:
#             add = (length - (count % length))
#             text = text + ('\0' * add)
#
#         # # print(len(text))
#
#         cryptor = AES.new(self.key, self.mode, b'0000000000000000')
#         self.ciphertext = cryptor.encrypt(text)
#         return b2a_hex(self.ciphertext)
#
#     def mydecrypt(self, text):
#         cryptor = AES.new(self.key, self.mode, b'0000000000000000')
#         plain_text = cryptor.decrypt(a2b_hex(text))
#         return plain_text.rstrip('\0')
# Class Part
# 返回一个类对象
class comp_device:
    def __init__(self, name, model_info, pin_, etch_, pin_net_, pinx_, piny_, coNNsch_):
        self.name = name
        self.model_info = model_info
        self.pin_ = tuple(sorted(list(pin_)))
        self.etch_ = etch_
        self.pin_net_ = pin_net_
        self.pinx_ = pinx_
        self.piny_ = piny_
        self.coNNsch_ = coNNsch_

    def GetName(self):
        return self.name

    def GetModel(self):
        return self.model_info

    def GetPinList(self):
        return self.pin_

    def GetNet(self, pin_n):
        return self.pin_net_.get(pin_n)

    def GetNetList(self):
        return self.etch_

    def GetXPoint(self, pin_n):
        return self.pinx_[pin_n]

    def GetYPoint(self, pin_n):
        return self.piny_[pin_n]

    def GetXY(self, pin_n):
        return (self.pinx_[pin_n], self.piny_[pin_n])


class etch_line:
    def __init__(self, etch_, seg_set, segwidth, seglen, seglayer, segp1, segp2, conn_schpin):
        self.etch_ = etch_
        self.seg_set = seg_set
        self.segwidth = segwidth
        self.seglen = seglen
        self.seglayer = seglayer
        self.segp1 = segp1
        self.segp2 = segp2
        self.conn_schpin = conn_schpin

    def GetName(self):
        return self.etch_

    def GetSegmentList(self):
        return self.seg_set

    def GetWidth(self, seg_id):
        return self.segwidth[seg_id]

    def GetLength(self, seg_id):
        return self.seglen[seg_id]

    def GetLayer(self, seg_id):
        return self.seglayer[seg_id]

    def GetXY1(self, seg_id):
        return self.segp1[seg_id]

    def GetXY2(self, seg_id):
        return self.segp2[seg_id]

    def GetConnectedSCHList(self, seg_id):
        output_sch = list()
        # {('R2', '1'): [[('1710.14', '1076.50'), (1, 2)]], ('U1', 'Y8'): [[('1673.68', '1124.77'), (2, 2)]]}
        for (SCH_name, pin_id), pin_seg_info in self.conn_schpin.items():
            # # print('get', (SCH_name, pin_id), pin_seg_info)
            for idx in xrange(len(pin_seg_info)):
                pinpoint = pin_seg_info[idx][0]
                seg_ind = pin_seg_info[idx][1]
                if seg_ind and seg_ind[0] == seg_id:
                    output_sch.append((SCH_name, pin_id, pinpoint, seg_ind))
                    # eg: ('R2', '1', ('1710.14', '1076.50'), (1, 2))
        return output_sch

    # 根据坐标点是否是线所连pin脚点来判断
    def GetConnComp(self, seg_ind_input):
        for (SCH_name, pin_id), pin_seg_info in self.conn_schpin.items():
            # # print((SCH_name, pin_id), pin_seg_info)
            for idx in xrange(len(pin_seg_info)):
                seg_ind = pin_seg_info[idx][1]
                # # print('seg_ind_input', seg_ind, seg_ind_input)
                if seg_ind and seg_ind == seg_ind_input:
                    return (SCH_name, pin_id)

        return None

    # 返回与芯片相连的信号线的芯片名称与pin脚id
    def GetConnectedSCHListBySegInd(self, seg_ind_input):
        connected_SCH_list = list()
        count = 0
        for (SCH_name, pin_id), pin_seg_info in self.conn_schpin.items():
            # if SCH_name == 'R737':
            #     for idx in xrange(len(pin_seg_info)):
            #         seg_ind = pin_seg_info[idx][1]
            # # print('R737', seg_ind, seg_ind_input)
            # # print('get', (SCH_name, pin_id), pin_seg_info)
            # # print('seg_ind_input', seg_ind_input)
            count += 1
            for idx in xrange(len(pin_seg_info)):
                seg_ind = pin_seg_info[idx][1]
                if seg_ind and seg_ind == seg_ind_input:
                    connected_SCH_list.append((SCH_name, pin_id))
            #
            # if SCH_name == 'R737':
            #     # print(connected_SCH_list)
        if connected_SCH_list != []:
            return connected_SCH_list
        else:
            return None

    # def GetSchPinListBySegInd(self,seg_ind_input):
    #     sch_pin_list = []
    #     for (SCH_name, pin_id), pin_seg_info in self.conn_schpin.items():
    #         if list(seg_ind_input) == list(pin_seg_info[0][-1]):
    #             return (SCH_name, pin_id)


# 转换数据：转换列表为对象
class brd_data:
    def __init__(self, data):
        self.data = data

    def GetData(self):
        return self.data


# 生成topology表格
class TopologyLayoutForm(QtGui.QDialog):
    def __init__(self, t_typ, parent=None):
        super(TopologyLayoutForm, self).__init__(parent)
        self.t_typ = t_typ
        self.setWindowTitle('Topology Table')

        self.label0 = QtGui.QLabel('<b>Signal Type: %s</b>' % t_typ, self)

        self.cb1 = QtGui.QCheckBox("Total Length", self)
        # 默认勾选
        self.cb1.setCheckState(QtCore.Qt.Checked)

        if t_typ == 'Differential':
            self.cb7 = QtGui.QCheckBox("Layer Mismatch", self)
            self.cb8 = QtGui.QCheckBox("Segment Mismatch", self)
            self.cb9 = QtGui.QCheckBox("Total Mismatch", self)
            self.cb10 = QtGui.QCheckBox("DQS To DQS", self)
            # self.cb11 = QtGui.QCheckBox("DIMM To DIMM", self)
        else:
            self.cb7 = QtGui.QCheckBox("DQ TO DQ", self)
            self.cb8 = QtGui.QCheckBox("DQ TO DQS", self)
            self.cb9 = QtGui.QCheckBox("CMD TO CLK", self)
            self.cb10 = QtGui.QCheckBox("CTL TO CLK", self)
            self.cb11 = QtGui.QCheckBox("DLL Group", self)

        self.cb4 = QtGui.QCheckBox("Skew to Target", self)
        self.cb5 = QtGui.QCheckBox("Skew to Bundle", self)
        self.cb6 = QtGui.QCheckBox("Group Mismatch", self)

        self.cb2 = QtGui.QCheckBox("Via Count", self)
        # 默认勾选
        self.cb2.setCheckState(QtCore.Qt.Checked)
        # 选择是否是DDR的特殊表格
        self.cb3 = QtGui.QCheckBox("For DDR", self)

        self.button1 = QtGui.QPushButton('Insert', self)
        self.button1.clicked.connect(self.load_topology_format)
        self.button2 = QtGui.QPushButton('Close', self)
        self.button2.clicked.connect(self.close_dialog)

        layout = QtGui.QGridLayout()
        layout.addWidget(self.label0, 0, 0)
        layout.addWidget(self.cb1, 1, 0)
        layout.addWidget(self.cb2, 1, 1)

        layout.addWidget(self.cb4, 2, 0)
        layout.addWidget(self.cb5, 2, 1)

        layout.addWidget(self.cb6, 3, 0)

        if t_typ == 'Differential':
            layout.addWidget(self.cb3, 4, 1)
            layout.addWidget(self.cb7, 3, 1)
            layout.addWidget(self.cb8, 3, 2)
            layout.addWidget(self.cb9, 3, 3)
            layout.addWidget(self.cb10, 4, 0)
        else:
            layout.addWidget(self.cb3, 5, 0)
            layout.addWidget(self.cb7, 4, 0)
            layout.addWidget(self.cb8, 4, 1)
            layout.addWidget(self.cb9, 4, 2)
            layout.addWidget(self.cb10, 4, 3)
            layout.addWidget(self.cb11, 3, 1)

        layout.addWidget(self.button1, 6, 0)
        layout.addWidget(self.button2, 6, 1)

        self.setLayout(layout)

    def close_dialog(self):
        # 相当于跳过
        QtGui.QDialog.accept(self)

    def load_topology_format(self):

        wb = Book(xlsm_path).caller()
        active_sheet = wb.sheets.active  # Get the active sheet object

        selection_range = wb.app.selection
        start_ind = (selection_range.row, selection_range.column)
        col_index = start_ind[1]

        col_title_list1 = ['Start Segment Name', 'Layer', 'Layer Change', 'Connect Component', 'Cross Over',
                           'Trace Width', 'Space', 'Min', 'Max', 'Net Name']

        active_sheet.range(start_ind).value = 'Topology'

        # 为特定的DDR表格设置标记
        if self.cb3.isChecked():
            active_sheet.range(start_ind[0], start_ind[1] + 1).value = 'For DDR'

        active_sheet.range((start_ind[0] + 1, col_index)).value = 'Signal Type'
        active_sheet.range((start_ind[0] + 1, col_index + 1)).value = self.t_typ

        active_sheet.range((start_ind[0] + 2, col_index)).value = 'Start Component Name-Pin Number'
        active_sheet.range((start_ind[0] + 2, col_index)).api.Interior.ColorIndex = 43
        # 合并单元格
        active_sheet.range((start_ind[0] + 2, col_index), (start_ind[0] + 2 + 9, start_ind[1])).api.MergeCells = True
        # 设置垂直居中
        active_sheet.range((start_ind[0] + 2, col_index)).api.VerticalAlignment = Constants.xlCenter

        col_index += 1
        active_sheet.range((start_ind[0] + 2, col_index)).value = 'End Component Name-Pin Number'
        active_sheet.range((start_ind[0] + 2, col_index)).api.Interior.ColorIndex = 43
        active_sheet.range((start_ind[0] + 2, col_index),
                           (start_ind[0] + 2 + 9, start_ind[1] + 1)).api.MergeCells = True
        active_sheet.range((start_ind[0] + 2, col_index)).api.VerticalAlignment = Constants.xlCenter

        col_index += 1
        for idx in xrange(len(col_title_list1)):
            active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_title_list1[idx]
            active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.ColorIndex = 44

        if self.cb3.isChecked():
            col_content_list1 = ['DOGBONE', 'MS', 'N', 'N', 'N', '4', 'NA', '0', '50', 'Length(mils)']
        else:
            col_content_list1 = ['BO', 'MS/SL/DSL', 'N', 'N', 'N', '3.5', 'NA', '0', '500', 'Length(mils)']
        col_index += 1
        for idx in xrange(len(col_content_list1)):
            active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list1[idx]
            if idx in [2, 3, 4]:
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightSteelBlue
            else:
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbSkyBlue

        if self.cb3.isChecked():
            col_content_list2 = ['CPU_PINFIELD', 'MS/SL/DSL', 'N', 'N', 'N', '3.5', 'NA', '0', '2000', 'Length(mils)']
        else:
            col_content_list2 = ['M1', 'MS/SL/DSL', 'N', 'N', 'N', '4/5/5.3', 'NA', '0', '2000', 'Length(mils)']
        col_index += 1
        for idx in xrange(len(col_content_list2)):
            active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list2[idx]
            if idx in [2, 3, 4]:
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightSteelBlue
            else:
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbSkyBlue

        if self.cb3.isChecked():
            col_content_list3 = ['MAIN_TRACE', 'MS/SL/DSL', 'N', 'N', 'N', '4/5/5.3', 'NA', '0', '2000', 'Length(mils)']
            col_index += 1
            for idx in xrange(len(col_content_list3)):
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list3[idx]
                if idx in [2, 3, 4]:
                    active_sheet.range(
                        (start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightSteelBlue
                else:
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbSkyBlue
        else:
            col_index += 1
            active_sheet.range((start_ind[0] + 2, col_index)).value = 'Component Name'
            active_sheet.range((start_ind[0] + 2, col_index)).api.Interior.ColorIndex = 43
            # 合并单元格
            active_sheet.range((start_ind[0] + 2, col_index), (start_ind[0] + 2 + 9, col_index)).api.MergeCells = True
            active_sheet.range((start_ind[0] + 2, col_index)).api.VerticalAlignment = Constants.xlCenter

            col_index += 1
            for idx in xrange(len(col_title_list1)):
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_title_list1[idx]
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.ColorIndex = 44

            col_content_list3 = ['M2', 'MS/SL/DSL', 'N', 'N', 'N', '4/5/5.3', 'NA', '0', '2000', 'Length(mils)']
            col_index += 1
            for idx in xrange(len(col_content_list3)):
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list3[idx]
                if idx in [2, 3, 4]:
                    active_sheet.range(
                        (start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightSteelBlue
                else:
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbSkyBlue

            col_content_list4 = ['BI', 'MS/SL/DSL', 'N', 'N', 'N', '3.5', 'NA', '0', '500', 'Length(mils)']
            col_index += 1
            for idx in xrange(len(col_content_list4)):
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list4[idx]
                if idx in [2, 3, 4]:
                    active_sheet.range(
                        (start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightSteelBlue
                else:
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbSkyBlue

            col_content_list5 = ['BO+M1', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '2000', 'Length(mils)']
            col_index += 1
            for idx in xrange(len(col_content_list5)):
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list5[idx]
                if idx in [2, 3, 4]:
                    active_sheet.range(
                        (start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightSteelBlue
                else:
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbSkyBlue

        if self.cb1.isChecked():
            col_index += 1
            col_content_list = ['Total Length', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '3000', 'Length(mils)']
            for idx in xrange(len(col_content_list)):
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

        if self.cb6.isChecked():
            col_index += 1
            col_content_list = ['Group Mismatch', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '500', 'Max Mismatch(mils)']
            for idx in xrange(len(col_content_list)):
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

        if self.t_typ == 'Differential':

            if self.cb7.isChecked():
                col_index += 1
                col_content_list = ['Layer Mismatch', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '10',
                                    'Max Mismatch(mils)']
                for idx in xrange(len(col_content_list)):
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

            if self.cb8.isChecked():
                col_index += 1
                col_content_list = ['Segment Mismatch', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '10',
                                    'Max Mismatch(mils)']
                for idx in xrange(len(col_content_list)):
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

            if self.cb9.isChecked():
                col_index += 1
                col_content_list = ['Total Mismatch', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '5', 'Mismatch(mils)']
                for idx in xrange(len(col_content_list)):
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

            if self.cb10.isChecked():
                col_index += 1
                col_content_list = ['Relative Length Spec(DQS to DQS)', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '5',
                                    'Mismatch(mils)']
                for idx in xrange(len(col_content_list)):
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

        else:
            if self.cb7.isChecked():
                col_index += 1
                col_content_list = ['Relative Length Spec(DQ to DQ)', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '100',
                                    'Mismatch(mils)']
                for idx in xrange(len(col_content_list)):
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

            if self.cb8.isChecked():
                col_index += 1
                col_content_list = ['Relative Length Spec(DQ to DQS)', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '300',
                                    'Mismatch(mils)']
                for idx in xrange(len(col_content_list)):
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

            if self.cb9.isChecked():
                col_index += 1
                col_content_list = ['CMD or ADD to CLK Length Matching', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '700',
                                    'Mismatch(mils)']
                for idx in xrange(len(col_content_list)):
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

            if self.cb10.isChecked():
                col_index += 1
                col_content_list = ['CTL to CLK Length Matching', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '700',
                                    'Mismatch(mils)']
                for idx in xrange(len(col_content_list)):
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

            if self.cb11.isChecked():
                col_index += 1
                col_content_list = ['DIMM TO DIMM Length', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '500',
                                    'Mismatch(mils)']
                for idx in xrange(len(col_content_list)):
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                    active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

        if self.cb4.isChecked():
            col_index += 1
            col_content_list = ['Skew to Target', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '500', 'Max Mismatch(mils)']
            for idx in xrange(len(col_content_list)):
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

        if self.cb5.isChecked():
            col_index += 1
            col_content_list = ['Skew to Bundle', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '50', 'Max Mismatch(mils)']
            for idx in xrange(len(col_content_list)):
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

        if self.cb2.isChecked():
            col_index += 1
            col_content_list = ['Via Count', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', '0', '2', 'Via Count']
            for idx in xrange(len(col_content_list)):
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
                active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbLightBlue

        col_index += 1
        col_content_list = ['Result', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA']
        for idx in xrange(len(col_content_list)):
            active_sheet.range((start_ind[0] + 2 + idx, col_index)).value = col_content_list[idx]
            active_sheet.range((start_ind[0] + 2 + idx, col_index)).api.Interior.Color = RgbColor.rgbSkyBlue

        SetCellFont_current_region(active_sheet, start_ind, 'Times New Roman', 12, 'l')
        SetCellBorder_current_region(active_sheet, start_ind)
        active_sheet.autofit('c')

        active_sheet.range((start_ind[0], start_ind[1] + 2), (start_ind[0] + 1, col_index)).api.MergeCells = True

        QtGui.QDialog.accept(self)


# 选择表格类型
class TopologyTypeSelection(QtGui.QDialog):
    def __init__(self, parent=None):
        super(TopologyTypeSelection, self).__init__(parent)
        self.setWindowTitle('Select Topology Type or Convert Old Format')
        self.se_select = TopologyLayoutForm('Single-ended')
        self.diff_select = TopologyLayoutForm('Differential')

        self.button1 = QtGui.QPushButton('Single-ended Type', self)
        self.button1.clicked.connect(self.se)
        self.button2 = QtGui.QPushButton('Differential Type', self)
        self.button2.clicked.connect(self.diff)
        self.button3 = QtGui.QPushButton('Convert Old Format', self)
        self.button3.clicked.connect(self.convertformat)
        self.button4 = QtGui.QPushButton('Cancel', self)
        self.button4.clicked.connect(self.close_dialog)

        layout = QtGui.QGridLayout()
        layout.addWidget(self.button1, 0, 0)
        layout.addWidget(self.button2, 1, 0)
        layout.addWidget(self.button3, 2, 0)
        layout.addWidget(self.button4, 2, 1)
        self.setLayout(layout)

    def close_dialog(self):
        QtGui.QDialog.accept(self)

    def se(self):
        QtGui.QDialog.accept(self)
        self.se_select.show()

    def diff(self):
        QtGui.QDialog.accept(self)
        self.diff_select.show()

    def convertformat(self):
        QtGui.QDialog.accept(self)
        FormatConverter()


# 生成Simple Topology表格
class SimpleTopology(QtGui.QDialog):
    def __init__(self, parent=None):
        super(SimpleTopology, self).__init__(parent)
        self.setWindowTitle('Simple Topology')
        self.b111 = QtGui.QPushButton('Insert Format', self)
        self.b111.clicked.connect(self.LoadFormat)
        self.b222 = QtGui.QPushButton('Load Topology', self)
        self.b222.clicked.connect(self.LoadTopology)
        self.b333 = QtGui.QPushButton('Cancel', self)
        self.b333.clicked.connect(self.close_dialog)

        layout = QtGui.QGridLayout()
        layout.addWidget(self.b111, 0, 0)
        layout.addWidget(self.b222, 1, 0)
        layout.addWidget(self.b333, 2, 0)
        self.setLayout(layout)

    def close_dialog(self):
        QtGui.QDialog.accept(self)

    def LoadFormat(self):
        QtGui.QDialog.accept(self)
        LoadSimpleTopologyFormat()

    def LoadTopology(self):
        QtGui.QDialog.accept(self)
        LoadSimpleTopology()


class Constants:
    xlNextToAxis = 4  # from enum Constants
    xlNoDocuments = 3  # from enum Constants
    xlNone = -4142  # from enum Constants
    xlNotes = -4144  # from enum Constants
    xlOff = -4146  # from enum Constants
    xl3DEffects1 = 13  # from enum Constants
    xl3DBar = -4099  # from enum Constants
    xl3DEffects2 = 14  # from enum Constants
    xl3DSurface = -4103  # from enum Constants
    xlAbove = 0  # from enum Constants
    xlAccounting1 = 4  # from enum Constants
    xlAccounting2 = 5  # from enum Constants
    xlAccounting3 = 6  # from enum Constants
    xlAccounting4 = 17  # from enum Constants
    xlAdd = 2  # from enum Constants
    xlAll = -4104  # from enum Constants
    xlAllExceptBorders = 7  # from enum Constants
    xlAutomatic = -4105  # from enum Constants
    xlBar = 2  # from enum Constants
    xlBelow = 1  # from enum Constants
    xlBidi = -5000  # from enum Constants
    xlBidiCalendar = 3  # from enum Constants
    xlBoth = 1  # from enum Constants
    xlBottom = -4107  # from enum Constants
    xlCascade = 7  # from enum Constants
    xlCenter = -4108  # from enum Constants
    xlCenterAcrossSelection = 7  # from enum Constants
    xlChart4 = 2  # from enum Constants
    xlChartSeries = 17  # from enum Constants
    xlChartShort = 6  # from enum Constants
    xlChartTitles = 18  # from enum Constants
    xlChecker = 9  # from enum Constants
    xlCircle = 8  # from enum Constants
    xlClassic1 = 1  # from enum Constants
    xlClassic2 = 2  # from enum Constants
    xlClassic3 = 3  # from enum Constants
    xlClosed = 3  # from enum Constants
    xlColor1 = 7  # from enum Constants
    xlColor2 = 8  # from enum Constants
    xlColor3 = 9  # from enum Constants
    xlColumn = 3  # from enum Constants
    xlCombination = -4111  # from enum Constants
    xlComplete = 4  # from enum Constants
    xlConstants = 2  # from enum Constants
    xlContents = 2  # from enum Constants
    xlContext = -5002  # from enum Constants
    xlCorner = 2  # from enum Constants
    xlCrissCross = 16  # from enum Constants
    xlCross = 4  # from enum Constants
    xlCustom = -4114  # from enum Constants
    xlDebugCodePane = 13  # from enum Constants
    xlDefaultAutoFormat = -1  # from enum Constants
    xlDesktop = 9  # from enum Constants
    xlDiamond = 2  # from enum Constants
    xlDirect = 1  # from enum Constants
    xlDistributed = -4117  # from enum Constants
    xlDivide = 5  # from enum Constants
    xlDoubleAccounting = 5  # from enum Constants
    xlDoubleClosed = 5  # from enum Constants
    xlDoubleOpen = 4  # from enum Constants
    xlDoubleQuote = 1  # from enum Constants
    xlDrawingObject = 14  # from enum Constants
    xlEntireChart = 20  # from enum Constants
    xlExcelMenus = 1  # from enum Constants
    xlExtended = 3  # from enum Constants
    xlFill = 5  # from enum Constants
    xlFirst = 0  # from enum Constants
    xlFixedValue = 1  # from enum Constants
    xlFloating = 5  # from enum Constants
    xlFormats = -4122  # from enum Constants
    xlFormula = 5  # from enum Constants
    xlFullScript = 1  # from enum Constants
    xlGeneral = 1  # from enum Constants
    xlGray16 = 17  # from enum Constants
    xlGray25 = -4124  # from enum Constants
    xlGray50 = -4125  # from enum Constants
    xlGray75 = -4126  # from enum Constants
    xlGray8 = 18  # from enum Constants
    xlGregorian = 2  # from enum Constants
    xlGrid = 15  # from enum Constants
    xlGridline = 22  # from enum Constants
    xlHigh = -4127  # from enum Constants
    xlHindiNumerals = 3  # from enum Constants
    xlIcons = 1  # from enum Constants
    xlImmediatePane = 12  # from enum Constants
    xlInside = 2  # from enum Constants
    xlInteger = 2  # from enum Constants
    xlJustify = -4130  # from enum Constants
    xlLTR = -5003  # from enum Constants
    xlLast = 1  # from enum Constants
    xlLastCell = 11  # from enum Constants
    xlLatin = -5001  # from enum Constants
    xlLeft = -4131  # from enum Constants
    xlLeftToRight = 2  # from enum Constants
    xlLightDown = 13  # from enum Constants
    xlLightHorizontal = 11  # from enum Constants
    xlLightUp = 14  # from enum Constants
    xlLightVertical = 12  # from enum Constants
    xlList1 = 10  # from enum Constants
    xlList2 = 11  # from enum Constants
    xlList3 = 12  # from enum Constants
    xlLocalFormat1 = 15  # from enum Constants
    xlLocalFormat2 = 16  # from enum Constants
    xlLogicalCursor = 1  # from enum Constants
    xlLong = 3  # from enum Constants
    xlLotusHelp = 2  # from enum Constants
    xlLow = -4134  # from enum Constants
    xlMacrosheetCell = 7  # from enum Constants
    xlManual = -4135  # from enum Constants
    xlMaximum = 2  # from enum Constants
    xlMinimum = 4  # from enum Constants
    xlMinusValues = 3  # from enum Constants
    xlMixed = 2  # from enum Constants
    xlMixedAuthorizedScript = 4  # from enum Constants
    xlMixedScript = 3  # from enum Constants
    xlModule = -4141  # from enum Constants
    xlMultiply = 4  # from enum Constants
    xlNarrow = 1  # from enum Constants
    xlOn = 1  # from enum Constants
    xlOpaque = 3  # from enum Constants
    xlOpen = 2  # from enum Constants
    xlOutside = 3  # from enum Constants
    xlPartial = 3  # from enum Constants
    xlPartialScript = 2  # from enum Constants
    xlPercent = 2  # from enum Constants
    xlPlus = 9  # from enum Constants
    xlPlusValues = 2  # from enum Constants
    xlRTL = -5004  # from enum Constants
    xlReference = 4  # from enum Constants
    xlRight = -4152  # from enum Constants
    xlScale = 3  # from enum Constants
    xlSemiGray75 = 10  # from enum Constants
    xlSemiautomatic = 2  # from enum Constants
    xlShort = 1  # from enum Constants
    xlShowLabel = 4  # from enum Constants
    xlShowLabelAndPercent = 5  # from enum Constants
    xlShowPercent = 3  # from enum Constants
    xlShowValue = 2  # from enum Constants
    xlSimple = -4154  # from enum Constants
    xlSingle = 2  # from enum Constants
    xlSingleAccounting = 4  # from enum Constants
    xlSingleQuote = 2  # from enum Constants
    xlSolid = 1  # from enum Constants
    xlSquare = 1  # from enum Constants
    xlStError = 4  # from enum Constants
    xlStar = 5  # from enum Constants
    xlStrict = 2  # from enum Constants
    xlSubtract = 3  # from enum Constants
    xlSystem = 1  # from enum Constants
    xlTextBox = 16  # from enum Constants
    xlTiled = 1  # from enum Constants
    xlTitleBar = 8  # from enum Constants
    xlToolbar = 1  # from enum Constants
    xlToolbarButton = 2  # from enum Constants
    xlTop = -4160  # from enum Constants
    xlTopToBottom = 1  # from enum Constants
    xlTransparent = 2  # from enum Constants
    xlTriangle = 3  # from enum Constants
    xlVeryHidden = 2  # from enum Constants
    xlVisible = 12  # from enum Constants
    xlVisualCursor = 2  # from enum Constants
    xlWatchPane = 11  # from enum Constants
    xlWide = 3  # from enum Constants
    xlWorkbookTab = 6  # from enum Constants
    xlWorksheet4 = 1  # from enum Constants
    xlWorksheetCell = 3  # from enum Constants
    xlWorksheetShort = 5  # from enum Constants


class LineStyle:
    xlContinuous = 1  # from enum XlLineStyle
    xlDash = -4115  # from enum XlLineStyle
    xlDashDot = 4  # from enum XlLineStyle
    xlDashDotDot = 5  # from enum XlLineStyle
    xlDot = -4118  # from enum XlLineStyle
    xlDouble = -4119  # from enum XlLineStyle
    xlLineStyleNone = -4142  # from enum XlLineStyle
    xlSlantDashDot = 13  # from enum XlLineStyle


class RgbColor:
    rgbAliceBlue = 16775408  # from enum XlRgbColor
    rgbAntiqueWhite = 14150650  # from enum XlRgbColor
    rgbAqua = 16776960  # from enum XlRgbColor
    rgbAquamarine = 13959039  # from enum XlRgbColor
    rgbAzure = 16777200  # from enum XlRgbColor
    rgbBeige = 14480885  # from enum XlRgbColor
    rgbBisque = 12903679  # from enum XlRgbColor
    rgbBlack = 0  # from enum XlRgbColor
    rgbBlanchedAlmond = 13495295  # from enum XlRgbColor
    rgbBlue = 16711680  # from enum XlRgbColor
    rgbBlueViolet = 14822282  # from enum XlRgbColor
    rgbBrown = 2763429  # from enum XlRgbColor
    rgbBurlyWood = 8894686  # from enum XlRgbColor
    rgbCadetBlue = 10526303  # from enum XlRgbColor
    rgbChartreuse = 65407  # from enum XlRgbColor
    rgbCoral = 5275647  # from enum XlRgbColor
    rgbCornflowerBlue = 15570276  # from enum XlRgbColor
    rgbCornsilk = 14481663  # from enum XlRgbColor
    rgbCrimson = 3937500  # from enum XlRgbColor
    rgbDarkBlue = 9109504  # from enum XlRgbColor
    rgbDarkCyan = 9145088  # from enum XlRgbColor
    rgbDarkGoldenrod = 755384  # from enum XlRgbColor
    rgbDarkGray = 11119017  # from enum XlRgbColor
    rgbDarkGreen = 25600  # from enum XlRgbColor
    rgbDarkGrey = 11119017  # from enum XlRgbColor
    rgbDarkKhaki = 7059389  # from enum XlRgbColor
    rgbDarkMagenta = 9109643  # from enum XlRgbColor
    rgbDarkOliveGreen = 3107669  # from enum XlRgbColor
    rgbDarkOrange = 36095  # from enum XlRgbColor
    rgbDarkOrchid = 13382297  # from enum XlRgbColor
    rgbDarkRed = 139  # from enum XlRgbColor
    rgbDarkSalmon = 8034025  # from enum XlRgbColor
    rgbDarkSeaGreen = 9419919  # from enum XlRgbColor
    rgbDarkSlateBlue = 9125192  # from enum XlRgbColor
    rgbDarkSlateGray = 5197615  # from enum XlRgbColor
    rgbDarkSlateGrey = 5197615  # from enum XlRgbColor
    rgbDarkTurquoise = 13749760  # from enum XlRgbColor
    rgbDarkViolet = 13828244  # from enum XlRgbColor
    rgbDeepPink = 9639167  # from enum XlRgbColor
    rgbDeepSkyBlue = 16760576  # from enum XlRgbColor
    rgbDimGray = 6908265  # from enum XlRgbColor
    rgbDimGrey = 6908265  # from enum XlRgbColor
    rgbDodgerBlue = 16748574  # from enum XlRgbColor
    rgbFireBrick = 2237106  # from enum XlRgbColor
    rgbFloralWhite = 15792895  # from enum XlRgbColor
    rgbForestGreen = 2263842  # from enum XlRgbColor
    rgbFuchsia = 16711935  # from enum XlRgbColor
    rgbGainsboro = 14474460  # from enum XlRgbColor
    rgbGhostWhite = 16775416  # from enum XlRgbColor
    rgbGold = 55295  # from enum XlRgbColor
    rgbGoldenrod = 2139610  # from enum XlRgbColor
    rgbGray = 8421504  # from enum XlRgbColor
    rgbGreen = 32768  # from enum XlRgbColor
    rgbGreenYellow = 3145645  # from enum XlRgbColor
    rgbGrey = 8421504  # from enum XlRgbColor
    rgbHoneydew = 15794160  # from enum XlRgbColor
    rgbHotPink = 11823615  # from enum XlRgbColor
    rgbIndianRed = 6053069  # from enum XlRgbColor
    rgbIndigo = 8519755  # from enum XlRgbColor
    rgbIvory = 15794175  # from enum XlRgbColor
    rgbKhaki = 9234160  # from enum XlRgbColor
    rgbLavender = 16443110  # from enum XlRgbColor
    rgbLavenderBlush = 16118015  # from enum XlRgbColor
    rgbLawnGreen = 64636  # from enum XlRgbColor
    rgbLemonChiffon = 13499135  # from enum XlRgbColor
    rgbLightBlue = 15128749  # from enum XlRgbColor
    rgbLightCoral = 8421616  # from enum XlRgbColor
    rgbLightCyan = 9145088  # from enum XlRgbColor
    rgbLightGoldenrodYellow = 13826810  # from enum XlRgbColor
    rgbLightGray = 13882323  # from enum XlRgbColor
    rgbLightGreen = 9498256  # from enum XlRgbColor
    rgbLightGrey = 13882323  # from enum XlRgbColor
    rgbLightPink = 12695295  # from enum XlRgbColor
    rgbLightSalmon = 8036607  # from enum XlRgbColor
    rgbLightSeaGreen = 11186720  # from enum XlRgbColor
    rgbLightSkyBlue = 16436871  # from enum XlRgbColor
    rgbLightSlateGray = 10061943  # from enum XlRgbColor
    rgbLightSlateGrey = 10061943  # from enum XlRgbColor
    rgbLightSteelBlue = 14599344  # from enum XlRgbColor
    rgbLightYellow = 14745599  # from enum XlRgbColor
    rgbLime = 65280  # from enum XlRgbColor
    rgbLimeGreen = 3329330  # from enum XlRgbColor
    rgbLinen = 15134970  # from enum XlRgbColor
    rgbMaroon = 128  # from enum XlRgbColor
    rgbMediumAquamarine = 11206502  # from enum XlRgbColor
    rgbMediumBlue = 13434880  # from enum XlRgbColor
    rgbMediumOrchid = 13850042  # from enum XlRgbColor
    rgbMediumPurple = 14381203  # from enum XlRgbColor
    rgbMediumSeaGreen = 7451452  # from enum XlRgbColor
    rgbMediumSlateBlue = 15624315  # from enum XlRgbColor
    rgbMediumSpringGreen = 10156544  # from enum XlRgbColor
    rgbMediumTurquoise = 13422920  # from enum XlRgbColor
    rgbMediumVioletRed = 8721863  # from enum XlRgbColor
    rgbMidnightBlue = 7346457  # from enum XlRgbColor
    rgbMintCream = 16449525  # from enum XlRgbColor
    rgbMistyRose = 14804223  # from enum XlRgbColor
    rgbMoccasin = 11920639  # from enum XlRgbColor
    rgbNavajoWhite = 11394815  # from enum XlRgbColor
    rgbNavy = 8388608  # from enum XlRgbColor
    rgbNavyBlue = 8388608  # from enum XlRgbColor
    rgbOldLace = 15136253  # from enum XlRgbColor
    rgbOlive = 32896  # from enum XlRgbColor
    rgbOliveDrab = 2330219  # from enum XlRgbColor
    rgbOrange = 42495  # from enum XlRgbColor
    rgbOrangeRed = 17919  # from enum XlRgbColor
    rgbOrchid = 14053594  # from enum XlRgbColor
    rgbPaleGoldenrod = 7071982  # from enum XlRgbColor
    rgbPaleGreen = 10025880  # from enum XlRgbColor
    rgbPaleTurquoise = 15658671  # from enum XlRgbColor
    rgbPaleVioletRed = 9662683  # from enum XlRgbColor
    rgbPapayaWhip = 14020607  # from enum XlRgbColor
    rgbPeachPuff = 12180223  # from enum XlRgbColor
    rgbPeru = 4163021  # from enum XlRgbColor
    rgbPink = 13353215  # from enum XlRgbColor
    rgbPlum = 14524637  # from enum XlRgbColor
    rgbPowderBlue = 15130800  # from enum XlRgbColor
    rgbPurple = 8388736  # from enum XlRgbColor
    rgbRed = 255  # from enum XlRgbColor
    rgbRosyBrown = 9408444  # from enum XlRgbColor
    rgbRoyalBlue = 14772545  # from enum XlRgbColor
    rgbSalmon = 7504122  # from enum XlRgbColor
    rgbSandyBrown = 6333684  # from enum XlRgbColor
    rgbSeaGreen = 5737262  # from enum XlRgbColor
    rgbSeashell = 15660543  # from enum XlRgbColor
    rgbSienna = 2970272  # from enum XlRgbColor
    rgbSilver = 12632256  # from enum XlRgbColor
    rgbSkyBlue = 15453831  # from enum XlRgbColor
    rgbSlateBlue = 13458026  # from enum XlRgbColor
    rgbSlateGray = 9470064  # from enum XlRgbColor
    rgbSlateGrey = 9470064  # from enum XlRgbColor
    rgbSnow = 16448255  # from enum XlRgbColor
    rgbSpringGreen = 8388352  # from enum XlRgbColor
    rgbSteelBlue = 11829830  # from enum XlRgbColor
    rgbTan = 9221330  # from enum XlRgbColor
    rgbTeal = 8421376  # from enum XlRgbColor
    rgbThistle = 14204888  # from enum XlRgbColor
    rgbTomato = 4678655  # from enum XlRgbColor
    rgbTurquoise = 13688896  # from enum XlRgbColor
    rgbViolet = 15631086  # from enum XlRgbColor
    rgbWheat = 11788021  # from enum XlRgbColor
    rgbWhite = 16777215  # from enum XlRgbColor
    rgbWhiteSmoke = 16119285  # from enum XlRgbColor
    rgbYellow = 65535  # from enum XlRgbColor
    rgbYellowGreen = 3329434  # from enum XlRgbColor


class BorderWeight:
    xlHairline = 1  # from enum XlBorderWeight
    xlMedium = -4138  # from enum XlBorderWeight
    xlThick = 4  # from enum XlBorderWeight
    xlThin = 2  # from enum XlBorderWeight


class BordersIndex:
    xlDiagonalDown = 5  # from enum XlBordersIndex
    xlDiagonalUp = 6  # from enum XlBordersIndex
    xlEdgeBottom = 9  # from enum XlBordersIndex
    xlEdgeLeft = 7  # from enum XlBordersIndex
    xlEdgeRight = 10  # from enum XlBordersIndex
    xlEdgeTop = 8  # from enum XlBordersIndex
    xlInsideHorizontal = 12  # from enum XlBordersIndex
    xlInsideVertical = 11  # from enum XlBordersIndex


# 专为DDR书写的生成表格函数
def LoadDDRTable():
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)
    topology_table = active_sheet.range(start_ind).current_region.value

    topology_width_table = topology_table[7][3::]

    topology_width_table = [x for x in topology_width_table if x not in ['NA', 'N']]

    if topology_width_table[-1] in [4.3, 6.75]:
        # MS
        active_sheet.range(start_ind[0] + 3, start_ind[1] + 2 + len(topology_width_table)).value = 'MS'
        active_sheet.range(start_ind[0] + 9, start_ind[1] + 3 + len(topology_width_table)).value = 800
        active_sheet.range(start_ind[0] + 10, start_ind[1] + 3 + len(topology_width_table)).value = 3200
    elif topology_width_table[-1] in [3.5, 5.7, 4, 6]:
        # SL
        active_sheet.range(start_ind[0] + 3, start_ind[1] + 2 + len(topology_width_table)).value = 'SL'
        active_sheet.range(start_ind[0] + 9, start_ind[1] + 3 + len(topology_width_table)).value = 800
        active_sheet.range(start_ind[0] + 10, start_ind[1] + 3 + len(topology_width_table)).value = 4700
    else:
        pass


def power_trace():
    global Net_brd_data, non_signal_net_list
    allegro_report_path, layer_type_dict, start_sch_name_list, progress_ind, All_Layer_List = GetSetting()
    SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data, npr_brd_data = read_allegro_data(
        allegro_report_path)
    _, _, _, non_signal_net_list = GetNetList()
    # Net_brd_data = read_allegro_data(allegro_report_path)
    wb = Book(xlsm_path).caller()
    net_data = Net_brd_data.GetData()
    # # print(net_data)

    netlist_sheet = wb.sheets['NetList']
    power_sheet = wb.sheets['power']
    power_data, net_width, list_temp = [], [], []

    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'Non-Signal Net List':
            non_signal_net_list = netlist_sheet.range((cell.Row + 1, cell.Column)).options(expand='table', ndim=1).value
            for i in non_signal_net_list:
                power_data.append([i])
            # # print(power_data)
    for cell in power_sheet.api.UsedRange.Cells:
        if cell.Value == 'Non-Signal Net List':
            power_sheet.range((cell.Row + 1, cell.Column)).value = power_data

    for power in power_data:
        for net in net_data:
            if net[0] == power[0]:
                net_width.append([net[0], net[5]])

    for width in net_width:
        if width not in list_temp:
            list_temp.append(width)

    for data in list_temp:
        if float(data[1]) > 10:
            data.append('Pass')
        else:
            data.append('Fail')
    # # print(list_temp)
    for cell in power_sheet.api.UsedRange.Cells:
        if cell.Value == 'Non-Signal Net List':
            power_sheet.range((cell.Row + 1, cell.Column)).options(expand='table', ndim=3).value = list_temp

    power_sheet.api.Cells.Font.Name = 'Times New Roman'
    power_sheet.api.Cells.Font.Size = 12
    start_ind = (8, 2)
    SetCellFont_current_region(power_sheet, start_ind, 'Times New Roman', 12, 'l')
    SetCellBorder_current_region(power_sheet, start_ind)
    power_sheet.autofit('c')
    for cell in power_sheet.api.UsedRange.Cells:
        if cell.Value == 'result':
            result_range = power_sheet.range((cell.Row + 1, cell.Column))
            for i in xrange(len(list_temp)):
                if power_sheet.range((cell.Row + i + 1, cell.Column)).value == 'Pass':
                    power_sheet.range((cell.Row + i + 1, cell.Column)).api.Interior.Color = RgbColor.rgbLightGreen
                else:
                    power_sheet.range((cell.Row + i + 1, cell.Column)).api.Interior.Color = RgbColor.rgbRed


def power_clear():
    wb = Book(xlsm_path).caller()
    netlist_sheet = wb.sheets['power']

    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'Non-Signal Net List':
            non_sig_idx = (cell.Row + 1, cell.Column)

    netlist_sheet.range(non_sig_idx).expand('table').clear()


def isfloat(string):
    try:
        float(string)
        return True
    except:
        return False


def getpinnumio():
    global pin_number_in_out_dict, mapped_model_name_list
    pin_number_in_out_dict = dict()
    mapped_model_name_list = list()

    # General Pin in-out mapping: 2-pin component, MOS, BJT
    pin_number_in_out_dict[('1', ('1', '2'))] = ['2']  # RES or CAP
    pin_number_in_out_dict[('2', ('1', '2'))] = ['1']  # RES or CAP

    # Update "pin_number_in_out_dict" from the "pin_mapping_library.csv"
    # 应该能改进代码
    library_path = os.path.join(os.path.dirname(xlsm_path), 'dist', 'pin_mapping_library.csv')
    data = open(library_path, 'r').read().split('\n')[1::]
    # # print(data)
    data = [x for x in data if x != '']
    data = list(map(lambda x: x.split(','), data))
    # # print (data)

    for line in data:
        key = (line[0], line[1])
        value = line[2]
        pin_number_in_out_dict[key] = value.split(';')
        mapped_model_name_list.append(line[0])
    # # print(pin_number_in_out_dict)
    mapped_model_name_list = list(set(mapped_model_name_list))
    # # print(mapped_model_name_list)


# 计算两点距离
def two_point_distance(xy1, xy2):
    x1, y1 = xy1[0], xy1[1]
    x2, y2 = xy2[0], xy2[1]

    d = ((float(x1) - float(x2)) ** 2 + (float(y1) - float(y2)) ** 2) ** 0.5
    return d


# 处理总报告的数据
def read_allegro_data(allegro_data_path):
    # Extract data from design file
    # mycrypt = MyCrypt(b'\x17\xBB\x50\xEA\x20\xA7\x4D\xE5\x2F\x7F\x29\x4C\x96\x7D\xE5\xA5')
    start_time = time.clock()
    # # print(1)
    rpt_content = open(allegro_data_path, 'r').read()
    # # print(2, rpt_content)
    content = rpt_content
    # content = mycrypt.mydecrypt(rpt_content)
    # f = open('C:\Users\Tommy\Desktop\password.txt', 'w')
    # f.write(content)
    # f.close()
    # # print(3, rpt_content)
    # 名称可能不同，需修改  ----Gorgeous
    # 17.2 Detailed Trace Length by Layer and Width Report
    # 16.6 Detailed Etch Length by Layer and Width Report
    content = re.split('Detailed Etch Length by Layer and Width Report\n|Detailed Trace Length by Layer and Width '
                       'Report\n|Symbol Pin Report\n|Allegro Report\n|Cross Section Report\n|'
                       'Properties on Nets Report\n', content)
    content = [x.split('\n') for x in content if x != '']
    # # print('content', content)

    # 为何数据不从分报告中提取，而要汇总到一个总报告，是因为Python IO操作占用内存大吗 ----Gorgeous
    SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data, npr_brd_data \
        = None, None, None, None, None
    for item in content:
        # # print(item[3])
        if item[3].find('REFDES,PIN_NUMBER,SYM_NAME,COMP_DEVICE_TYPE,PAD_STACK_NAME,PIN_X,PIN_Y,NET_NAME') > -1:
            SCH_brd_data = brd_data([x.upper().split(',') for x in item[4:-1]])
        elif item[3].find(
                'Net Name,Total Net Length (mils),Layer Name,Total Layer Length (mils),Layer Length % of Total,Line Width (mils),Contiguous Length at Width (mils),Contiguous Length % Layer Length,Contiguous Length End Points') > -1:
            # DEL_content = [x.upper().split(',') for x in item[4:-1]]
            DEL_content = list(map(lambda x: x.upper().split(','), item[4:-1]))
            for ind1 in xrange(len(DEL_content)):
                xy_point = DEL_content[ind1][-1].split(' ')
                # # print(xy_point)
                DEL_content[ind1][-1] = [tuple(xy_point[0:2]), tuple(xy_point[2::])]
            Net_brd_data = brd_data(DEL_content)
        elif item[3].find(
                'Diffpair (Nets),Nominal Gap (mils),Actual Gap (mils),Gap Deviation (mils),Segment Length (mils),Segment End Points') > -1:
            diff_pair_brd_data = brd_data([x.upper().split(',') for x in item[4:-1]])

        # 不知为何要分三种情况
        elif item[4].find(
                'Subclass Name,Type,Material,Thickness (MIL),Conductivity (mho/cm),Dielectric Constant,Loss Tangent,Negative Artwork,Shield,Width (MIL),Unused Pin Pad Suppression,Unused Via Pad Suppression') > -1:
            # Stackup_content = [x.upper() for x in item[5:-1]]
            Stackup_content = list(map(lambda x: x.upper(), item[5:-1]))
            stackup_brd_data = brd_data(Stackup_content)
        elif item[4].find(
                'Subclass Name,Type,Material,Thickness (MIL),Tol +,Tol -,Conductivity (mho/cm),Dielectric Constant,Loss Tangent,Negative Artwork,Shield,Width (MIL),Unused Pin Pad Suppression,Unused Via Pad Suppression') > -1:
            # Stackup_content = [','.join(x[0:4]+x[6::]) for x in map(lambda y: y.upper().split(','), item[5:-1])]
            Stackup_content = list(
                map(lambda x: ','.join(x[0:4] + x[6::]), map(lambda y: y.upper().split(','), item[5:-1])))
            stackup_brd_data = brd_data(Stackup_content)
        elif item[4].find(
                'Subclass Name,Type,Material,Thickness (MIL),Conductivity (mho/cm),Dielectric Constant,Loss Tangent,Negative Artwork,Shield,Width (MIL),Single Impedance (ohm),Unused Pin Pad Suppression,Unused Via Pad Suppression') > -1:
            Stackup_content = [','.join(x[0:9] + x[10::]) for x in map(lambda y: y.upper().split(','), item[5:-1])]
            stackup_brd_data = brd_data(Stackup_content)
        elif item[3].find(
                'NET_NAME,NET_BUS_NAME,NET_SPACING_TYPE,NET_PHYSICAL_TYPE,NET_ELECTRICAL_CONSTRAINT_SET,NET_ECL,NET_DRIVER_TERM_VAL,NET_LOAD_TERM_VAL,NET_FIXED,NET_RATSNEST_SCHEDULE,NET_NO_RAT,NET_NO_GLOSS,NET_NO_PIN_ESCAPE,NET_NO_RIPUP,NET_NO_ROUTE,NET_NO_TEST,NET_PROBE_NUMBER,NET_ROUTE_PRIORITY,NET_ROUTE_TO_SHAPE,NET_SAME_NET,NET_MIN_BVIA_GAP,NET_MIN_BVIA_STAGGER,NET_MAX_BVIA_STAGGER,NET_MIN_LINE_WIDTH,NET_MIN_NECK_WIDTH,NET_DIFFERENTIAL_PAIR,NET_DIFFP_COUPLED_MINUS,NET_DIFFP_COUPLED_PLUS,NET_DIFFP_GATHER_CONTROL,NET_DIFFP_MIN_SPACE,NET_DIFFP_NECK_GAP,NET_DIFFP_PHASE_CONTROL,NET_DIFFP_PHASE_TOL,NET_DIFFP_PRIMARY_GAP,NET_DIFFP_UNCOUPLED_LENGTH,NET_TS_ALLOWED,NET_STUB_LENGTH,NET_IMPEDANCE_RULE,NET_PROPAGATION_DELAY,NET_RELATIVE_PROPAGATION_DELAY,NET_MAX_PARALLEL,NET_MAX_XTALK,NET_MAX_PEAK_XTALK,NET_MAX_INTER_XTALK,NET_MAX_INTRA_XTALK,NET_MAX_EXPOSED_LENGTH,NET_TOTAL_ETCH_LENGTH,NET_MIN_NOISE_MARGIN,NET_MAX_OVERSHOOT,NET_MIN_FIRST_SWITCH,NET_MAX_FINAL_SETTLE,NET_MIN_HOLD,NET_MIN_SETUP,NET_MAX_SSN,NET_MAX_SLEW_RATE,NET_PULSE_PARAM,NET_FIRST_INCIDENT,NET_EDGE_SENS,NET_CLK_2OUT_MAX,NET_CLK_2OUT_MIN,NET_CLK_SKEW_MAX,NET_CLK_SKEW_MIN,NET_CLOCK_NET,NET_TIMING_DELAY_OVERRIDE,NET_ASSIGN_TOPOLOGY,NET_TOPOLOGY_TEMPLATE,NET_TOPOLOGY_TEMPLATE_MAPPING_MODE,NET_TOPOLOGY_TEMPLATE_REVISION,NET_VOLTAGE,NET_SUBNET_NAME,NET_TESTER_GUARDBAND,NET_SHORT,NET_SHORTING_SCHEME,NET_WEIGHT,LAYERSET_GROUP') > -1:
            # # print(item)
            # # print('')
            npr_brd_data = [x.split(',') for x in item[5:-1]]
    # # print('npr_brd_data', npr_brd_data)
    # # print(stackup_brd_data.GetData())
    # # print(diff_pair_brd_data.GetData())
    # # print(1,SCH_brd_data)
    # # print(2,Net_brd_data)
    # # print(3,diff_pair_brd_data)
    # # print(4,stackup_brd_data)

    end_time = time.clock()

    # print('read_allegro_data', end_time - start_time)
    if SCH_brd_data and Net_brd_data and diff_pair_brd_data and stackup_brd_data and npr_brd_data:
        return SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data, npr_brd_data


# 获得差分信号键值对
def diff_detect(diff_pair_brd_data, npr_brd_data):
    # Get diff. pair list
    rpt_content = diff_pair_brd_data.GetData()
    both_diff = []

    for i in xrange(len(rpt_content)):
        if len(rpt_content[i]) > 1:
            both_diff.append(rpt_content[i])
    # 单向键值对
    diff_pair_dict = dict(both_diff)
    # 双向键值对
    diff_pair_dict.update(dict([(x[1], x[0]) for x in both_diff]))
    # # print(diff_pair_dict)
    npr_diff_pair_net_list = []
    part_diff_net_list = list(diff_pair_dict.keys())
    for items in npr_brd_data:
        if items[-1] and items[0] not in part_diff_net_list:
            npr_diff_pair_net_list.append(items[0])

    return diff_pair_dict, npr_diff_pair_net_list


# 获得电源线和地线的名称列表
def get_exclude_netlist(netlist, npr_brd_data):  # netlist = All_Net_List
    # Get pwr and gnd net list
    PWR_Net_List_pro = []
    # 从properties on Nets Report中获取电源线
    for item in npr_brd_data:
        # # print(item)
        if item[1] == 'PWR':
            PWR_Net_List_pro.append(item[0])

    PWR_Net_List_pro = list(set(PWR_Net_List_pro))

    PWR_Net_KeyWord_List = ['^\+.*', '^-.*',
                            'VREF|VPP|VSS|PWR|VREG|VCORE|VCC|VT|VDD|VLED|PWM|VDIMM|VGT|VIN|[^S](VID)|VR',
                            'VOUT|VGG|VGPS|VNN|VOL|VSD|VSYS|VCM|VSA', '.*V[0-9]A.*', '.*V[0-9]\.[0-9]A.*',
                            '.*V[0-9]_[0-9]A.*', '.*V[0-9]S.*', '^V[0-9].*', '.*_V[0-9]', '.*_V[0-9][0-9]',
                            '.*V[0-9]P.*', '.*V[0-9]V.*', '.*[0-9]V[0-9].*', '^[0-9]V.*', '^[0-9][0-9]V.*',
                            '.*[0-9]\.[0-9]V.*', '.*[0-9]_[0-9]V.*', '.*_[0-9]V.*', '.*_[0-9][0-9]V.*',
                            '.*_[0-9]\.[0-9]V.*', '.*[0-9]P[0-9]V.*', '.[0-9]*P[0-9][0-9]V.*', '.*V_[0-9]P[0-9].*',
                            '.*\+[0-9]V.*', '.*\+[0-9][0-9]V.*', '.*?\d+V_S\d$', '\dV']
    PWR_Net_List = [net for net in netlist for keyword in PWR_Net_KeyWord_List if re.findall(keyword, net) != []] \
                   + PWR_Net_List_pro
    # # print(PWR_Net_List)
    PWR_Net_List = sorted(list(set(PWR_Net_List)))

    GND_Net_List = [net for net in netlist if net.find('GND') > -1]
    GND_Net_List = sorted(list(set(GND_Net_List)))

    # 被排除的线：地线和电源线
    Exclude_Net_List = sorted(list(set(PWR_Net_List + GND_Net_List)))

    return Exclude_Net_List, PWR_Net_List, GND_Net_List


# 取 SCH_brd_data 中每个元素最后一个值，线名
def getallnetlist(SCH_brd_data):
    SCH_Data = SCH_brd_data.GetData()
    All_Net_List = list(set([x[-1] for x in SCH_Data if x[-1] not in ['', None]]))
    return All_Net_List


# 对SCH_brd_data信号报告做一个处理，返回一个对象列表
def SCH_detect(SCH_brd_data):
    # Extract schematic info. from report file
    start_time = time.clock()
    global SCH_Name_list, non_signal_net_list, All_Net_List

    All_Net_List = getallnetlist(SCH_brd_data)
    SCH_object_list = list()
    SCH_Data = SCH_brd_data.GetData()

    SCH_Name_list = list()
    SCH_dict_tmp = dict()
    # # print(SCH_Data)
    for line in SCH_Data:
        if line[0]:
            SCH_Name_list.append(line[0])
            # 出现多次情况多次输入
            # get(key, default = None):如果键不存在，返回默认值
            if not SCH_dict_tmp.get(line[0], None):
                # # print(line[0])
                # # print(line[0], line[1], line[-1], line[3], line[-3], line[-2])
                SCH_dict_tmp[line[0]] = [(line[1], line[-1], line[3], line[-3], line[-2])]
            else:
                SCH_dict_tmp[line[0]].append((line[1], line[-1], line[3], line[-3], line[-2]))

    SCH_Name_list = set(SCH_Name_list)
    # # print(SCH_Name_list)
    # # print(SCH_dict_tmp)
    for x in SCH_Name_list:
        pin_list = []
        net_list = []
        model_description = []
        x_point_list = []
        y_point_list = []
        for xttt in SCH_dict_tmp[x]:
            pin_list.append(xttt[0])
            # 可能为空，表示pin脚无连线
            net_list.append(xttt[1])
            # 可能为空，描述为空（不确定）
            model_description.append(xttt[2])
            x_point_list.append(xttt[3])
            y_point_list.append(xttt[4])

        pin_net_dict = dict()
        pin_x_dict = dict()
        pin_y_dict = dict()
        for idx_ppp in xrange(len(pin_list)):
            pin_net_dict[pin_list[idx_ppp]] = net_list[idx_ppp]
            pin_x_dict[pin_list[idx_ppp]] = x_point_list[idx_ppp]
            pin_y_dict[pin_list[idx_ppp]] = y_point_list[idx_ppp]

        # 留下 diff 与 se 信号
        # In order to check connected SCH for any SCH, any net in "Exclude_Net_List" should be excluded
        # check_SCH_net_list也包含空值
        check_SCH_net_list = set(net_list) - set(non_signal_net_list)
        # # print(check_SCH_net_list)

        # connected_SCH_list does not include the case that 0ohm resistor connected btw two SCH
        connected_SCH_list = []
        for d in SCH_Data:
            if d[-1] in check_SCH_net_list:  # and d[0] not in [x, '', None]:
                connected_SCH_list.append(d[0])
        # # print(2)
        # # print(connected_SCH_list)

        # x: 所有元件名  model_description: 描述名   pin_list: 元件个数列表   net_list: 线名列表
        # pin_x_dict: 线的x轴坐标值   pin_y_dict: 线的y轴坐标值  connected_SCH_list: 去除其他信号后的线名（只有diff与se），可能为空
        SCH_object_list.append(comp_device(x, model_description[0], \
                                           pin_list, net_list, pin_net_dict, pin_x_dict, pin_y_dict,
                                           connected_SCH_list))
    # # print (pin_x_dict, pin_y_dict)
    # # print(x, model_description[0],pin_list, net_list, pin_net_dict, pin_x_dict, pin_y_dict,connected_SCH_list)
    #     # print(net_list)
    end_time = time.clock()

    # print('SCH_detect', end_time - start_time)
    return SCH_object_list


# 分别 type_list 和 layer_list
def getalllayerlist(stackup_brd_data):
    # Get Layer List from report file

    data = stackup_brd_data.GetData()
    data = data[0:-1]
    for idx in xrange(len(data)):
        data[idx] = data[idx].split(',')
    layer_list = list()
    type_list = list()

    for layer in data:
        try:
            if layer[1] in ['CONDUCTOR', 'PLANE']:
                layer_list.append(layer[0])
                type_list.append(layer[1])
        except:
            pass

    return layer_list, type_list


# 对Net_brd_data报告做处理，与SCH_object_list信号对象报告中的pin脚坐标值比较0
# 即找出信号线与pin脚相连的坐标点
def net_detect(Net_brd_data, SCH_object_list):
    # Pre-processing the report file to get the etch_line list
    start_time = time.clock()
    global non_signal_net_list, All_Net_List

    net_data = Net_brd_data.GetData()
    net_object_list = list()

    # 去除非信号后的线名
    # 这个check_net_list与SCH_detect中的check_net_list不同
    # 这个check_net_list是完整的所有的线名，而SCH_detect中的check_net_list只是所有与pin脚相接的线名
    # 并且这里的线名不会为空
    check_net_list = list(set([x[0] for x in net_data]) - set(non_signal_net_list))

    for net in check_net_list:
        # # print (net)
        # 按 check_net_list 中的 net 名顺序列出每个 net 对应的叠层，长宽及过孔的坐标值
        layer_list, width_list, length_list, xy1_list, xy2_list = zip(
            *[(x[2], float(x[5]), float(x[6]), x[-1][0], x[-1][1]) for x in net_data if net == x[0]])

        # 将坐标值取一位小数
        # # print('xy1_list', xy1_list)
        xy1_list = [[y[:-1] for y in x] for x in xy1_list]
        xy2_list = [[y[:-1] for y in x] for x in xy2_list]
        # # print('xy2_list', xy2_list)

        segment_list = xrange(1, len(layer_list) + 1)
        # # print (net)
        # # print(layer_list, width_list, length_list, xy1_list, xy2_list)
        seg_width_dict = dict()
        seg_length_dict = dict()
        seg_layer_dict = dict()
        seg_xy_dict1 = dict()
        seg_xy_dict2 = dict()
        # for id_ss in xrange(len(segment_list)):
        for id_ss in xrange(len(layer_list)):
            # # print(segment_list[id_ss])
            seg_width_dict[segment_list[id_ss]] = width_list[id_ss]
            seg_length_dict[segment_list[id_ss]] = length_list[id_ss]
            seg_layer_dict[segment_list[id_ss]] = layer_list[id_ss]
            seg_xy_dict1[segment_list[id_ss]] = xy1_list[id_ss]
            seg_xy_dict2[segment_list[id_ss]] = xy2_list[id_ss]

        # Get the connected SCH and Pin (x, y) information to help topology construction
        connected_SCH_Pin_dict_temp = dict()
        for sch in SCH_object_list:
            if net in sch.GetNetList():
                # # print(sch.GetNetList())
                for pin in sch.GetPinList():
                    # # print (pin)
                    # 按 check_net_list 中的 net 名顺序排序
                    if sch.GetNet(pin) == net:
                        # if net == 'PCH_HSOP0':
                        #     # print(sch.GetName())
                        # # print (pin + '\n')
                        connected_SCH_Pin_dict_temp[(sch.GetName(), pin)] = (
                            (sch.GetXPoint(pin), sch.GetYPoint(pin)), '')

        # if net == 'PCH_HSOP0':
        #     # print(111, connected_SCH_Pin_dict_temp)

        connected_SCH_Pin_dict = dict()

        # # print(net)
        # # print(connected_SCH_Pin_dict_temp)
        # Determine which segment has SCH connected

        # 存在的意义是什么，不懂
        min_pin_net_connect_distance = 20  # unit = mils

        for (sch, pin), (pin_xy_point, sch_layer) in connected_SCH_Pin_dict_temp.items():

            connected_SCH_Pin_dict[(sch, pin)] = list()

            d1_list = []
            d2_list = []
            for seg in segment_list:
                # SCH_data 中点的坐标值 与 net_data 坐标值的距离
                # # print (net, seg)
                # # print (pin_xy_point)
                # # print (seg_xy_dict1[seg])
                # # print (seg_xy_dict2[seg])
                # # print ('\n')

                d1_list.append(two_point_distance(pin_xy_point, seg_xy_dict1[seg]))
                d2_list.append(two_point_distance(pin_xy_point, seg_xy_dict2[seg]))

            # # print (d1_list)
            d_min = min(d1_list + d2_list + [min_pin_net_connect_distance])
            # # print(sch,pin)
            # # print(d_min)
            # # print (d_min) # 大部分为0，小部分距离在10以内，大于十的去除
            # d_min为0表示线与pin脚相连接
            ind = -1
            # 拿到每条信号线上pin的坐标点
            for idx_d12 in xrange(len(d1_list)):
                ind += 1
                # 找出与pin脚相接的线
                if d1_list[idx_d12] == d_min:
                    connected_SCH_Pin_dict[(sch, pin)] += [[pin_xy_point, (ind + 1, 1)]]

                    # if net == 'PCH_HSOP0':
                    #     # print('01', connected_SCH_Pin_dict)
                elif d2_list[idx_d12] == d_min:
                    connected_SCH_Pin_dict[(sch, pin)] += [[pin_xy_point, (ind + 1, 2)]]
                    # if net == 'PCH_HSOP0':
                    #     # print('02', connected_SCH_Pin_dict)

            if connected_SCH_Pin_dict[(sch, pin)] == []:
                connected_SCH_Pin_dict[(sch, pin)] += [[pin_xy_point, None]]
            # # print(connected_SCH_Pin_dict)

        # 信号名，每种信号名的range，宽度字典，长度字典，叠层字典，不间断线信号的起始坐标值，经过所有pin脚的坐标值
        net_object_list.append(
            etch_line(net, segment_list, seg_width_dict, seg_length_dict, seg_layer_dict, seg_xy_dict1, seg_xy_dict2,
                      connected_SCH_Pin_dict))
        # # print(connected_SCH_Pin_dict)
        # if net == 'DP1_AUX_CPU_P':
        #     # print(99999, connected_SCH_Pin_dict)
        # # print(net, segment_list, seg_width_dict, seg_length_dict, seg_layer_dict, seg_xy_dict1, seg_xy_dict2, connected_SCH_Pin_dict)
    end_time = time.clock()
    # print('net_detect', end_time - start_time)
    # # print(net_object_list)
    return net_object_list


# 判断信号名是否在net_data中，并通过net名返回net_object
def get_net_object_by_name(net_name):
    # Get the net object from the net name
    global net_object_list
    find = False
    for net in net_object_list:
        if net.GetName() == net_name:
            return net
    if ~find:
        # # print(net.GetName())
        # # print(net_name)
        return None


def get_SCH_object_by_name(SCH_name):
    # Get the sch_object from the component name
    global SCH_object_list
    for sch in SCH_object_list:
        if sch.GetName() == SCH_name:
            return sch


# 找出与给定芯片连接的信号线的名称
def get_connected_net_list_by_SCH_name(SCH_name):  # Nets in Exclude_Net_List are excluded
    # Get the connected net list of specified component by name
    global SCH_object_list, non_signal_net_list
    # SCH_object_list包含芯片名称，使用的pin脚个数和名称，pin脚连接的线名与坐标值
    for sch in SCH_object_list:
        # # print(sch.GetName(), SCH_name)
        if sch.GetName() == SCH_name:
            check_net_list = sch.GetNetList()
            break

    check_net_list_output = []
    for x in check_net_list:
        # if x not in non_signal_net_list and get_net_object_by_name(x) == None:
        # # print(x)
        # 有必要用get_net_object_by_name(x)做判断吗，通过sch取得的netlist是否一定在net_object_list中
        # 事实证明有必要，PP_HV 与 LX 都不在net_object_list中，因为他们是长度为0的面，不能算作真正的信号线所以要排除
        # net_object_list中包含信号名，同名信号线的个数，叠层，坐标
        if x not in non_signal_net_list and len(x) > 0 and get_net_object_by_name(x):
            check_net_list_output.append(x)

    return check_net_list_output


# 列出各段的芯片pin脚，叠层，宽度，长度
# 例如 ['USB3_CMC_TXDN1', '[FRONT_USB_HEADER-5]:BOTTOM:5.1', 976.98, 'TOP:5.1:[U4-2]', 65.81, '[U4-2]:TOP:5.1:[U4-9]', 40.0, '[U4-9]:TOP:5.1:[LU2-2]', 81.27]
def topology_format(net_name, topology_seg_ind):
    # Formatting function of exported results
    # print('net_name', net_name)
    # print('topology_seg_ind', topology_seg_ind)
    net = get_net_object_by_name(net_name)
    # if net_name == 'PCH_HSOP0':
    #     print(net.GetName(), net.GetSegmentList(), net.GetWidth(), net.GetLength(), net.GetLayer())
    topology_formatted, topology_formatted_extra, topology_formatted_extra3 = [], [], []
    double_flag = False
    three_flag = False

    for i in xrange(len(topology_seg_ind)):
        layer = net.GetLayer(topology_seg_ind[i][0])
        width = net.GetWidth(topology_seg_ind[i][0])
        length = net.GetLength(topology_seg_ind[i][0])

        # 不懂为什么要将奇偶情况分开讨论
        # 长度大于1才匹配
        if float(length) > 1:
            items = net.GetConnectedSCHListBySegInd(topology_seg_ind[i])
            # print('items', items)
            if i % 2 == 0:
                # print(topology_seg_ind[i])
                # print('items', items)
                # 判断是否是芯片相接
                # 与芯片相接则写成 [sch-pin]:layer:width 的形式
                # print('out', topology_seg_ind[i], net.GetConnectedSCHListBySegInd(topology_seg_ind[i]))
                if items is not None:
                    content = '%s:%s' % (layer, width)
                    content_extra = '%s:%s' % (layer, width)
                    content_extra3 = '%s:%s' % (layer, width)
                    # print('in', topology_seg_ind[i], net.GetConnectedSCHListBySegInd(topology_seg_ind[i]))
                    if len(items) == 3:
                        # print(3, items)
                        three_flag = True
                        # 顺序存入
                        for (sch, pin) in net.GetConnectedSCHListBySegInd(topology_seg_ind[i]):
                            content = '[%s-%s]:' % (sch, pin) + content
                            # content_extra = '[%s-%s]:' % (sch, pin) + content_extra
                        # 反向存入
                        for (sch, pin) in items[::-1]:
                            content_extra = content_extra + ':[%s-%s]' % (sch, pin)

                        for (sch, pin) in items[::2] + [items[1]]:
                            # if net_name == 'SPI_CLK_ROM1':
                            # print(sch, pin)
                            content_extra3 = content_extra3 + ':[%s-%s]' % (sch, pin)
                    elif len(items) == 1:
                        for (sch, pin) in net.GetConnectedSCHListBySegInd(topology_seg_ind[i]):
                            content = '[%s-%s]:' % (sch, pin) + content
                            content_extra = '[%s-%s]:' % (sch, pin) + content_extra
                            content_extra3 = '[%s-%s]:' % (sch, pin) + content_extra3
                    elif len(items) > 1:
                        # print('before_content', content)
                        double_flag = True
                        # 顺序存入
                        for (sch, pin) in net.GetConnectedSCHListBySegInd(topology_seg_ind[i]):
                            # if net_name == 'SPI_CLK_ROM1':
                            #     print(sch, pin)
                            content = '[%s-%s]:' % (sch, pin) + content
                            # content_extra = '[%s-%s]:' % (sch, pin) + content_extra
                        # 反向存入
                        for (sch, pin) in items[::-1]:
                            content_extra = '[%s-%s]:' % (sch, pin) + content_extra
                            content_extra3 = '[%s-%s]:' % (sch, pin) + content_extra3
                        # print('content', content)
                        # print('content_extra', content_extra)
                # 如不与芯片相接则写成 layer:width 的形式
                else:
                    content = '%s:%s' % (layer, width)
                    content_extra = '%s:%s' % (layer, width)
                    content_extra3 = '%s:%s' % (layer, width)
                # print(1, content)
                # print(1, content_extra)
            else:
                # print(2, items)
                # print(2, content)
                # print(2, content_extra)
                # print('items', items)
                if items is not None:
                    if len(items) == 3:
                        # print(items)
                        # items_copy = copy.deepcopy(items)
                        three_flag = True
                        for (sch, pin) in items:
                            content = content + ':[%s-%s]' % (sch, pin)

                        topology_formatted.append(content)
                        topology_formatted.append(length)
                        # print(3, content)

                        for (sch, pin) in items[::-1]:
                            content_extra = content_extra + ':[%s-%s]' % (sch, pin)

                        for (sch, pin) in items[::2] + [items[1]]:
                            print(sch, pin)
                            content_extra3 = content_extra3 + ':[%s-%s]' % (sch, pin)

                        # print(3, content_extra)
                        topology_formatted_extra.append(content_extra)
                        topology_formatted_extra.append(length)
                        topology_formatted_extra3.append(content_extra3)
                        topology_formatted_extra3.append(length)
                    elif len(items) == 1:
                        for (sch, pin) in items:
                            content = content + ':[%s-%s]' % (sch, pin)
                            content_extra = content_extra + ':[%s-%s]' % (sch, pin)
                            content_extra3 = content_extra3 + ':[%s-%s]' % (sch, pin)
                        topology_formatted.append(content)
                        topology_formatted.append(length)
                        topology_formatted_extra.append(content_extra)
                        topology_formatted_extra.append(length)
                        topology_formatted_extra3.append(content_extra3)
                        topology_formatted_extra3.append(length)
                        # print(4, content)
                    # print(4, content_extra)
                    elif len(items) > 1:
                        double_flag = True
                        for (sch, pin) in items:
                            content = content + ':[%s-%s]' % (sch, pin)

                        topology_formatted.append(content)
                        topology_formatted.append(length)
                        # print(3, content)

                        for (sch, pin) in items[::-1]:
                            content_extra = content_extra + ':[%s-%s]' % (sch, pin)
                            content_extra3 = content_extra3 + ':[%s-%s]' % (sch, pin)

                        # print(3, content_extra)
                        topology_formatted_extra.append(content_extra)
                        topology_formatted_extra.append(length)
                        topology_formatted_extra3.append(content_extra3)
                        topology_formatted_extra3.append(length)
                else:

                    topology_formatted.append(content)
                    topology_formatted.append(length)
                    topology_formatted_extra.append(content_extra)
                    topology_formatted_extra.append(length)
                    topology_formatted_extra3.append(content_extra3)
                    topology_formatted_extra3.append(length)
    # print('')
    # print('topology_formatted', topology_formatted)
    # print('topology_formatted_extra', topology_formatted_extra)
    # print('topology_formatted_extra3', topology_formatted_extra3)
    if three_flag:
        # print('')
        # print(3, topology_formatted_extra3)
        # print('')
        # print('double_flag', [[net_name] + topology_formatted, [net_name] + topology_formatted_extra])
        return [[net_name] + topology_formatted, [net_name] + topology_formatted_extra,
                [net_name] + topology_formatted_extra3]
    elif double_flag:
        # print('')
        # print(2, topology_formatted_extra3)
        # print('')
        return [[net_name] + topology_formatted, [net_name] + topology_formatted_extra]
    else:
        # print('single_flag', [[net_name] + topology_formatted])
        return [[net_name] + topology_formatted]


# 返回每个信号线的各段ind
# 其中有错误逻辑
def topology_extract1(net_name, start_sch_name, start_sch_pin=None):
    # Topology Extraction Function
    # start_time = time.clock()
    # 通过net_name返回net_object
    check_net = get_net_object_by_name(net_name)
    # 同一个net_name的线的个数
    seg_number_list = check_net.GetSegmentList()
    # # print(seg_number_list)
    # # print('')
    seg_ind_list_original = list()
    seg_inter_map_dict = dict()
    xy_point_list = list()
    # if net_name == 'HDA_BCLK_R' and start_sch_name == 'U3':
    #     # print(0, seg_number_list)

    # # print(seg_number_list)
    # 每条信号线宽度分割段的数量list,eg:seg = xrange(1,3)
    for seg in seg_number_list:
        # 每条连续信号线都有起始和终止点坐标所以分1,2

        # ind_list
        seg_ind_list_original.append((seg, 1))
        seg_ind_list_original.append((seg, 2))
        # ind_dict
        seg_inter_map_dict[(seg, 1)] = (seg, 2)
        seg_inter_map_dict[(seg, 2)] = (seg, 1)
        # ind_坐标点
        xy_point_list.append(check_net.GetXY1(seg))
        xy_point_list.append(check_net.GetXY2(seg))
    # if net_name == 'HDA_BCLK_R' and start_sch_name == 'U3':
    #     # print(1, seg_ind_list_original)
    # # print(seg_ind_list_original)
    seg_ind_xy_point_dict = dict()
    for iddx_seg in xrange(len(seg_ind_list_original)):
        seg_ind_xy_point_dict[seg_ind_list_original[iddx_seg]] = xy_point_list[iddx_seg]
    # # print(seg_ind_xy_point_dict)
    # if net_name == 'PCH_HSOP0' and start_sch_name == 'JN44':
    #     # print(0, seg_ind_xy_point_dict)

    # if net_name == 'HDA_BCLK_R' and start_sch_name == 'U3':
    #     # print(start_sch_name)
    #     # print (seg_ind_list_original)
    #     # print (seg_inter_map_dict)
    #     # print (xy_point_list)
    #     # print(seg_ind_xy_point_dict)
    # # print(net_name)
    # # print(seg_number_list)
    sch_list = []
    seg_ind_list_original_start = [x for x in seg_ind_list_original]
    # # print(seg_ind_list_original_start)
    for seg in seg_number_list:  # eg: 1 2 3
        # 遍历每条线所连接的器件名
        for sch in check_net.GetConnectedSCHList(seg):
            ######################
            sch_list.append(sch)
            ######################
    sch_name_list = [x[0] for x in sch_list]
    sch_name_list_start = [x for x in sch_name_list]
    # if net_name == 'SUSCLK_M2230' and start_sch_name == 'J38':
    #     # print(sch_name_list_start)
    topology_seg_ind_list_all = list()
    # previous_sch_pin_keys = []
    # previous_sch_pin_key1 = check_net.conn_schpin.get(previous_sch_pin1, None)
    # if previous_sch_pin_key1:
    #     previous_sch_pin_key1 = previous_sch_pin_key1[0][-1]
    #     previous_sch_pin_keys.append(previous_sch_pin_key1)
    #
    # previous_sch_pin_key2 = check_net.conn_schpin.get(previous_sch_pin2, None)
    # if previous_sch_pin_key2:
    #     previous_sch_pin_key2 = previous_sch_pin_key2[0][-1]
    #     previous_sch_pin_keys.append(previous_sch_pin_key2)

    # print('start_sch_name', start_sch_name)
    # print('sch_name_list', sch_name_list)
    # start_seg_ind1_list = [1]
    while start_sch_name in sch_name_list:
        start_seg_ind1_list = list()
        # if net_name == 'HDA_BCLK_R' and start_sch_name == 'U3':
        #     # print(2, sch_name_list)
        sch_name_list = [x[0] for x in sch_list]

        for sch in sch_list:
            # if net_name == 'HDA_BCLK_R' and start_sch_name == 'U3':
            #     # print(sch)
            if start_seg_ind1_list == []:

                # 信号名相同
                # # print('sch', sch)
                # sch = (SCH_name, pin_id, pinpoint, seg_ind)
                # 找到开始的芯片所连的信号线
                # if previous_sch_pin1 and sch[0] == start_sch_name and previous_sch_pin1[0] == start_sch_name:
                #     # print('ok1')
                #     # print(previous_sch_pin1[1], sch[1])
                #     if start_sch_pin != None and previous_sch_pin1[1] == sch[1]:
                #         # print('ok2')
                #         previous_sch_pin_keys.append((sch[0], sch[1]))
                # if previous_sch_pin2 and sch[0] == start_sch_name and previous_sch_pin2[0] == start_sch_name:
                #     # print('ok3')
                #     # print(previous_sch_pin2[1], sch[1])
                #     if start_sch_pin != None and previous_sch_pin2[1] == sch[1]:
                #         # print('ok4')
                #         previous_sch_pin_keys.append((sch[0], sch[1]))
                if sch[0] == start_sch_name:
                    # if net_name == 'USB_N12_R' and start_sch_name == 'U300':
                    #     # print(0, sch)
                    # 第一个判断是自己输入起始pin名，第二个判断是默认全部pin名
                    if start_sch_pin != None and start_sch_pin == sch[1] or start_sch_pin == None:
                        start_seg_ind1_list.append(sch[-1])
                    sch_list.remove(sch)
        if start_seg_ind1_list != [] and seg_ind_list_original == []:
            # # print('in')
            # # print(seg_ind_list_original_start)
            seg_ind_list_original = seg_ind_list_original_start
        start_seg_ind2_list = []
        # # print(net_name)
        # # print(1)
        # # print(start_seg_ind1_list)
        # # print(2)
        for x in start_seg_ind1_list:
            start_seg_ind2_list.append(seg_inter_map_dict[x])
        if start_seg_ind1_list != []:
            for start_idx in xrange(len(start_seg_ind1_list)):
                try:
                    seg_ind_list_original.remove(start_seg_ind1_list[start_idx])
                    seg_ind_list_original.remove(start_seg_ind2_list[start_idx])
                except ValueError:
                    pass
                # # print('B',seg_ind_list_original)
                end_all = False
                i = 0
                while end_all is False:
                    i += 1
                    # # print(i)
                    topology_seg_ind_list = list()
                    # 存入从开始到结束的信号线名称
                    topology_seg_ind_list.append(start_seg_ind1_list[start_idx])
                    topology_seg_ind_list.append(start_seg_ind2_list[start_idx])
                    # 没必要，因为两者无区别
                    seg_ind_list = list(seg_ind_list_original)
                    end = False
                    j = 0
                    while end is False:
                        j += 1
                        # 从开始的线依次向后找寻坐标点重合的信号线代表是信号下一个经过的信号线
                        for seg_ind in seg_ind_list:
                            # seg_ind_list中的每一个与topology_seg_ind_list最后一个相比，找出坐标相同点
                            # 判断线是否完整
                            if seg_ind != topology_seg_ind_list[-1] and seg_ind_xy_point_dict[seg_ind] == \
                                    seg_ind_xy_point_dict[topology_seg_ind_list[-1]]:
                                topology_seg_ind_list.append(seg_ind)
                                topology_seg_ind_list.append(seg_inter_map_dict[seg_ind])
                                seg_ind_list.remove(seg_ind)
                                seg_ind_list.remove(seg_inter_map_dict[seg_ind])
                                break
                        # Check if any connected segment to the topology_seg_ind_list[-1]
                        end = True

                        # 说明线没找完，继续往下寻找
                        for seg_ind in seg_ind_list:
                            if seg_ind != topology_seg_ind_list[-1] and seg_ind_xy_point_dict[seg_ind] == \
                                    seg_ind_xy_point_dict[topology_seg_ind_list[-1]]:
                                end = False
                                break
                        # Find the last split point to the end for this line, and remove it from the seg_ind_list_original
                        split_ind_list = []

                        if end is True:
                            # # print(topology_seg_ind_list)
                            # # print(seg_ind_list)
                            # # print('\n')

                            # 找分叉点
                            for ind1 in xrange(len(topology_seg_ind_list)):
                                for ind2 in xrange(len(seg_ind_list)):
                                    if topology_seg_ind_list[ind1] != seg_ind_list[ind2] and seg_ind_xy_point_dict[ \
                                            topology_seg_ind_list[ind1]] == seg_ind_xy_point_dict[seg_ind_list[ind2]]:
                                        split_ind_list.append(ind1)

                            # Record the split point
                            # 看不懂2代表什么，有什么作用
                            if 0 not in split_ind_list:
                                split_ind_list.append(2)

                            for ind in xrange(len(topology_seg_ind_list)):
                                if ind >= max(split_ind_list):
                                    if topology_seg_ind_list[ind] not in [start_seg_ind1_list[start_idx], \
                                                                          start_seg_ind2_list[start_idx]]:
                                        # if net_name == 'MXM_DPB_AUX_DN_C':
                                        #     # print(start_seg_ind1_list)
                                        #     # print(start_seg_ind2_list)
                                        #     # print(topology_seg_ind_list[ind])
                                        #     # print(split_ind_list)
                                        #     # print(ind)
                                        seg_ind_list_original.remove(topology_seg_ind_list[ind])
                            # if net_name == 'MXM_DPB_AUX_DN_C':
                            #     # print(seg_ind_list_original)
                            if split_ind_list == []:
                                end = False

                        # 为什么为50
                        if i > 20 or j > 20:
                            raise ValueError('%s-%s can\'t be extracted' % (start_sch_name, net_name))

                    topology_seg_ind_list_all.append(topology_seg_ind_list)

                    end_all = True

                    for seg_ind in seg_ind_list_original:
                        if seg_ind != start_seg_ind2_list[start_idx] and seg_ind_xy_point_dict[seg_ind] \
                                == seg_ind_xy_point_dict[start_seg_ind2_list[start_idx]]:
                            end_all = False
                            break
        # if net_name == 'PCH_HSOP0' and start_sch_name == 'JN44':
        #     # print(6, topology_seg_ind_list_all)
        #     # print(7, seg_ind_list_original)
        # if net_name == 'HDA_BCLK_R' and start_sch_name == 'U3':
        #     # print(6, seg_ind_list_original)
        #     # print(7, topology_seg_ind_list_all)
        #     # print(sch_name_list_start)
        #     # print(len(seg_ind_list_original), sch_name_list_start.count(start_sch_name))

    # if sch_name_list_start.count(start_sch_name) == 1 and len(seg_ind_list_original) == 0\
    #         or sch_name_list_start.count(start_sch_name) >= 2 and \
    #         len(seg_ind_list_original) <= sch_name_list_start.count(start_sch_name)\
    #         or start_seg_ind1_list == []:
    # # print('start_sch_name', start_sch_name)
    # # print('previous_sch_pin_keys', previous_sch_pin_keys)
    # if previous_sch_pin_keys:
    #     for i in previous_sch_pin_keys:
    #         topology_seg_ind_list_all = [line for line in topology_seg_ind_list_all
    #                                      if previous_sch_pin_keys not in line]
    if len(seg_ind_list_original) == 0 or start_seg_ind1_list == []:
        # # print(net_name)
        # # print(topology_seg_ind_list_all)
        connected_sch_list = list()
        end = False
        while end is False:
            end = True
            for line in topology_seg_ind_list_all:
                # # print('line', line)
                for idx in xrange(len(line)):
                    if len(line) >= 4:
                        # 去除掉第一个与最后一个pin点坐标
                        # # print(len(line))
                        if 0 < idx <= len(line) - 3:
                            # # print(idx)
                            # Construct the tree line
                            # 返回(sch, pin_id)
                            connected_sch_tmp = check_net.GetConnComp(line[idx])
                            # # print(line[idx])
                            # # print(connected_sch_tmp)
                            # 如果是pin脚坐标点
                            if connected_sch_tmp not in connected_sch_list + [None]:
                                # 去除最后一段线
                                topology_seg_ind_list_all.append(line[0:idx + 1])
                                # # print(line[0:idx+1])
                                # 保存从一个pin脚到另一个pin脚中间不经过芯片的同名信号线
                                connected_sch_list.append(connected_sch_tmp)
                                end = False
                                break
                if not end:
                    break
        # # print(topology_seg_ind_list_all)
        # 所以topology_seg_ind_list_all中就包含从开头到结尾的整段信号线，与每两个芯片之间的信号线
        # if net_name == 'MXM_DPB_AUX_DN_C':
        #     # print(topology_seg_ind_list_all)
        # if net_name == 'PCH_HSOP0' and start_sch_name == 'JN44':
        #     # print('OK', topology_seg_ind_list_all)
        # # print('topology_seg_ind_list_all1', topology_seg_ind_list_all)
        # # print('previous_sch_pin_point', previous_sch_pin_point)
        # topology_seg_ind_list_all = [i for i in topology_seg_ind_list_all if previous_sch_pin_point not in i]
        # # print('topology_seg_ind_list_all2', topology_seg_ind_list_all)
        end_time = time.clock()
        # # print('topo1', end_time - start_time)
        return topology_seg_ind_list_all
    else:
        # 类似这种可更换器件的情况会fail
        # if net_name == 'USB3_TXDN1_C':
        # # print(net_name)
        # # print(start_sch_name)
        # # print(seg_ind_list_original)
        # for ind in seg_ind_list_original:
        #     # print(seg_ind_xy_point_dict[ind])
        raise ValueError('%s-%s can\'t be extracted' % (start_sch_name, net_name))


# 返回与本信号线相连的下一段信号线的名称
def net_mapping(sch_name, input_pin, previous_net=None):
    # Find the input-output mapping of component
    if sch_name is not None:
        sch = get_SCH_object_by_name(sch_name)
        model_name = sch.GetModel()
        # Two pin component: RES, CAP,...
        output_pin_list = pin_number_in_out_dict.get((input_pin, sch.GetPinList()))
        # if previous_sch_pin_list and sch_name == previous_sch_pin_list[0][0]:
        #     output_pin_list = [i for i in output_pin_list if previous_sch_pin_list[0][1] in i]
        # Find output pin by model name
        if output_pin_list is None:
            output_pin_list = pin_number_in_out_dict.get((model_name, input_pin))
            # if previous_sch_pin_list and sch_name == previous_sch_pin_list[0][0]:
            #     output_pin_list = [i for i in output_pin_list if previous_sch_pin_list[0][1] in i]
        if output_pin_list is None:
            return [None], [None]
        else:
            # # print([sch.GetNet(output_pin) for output_pin in output_pin_list], output_pin_list)
            return [sch.GetNet(output_pin) for output_pin in output_pin_list if sch.GetNet(output_pin) != previous_net], \
                   output_pin_list
    else:
        return [None], [None]


# 返回每个信号线的经过芯片，叠层，pin脚名，线长
def topology_extract2(start_net_name, start_sch_name):
    # Topology Extraction Function 2
    # start_time = time.clock()
    # Initialize
    global All_Net_List, All_Layer_List, non_signal_net_list
    topology_list = list()

    try:
        topology_seg_ind = topology_extract1(start_net_name, start_sch_name, start_sch_pin=None)
    # # print('topology_seg_ind', topology_seg_ind)
    # print('topology_seg_ind', topology_seg_ind)
    except:
        # print('')
        # print('topology_extract1_error')
        # print('')
        raise ValueError('%s-%s can\'t be extracted!' % (start_sch_name, start_net_name))

    etch_line_old = get_net_object_by_name(start_net_name)
    # # print('etch_line_old1', etch_line_old.conn_schpin)
    topology_return_list = []
    # 显示一根信号线的数据，并指出与之相连的下一根信号线
    for line in topology_seg_ind:
        sch_list, pin_list, net_list, next_pin_list, previous_pin_list1, previous_pin_list2 = [], [], [], [], [], []

        # 判断最后一个线名是否与芯片相连接,因为经过过孔或者信号线宽度变化也会分段
        # 最后一个线名是否与芯片相连接才能进入循环
        # # print('etch_line', etch_line.conn_schpin)
        # # print('etch_line_old', etch_line_old.conn_schpin)
        # # print(1, line[-1], etch_line_old.GetConnectedSCHListBySegInd(line[-1]))
        if etch_line_old.GetConnectedSCHListBySegInd(line[-1]):
            # # print('ok')
            # Find the connected component of etch line
            # 找出最后一个线名的坐标，其实找出了每条信号线上所经过的所有芯片（pin脚）的坐标值
            if line[-1][1] == 1:
                net_xy_point = etch_line_old.GetXY1(line[-1][0])
            elif line[-1][1] == 2:
                net_xy_point = etch_line_old.GetXY2(line[-1][0])
            # # print('net_xy_point', net_xy_point)
            # 找出与最后条线相连的芯片的坐标
            # 因此net_xy_point与sch_pin_xy_point_list内坐标其实是相同的
            sch_pin_xy_point_list = list()
            for (sch, pin) in etch_line_old.GetConnectedSCHListBySegInd(line[-1]):
                sch_object = get_SCH_object_by_name(sch)
                sch_pin_xy_point_list.append((sch_object.GetXPoint(pin), sch_object.GetYPoint(pin)))
            d_xy_list = []
            # # print('sch_pin_xy_point_list', sch_pin_xy_point_list)
            # 因为两者值相同，所以d_xy_list全为0，考虑是否能省略这部分代码
            for x in sch_pin_xy_point_list:
                d_xy_list.append(two_point_distance(net_xy_point, x))

            ######################################################################################
            # index()方法检测字符创中是否包含字符串str，存在返回索引值，不存在抛出异常
            for ind in xrange(len(etch_line_old.GetConnectedSCHListBySegInd(line[-1]))):
                sch_list.append(etch_line_old.GetConnectedSCHListBySegInd(line[-1])[ind][0])
                pin_list.append(etch_line_old.GetConnectedSCHListBySegInd(line[-1])[ind][1])
            # # print('sch_list_in', sch_list)

        else:
            # 得出每条信号线所经过的芯片的名称以及pin脚名称列表
            sch_list.append(None)
            pin_list.append(None)

        sch_nochange_list = sch_list
        # print('first_sch_list', sch_list)
        # print('first_pin_list', pin_list)
        # previous_pin_list1 = copy.deepcopy(pin_list)
        # previous_pin_list2 = copy.deepcopy(pin_list)
        # previous_sch_pin_list1 = [etch_line_old.GetConnComp(line[-1])]
        # previous_sch_pin_list2 = [etch_line_old.GetConnComp(line[-2])]
        # print('first_previous_pin_list1', previous_pin_list1)
        # print('first_previous_sch_pin_list1', previous_sch_pin_list1)
        for ind in xrange(len(sch_nochange_list)):
            next_net, next_pin = net_mapping(sch_nochange_list[ind], pin_list[ind])
            # print('next_net, next_pin', next_net, next_pin)
            # next_net通常为单个线名，下列代码是否可以改写为 if next_net:
            for idxx_ in xrange(len(next_net)):
                # start_net_name为每段信号线的名称， line为从两段线开始的ind对

                # format_items = topology_format(start_net_name, line)

                topology_list = topology_format(start_net_name, line)

                net_list.append(next_net[idxx_])
                next_pin_list.append(next_pin[idxx_])

                # 如果存在信号线，则存入信号线所经过的芯片名称
                if idxx_ > 0:
                    sch_list.append(sch_list[-1])
                # 没看懂什么意思
                if net_list[-1] in non_signal_net_list:
                    topology_list[-1].append(net_list[-1])

            # print('topology_list', topology_list)
            # print('first_net_list', net_list)
            # print('first_next_pin_list', next_pin_list)
            # print('first_sch_list', sch_list)
            # print('')
            net_list_start = net_list
            end = True
            # Ending Condition Detect
            for net in [x for x in net_list if x not in non_signal_net_list]:
                if net is not None:
                    end = False
            j = -1

            pre_net_list = []
            topology_out_list = [topology_list[0]]  # F7684584

            # Topology Detect for secondary part (net change)
            while end is False:
                # if start_net_name == 'SUSCLK_M2230' and start_sch_name == 'J38':
                #     # print(j, end)
                j += 1
                topology_list_temp, net_list_temp, sch_list_temp, pin_list_temp, next_pin_list_temp = [], [], [], [], []
                previous_sch_pin_temp1, previous_sch_pin_temp2 = [], []

                i = -1
                end = True
                my_flag = 0
                # # print(1111, topology_list)
                # print('')
                # print('begin', net_list)
                # print('begin', sch_list)
                # print('begin', next_pin_list)
                for ij1 in xrange(len(net_list)):
                    i += 1
                    # print(ij1, net_list[ij1], net_list)
                    # 下段代码为重复代码，可以写成函数简化
                    # 只进入信号线名
                    if net_list[ij1] not in non_signal_net_list + [None]:
                        etch_line = get_net_object_by_name(net_list[ij1])

                        pre_net_list.append(net_list[ij1])
                        #
                        # if start_net_name == 'SPI_SCK':
                        #     # print(j,net_list[ij1])
                        #     # print(j,next_pin_list[ij1])
                        #     # print(j,sch_list[ij1])

                        try:
                            # 找出下一条线的段id
                            # print('before', net_list[ij1], sch_list[ij1], next_pin_list[ij1])
                            # if previous_pin_list:
                            topology_seg_ind = topology_extract1(net_list[ij1], sch_list[ij1],
                                                                 start_sch_pin=next_pin_list[ij1])
                            # print('topology_seg_ind', topology_seg_ind)
                            for ind in xrange(len(topology_seg_ind)):
                                # # print(i, ind, topology_seg_ind[ind])
                                # 找出与芯片相接的线
                                if etch_line.GetConnectedSCHListBySegInd(topology_seg_ind[ind][-1]) is not None:
                                    # previous_sch_pin_temp1.append(
                                    #     etch_line.GetConnComp(topology_seg_ind[ind][-1]))
                                    # previous_sch_pin_temp2.append(
                                    #     etch_line.GetConnComp(topology_seg_ind[ind][-2]))
                                    # 存入线的两端坐标值
                                    if topology_seg_ind[ind][-1][1] == 1:
                                        net_xy_point = etch_line.GetXY1(topology_seg_ind[ind][-1][0])
                                    elif topology_seg_ind[ind][-1][1] == 2:
                                        net_xy_point = etch_line.GetXY2(topology_seg_ind[ind][-1][0])
                                    sch_pin_xy_point_list = list()
                                    # if start_net_name == 'PCH_HSOP3':
                                    #     # print(net_xy_point)
                                    for (sch, pin) in etch_line.GetConnectedSCHListBySegInd(topology_seg_ind[ind][-1]):
                                        sch_object = get_SCH_object_by_name(sch)
                                        sch_pin_xy_point_list.append(
                                            (sch_object.GetXPoint(pin), sch_object.GetYPoint(pin)))
                                    # if start_net_name == 'PCH_HSOP3':
                                    #     # print(sch_pin_xy_point_list)
                                    d_xy_list = []
                                    for x in sch_pin_xy_point_list:
                                        d_xy_list.append(two_point_distance(net_xy_point, x))

                                    sch_list_temp.append(
                                        etch_line.GetConnectedSCHListBySegInd(topology_seg_ind[ind][-1])
                                        [d_xy_list.index(max(d_xy_list))][0])
                                    # sch_previous_list.append(
                                    #     etch_line.GetConnectedSCHListBySegInd(topology_seg_ind[ind][-1])[-2][0])
                                    pin_list_temp.append(
                                        etch_line.GetConnectedSCHListBySegInd(topology_seg_ind[ind][-1])
                                        [d_xy_list.index(max(d_xy_list))][1])
                                    # pin_previous_list.append(
                                    #     etch_line.GetConnectedSCHListBySegInd(topology_seg_ind[ind][-1])[-2][1])

                                    # if start_net_name == 'PCH_HSOP3':
                                    #     # print(sch_list_temp)
                                    #     # print(pin_list_temp)
                                else:
                                    sch_list_temp.append(None)
                                    pin_list_temp.append(None)
                                    # previous_sch_pin_temp1.append(None)
                                    # previous_sch_pin_temp2.append(None)
                                # print('net_list[ij1]', net_list[ij1])
                                # print('topology_seg_ind[ind]', topology_seg_ind[ind])
                                format_item_list = topology_format(net_list[ij1], topology_seg_ind[ind])
                                # print('format_item_list', len(format_item_list), format_item_list)
                                format_flag = True if len(format_item_list) > 1 else False
                                previous_net = format_item_list[0][0]

                                next_net, next_pin = net_mapping(sch_list_temp[-1], pin_list_temp[-1],
                                                                 previous_net=previous_net)
                                # print('next_net', next_net)
                                # print('next_pin', next_pin)

                                ########################################################
                                # 自己修改的代码
                                # print('net_list_start', len(net_list_start), net_list_start)
                                start_flag = 1
                                for idxx_ in xrange(len(next_net)):
                                    for start_ind in xrange(len(net_list_start)):
                                        # 如果回到最初的net_list，意思是转了一圈回来，为了防止无限循环，则退出
                                        if next_net[idxx_] == net_list_start[start_ind] \
                                                or next_net[idxx_] == start_net_name:  # -F7684584
                                            # and next_pin[idxx_] == next_pin_list_start[start_ind]\
                                            # and sch_list_temp[idxx_] == sch_list_start[start_ind]:
                                            net_list_temp.append(None)
                                            next_pin_list_temp.append(None)
                                            start_flag = 0
                                        # 如果next_net不在初始net_list中，则保存
                                        if start_ind == len(net_list_start) - 1 and start_flag:
                                            # net_list_temp.append(next_net[idxx_])
                                            # next_pin_list_temp.append(next_pin[idxx_])
                                            # if format_flag:
                                            for x in range(len(format_item_list)):
                                                net_list_temp.append(next_net[idxx_])
                                                next_pin_list_temp.append(next_pin[idxx_])
                                                if format_flag and x > 0:
                                                    sch_list_temp.append(sch_list_temp[-1])
                                                    pin_list_temp.append(pin_list_temp[-1])

                                    ########################################################

                                    # if idxx_ > 0:
                                    #     sch_list_temp.append(sch_list_temp[-1])
                                    # # print(222, topology_list[i])
                                    # # print(333, topology_format(net_list[ij1], topology_seg_ind[ind]))
                                    # format_item_list = []
                                    # # print('format_item_before_list', format_item_before_list)
                                    # for x in format_item_before_list:
                                    #     if previous_net in x:
                                    #         format_item_list.append(x[1:])
                                    #     else:
                                    #         format_item_list.append(x)
                                    # # print('format_item_list', format_item_list)
                                    # # print('topology_list', topology_list)
                                    if format_item_list:
                                        if len(format_item_list) > 1:
                                            for format_item in format_item_list:
                                                # if next_net[0] is None:
                                                #     topology_list_temp.append(
                                                #         topology_list[i] + format_item[1:])
                                                #     topology_out_list.append(
                                                #         topology_list[i] + format_item[1:])
                                                # # print('topology_out_list1', topology_list[i] + format_item[1:])
                                                # else:
                                                topology_list_temp.append(
                                                    topology_list[i] + format_item)
                                                topology_out_list.append(
                                                    topology_list[i] + format_item)
                                                # print('topology_out_list1', topology_list[i] + format_item)
                                        else:
                                            # if next_net[0] is None:
                                            #     topology_list_temp.append(
                                            #         topology_list[i] + format_item_list[0][1:])
                                            #     topology_out_list.append(
                                            #         topology_list[i] + format_item_list[0][1:])
                                            #     # print('topology_out_list2', topology_list[i] + format_item_list[0][1:])
                                            # else:
                                            topology_list_temp.append(
                                                topology_list[i] + format_item_list[0])
                                            topology_out_list.append(
                                                topology_list[i] + format_item_list[0])

                                # print('in_net_list_temp', net_list_temp)
                                # print('in_sch_list_temp', sch_list_temp)

                        except:
                            topology_list_temp.append(topology_list[i] + ['Exception;%s' % net_list[ij1]])
                            topology_out_list.append(topology_list[i] + ['Exception;%s' % net_list[ij1]])
                            sch_list_temp.append(None)
                            # pin_list_temp.append(None)
                            previous_sch_pin_temp1.append(None)
                            previous_sch_pin_temp2.append(None)
                            net_list_temp.append(None)
                            next_pin_list_temp.append(None)

                    elif net_list[ij1] in non_signal_net_list and j != 0:
                        topology_list_temp.append(topology_list[i] + [net_list[ij1]])
                        topology_out_list.append(topology_list[i] + [net_list[ij1]])
                        # # print('topology_out_list3', topology_list[i] + [net_list[ij1]])
                        sch_list_temp.append(None)
                        # pin_list_temp.append(None)
                        previous_sch_pin_temp1.append(None)
                        previous_sch_pin_temp2.append(None)
                        net_list_temp.append(None)
                        next_pin_list_temp.append(None)
                    else:
                        topology_list_temp.append(topology_list[i])
                        topology_out_list.append(topology_list[i])
                        # # print('topology_out_list4', topology_list[i])
                        # if start_net_name == 'GPP_CLK1N_LAN':
                        #     # print(topology_list[j])
                        sch_list_temp.append(None)
                        previous_sch_pin_temp1.append(None)
                        previous_sch_pin_temp2.append(None)
                        # pin_list_temp.append(None)
                        net_list_temp.append(None)
                        next_pin_list_temp.append(None)

                    # if start_net_name == 'DPC_AUX_DP_C':
                    #     # print(topology_list)
                    # print('topology_list_temp', topology_list_temp)
                    # print('net_list_temp', net_list_temp)
                    # print('sch_list_temp', sch_list_temp)
                    # print('')
                topology_list = list(topology_list_temp)

                # print('topology_list_final', topology_list)
                # if start_net_name == 'DPC_AUX_DP_C':
                #     # print(topology_list)

                net_list = list(net_list_temp)
                # print('last_net_list', net_list)
                # print('pre_net_list', pre_net_list)

                next_pin_list = list(next_pin_list_temp)
                # print('last_next_pin_list', next_pin_list)

                # previous_sch_pin_list1 = list(previous_sch_pin_temp1)
                # # print('last_previous_sch_pin_list1', previous_sch_pin_list1)
                #
                # previous_sch_pin_list2 = list(previous_sch_pin_temp2)
                # # print('last_previous_sch_pin_list2', previous_sch_pin_list2)
                #
                # for i in previous_sch_pin_list1:
                #     item = i if i is None else i[-1]
                #     previous_pin_list1.append(item)
                # # print('last_previous_pin_list1', previous_pin_list1)
                #
                # for i in previous_sch_pin_list2:
                #     item = i if i is None else i[-1]
                #     previous_pin_list2.append(item)
                # # print('last_previous_pin_list2', previous_pin_list2)

                sch_list = list(sch_list_temp)
                # # print('last_sch_list', sch_list)
                # if start_net_name == 'USB2_N14':
                #     # print(1111, net_list)
                #     # print(2222, pre_net_list)
                #     # print(333, sch_list)
                #     # print(j)

                for net in net_list:
                    if net in pre_net_list:
                        end = True
                        break
                    if net is not None:
                        end = False
                        break
                    if net is None:
                        end = True

                if j > 20:
                    raise ValueError("!!!!!!!Can't Extract %s-%s!!!!!!!" % (start_sch_name, start_net_name))

            topology_return_half_list = []
            # if start_net_name == 'PCH_HSON2' and start_sch_name == 'JN44':
            #     # print('ok')
            # # print(sch_list[ind])
            # # print(33333, topology_out_list)
            # # print('topology_out_list', topology_out_list)
            if topology_out_list:
                for x in topology_out_list:
                    if x not in topology_return_half_list:
                        topology_return_half_list.append(x)

                if topology_return_list:
                    topology_return_list += topology_return_half_list
                else:
                    topology_return_list = topology_return_half_list
            else:
                if topology_return_list:
                    topology_return_list += topology_list
                else:
                    topology_return_list = topology_list
            # # print('in', topology_return_list)
    # if topology_return_list:
    #     # print('topology_return_list', topology_return_list)
    #     # print('')
    # # print('topology_return_list', topology_return_list)
    # topology_output_list = []
    # if topology_return_list:
    #     for x in topology_return_list:
    #         if x not in topology_output_list:
    #             topology_output_list.apppend(x)
    # print(topology_output_list)
    return topology_return_list
    # else:
    #     # print('topology_list', topology_list)
    #     # print('')
    #     return topology_list


def corss_over_check(topology_list):
    global All_Net_List, All_Layer_List
    topology_list_cross = []
    for idx1 in xrange(len(topology_list)):
        topology_list_copy = copy.deepcopy(topology_list)
        topology_list_unit_check0 = copy.deepcopy(topology_list[idx1])
        topology_list_unit_check = copy.deepcopy(topology_list[idx1])
        topology_list_copy.remove(topology_list_unit_check)
        cross_idx = []
        for idx2 in xrange(len(topology_list_copy)):
            branch_len = min(len(topology_list_copy[idx2]), len(topology_list_unit_check0))
            for idx3 in xrange(branch_len):
                try:
                    if topology_list_unit_check0[idx3] and topology_list_copy[idx2][idx3]:
                        if topology_list_unit_check0[idx3] != topology_list_copy[idx2][idx3]:
                            # print('A:', topology_list_unit_check0)
                            # print('B:', topology_list_copy[idx2])

                            if str(topology_list_unit_check0[idx3]).find('GND') > -1 and idx3 not in cross_idx:
                                if str(topology_list_unit_check0[idx3 - 2]).split(':')[-1].find('[') == 0:
                                    topology_list_unit_check[idx3 - 2] = ':'.join(
                                        topology_list_unit_check[idx3 - 2].split(':') + ['$CROSS'])
                                cross_idx.append(idx3)
                                break
                            elif str(topology_list_unit_check0[idx3]).find(':') > -1 and \
                                    str(topology_list_unit_check0[idx3]).split(':')[0].find(
                                            '[') > -1 and idx3 not in cross_idx:
                                if str(topology_list_unit_check0[idx3]).split(':')[0].find('[') == 0:
                                    topology_list_unit_check[idx3] = ':'.join(
                                        ['CROSS$'] + topology_list_unit_check[idx3].split(':'))
                                if str(topology_list_unit_check0[idx3 - 2]).split(':')[-1].find('[') == 0:
                                    topology_list_unit_check[idx3 - 2] = ':'.join(
                                        topology_list_unit_check[idx3 - 2].split(':') + ['$CROSS'])
                                cross_idx.append(idx3)
                                break
                            elif str(topology_list_unit_check0[idx3]).find(':') > -1 and \
                                    str(topology_list_unit_check0[idx3]).split(':')[
                                        0] in All_Layer_List and idx3 not in cross_idx:
                                topology_list_unit_check[idx3] = ':'.join(
                                    ['CROSS$'] + topology_list_unit_check[idx3].split(':'))
                                cross_idx.append(idx3)
                                break
                            elif str(topology_list_unit_check0[idx3]) in All_Net_List and idx3 not in cross_idx:
                                if str(topology_list_unit_check0[idx3 - 2]).split(':')[-1].find('[') == 0:
                                    topology_list_unit_check[idx3 - 2] = ':'.join(
                                        topology_list_unit_check[idx3 - 2].split(':') + ['$CROSS'])
                                cross_idx.append(idx3)
                                break
                            # elif idx3 in cross_idx:   #当与上一个分叉相同时，退出
                            #     break
                            else:
                                # print('diff_A:', topology_list_unit_check0[idx3])
                                # print('diff_B:', topology_list_copy[idx2][idx3])
                                break
                        else:  # 过孔同层分叉
                            if str(topology_list_unit_check0[idx3]).split(':')[0] in All_Layer_List and \
                                    topology_list_unit_check0[idx3 + 1] != topology_list_copy[idx2][idx3 + 1] \
                                    and idx3 not in cross_idx:
                                topology_list_unit_check[idx3] = ':'.join(
                                    ['CROSS$'] + topology_list_unit_check[idx3].split(':'))
                                cross_idx.append(idx3)
                                break
                            # elif idx3 in cross_idx:   #当与上一个分叉相同时，退出
                            #     break

                except:
                    print("ununsxpected error:", sys.exc_info())
                    raise
        # print('C:', topology_list_unit_check)
        topology_list_cross.append(topology_list_unit_check)

    # 补齐没有显示分叉的分支
    for topology_branch in topology_list_cross:
        topology_list_cross_copy = copy.deepcopy(topology_list_cross)
        if str(topology_branch[-2]).split(':')[-1] != '$CROSS':
            topology_list_cross_copy.remove(topology_branch)
            last_part = str(topology_branch[-2]).split(':')[-1]
            for idx2 in xrange(len(topology_list_cross_copy)):
                if len(topology_list_cross_copy[idx2]) > len(topology_branch) + 1:
                    if str(topology_list_cross_copy[idx2][len(topology_branch)]).split(':')[0] == 'CROSS$' and \
                            str(topology_list_cross_copy[idx2][len(topology_branch)]).split(':')[1] == last_part:
                        topology_branch[-2] = ':'.join(topology_branch[-2].split(':') + ['$CROSS'])
                        # print('D:', topology_branch)
                        # print("D:", topology_list_cross)
                        break

    topology_list_cross.sort()
    return topology_list_cross


# 简化 topology_extract2 生成的数据格式
# 返回每个信号线的起始芯片和终止芯片，换层次数，总长，加上topology_extract2的数据
def topology_list_format_simplified(topology_list):
    # # print(1)
    # Topology Formatting Function
    # start_time = time.clock()
    topology_out_list = []
    # # print(1, topology_list)
    global All_Net_List, All_Layer_List, non_signal_net_list
    signal_net_list = []

    start1_time = time.clock()
    # for n1 in All_Net_List:
    #     if n1 not in non_signal_net_list:
    # signal_net_list = [i for i in All_Net_List if i not in non_signal_net_list]
    # 这种方法要优于立即执行函数表达式
    signal_net_list = list(set(All_Net_List) ^ set(non_signal_net_list))
    topology_list = corss_over_check(topology_list)
    # start2_time = time.clock()
    # # print('sm1', start2_time - start1_time)
    # , layout_error_diff_list, layout_error_se_list
    # count = 1
    # # print(count)
    # count += 1
    # Calculation Total Length and Via Count for each line
    for idx1 in xrange(len(topology_list)):
        via_count = 0
        total_length = 0
        net_count = 0
        remove_ind = None
        # # print(topology_list)

        if str(topology_list[idx1][-1]).find('Exception;') > -1:
            # # print(2, topology_list)
            end_sch_name = topology_list[idx1][-1]
            topology_list[idx1] = [topology_list[idx1][0], 'via_count', 'total_length'] + topology_list[idx1][1::]
        else:
            # 获得信号线最后芯片名并去除[]符号
            # # print(0, topology_list[idx1])
            # # print(1, topology_list[idx1][-2])
            # # print(2, str(topology_list[idx1][-2]).split(':')[-1].find('['))
            if str(topology_list[idx1][-2]).split(':')[-1].find('$') == 0 and str(topology_list[idx1][-2]).split(':')[
                -2].find('[') == 0:  # f7684584
                end_sch_name = topology_list[idx1][-2].split(':')[-2][1:-1]

            elif str(topology_list[idx1][-3]).split(':')[-1].find('$') == 0 and str(topology_list[idx1][-3]).split(':')[
                -2].find('[') == 0:  # f7684584
                end_sch_name = topology_list[idx1][-3].split(':')[-2][1:-1]
            elif str(topology_list[idx1][-2]).split(':')[-1].find('[') == 0:
                end_sch_name = topology_list[idx1][-2].split(':')[-1][1:-1]

            elif str(topology_list[idx1][-3]).split(':')[-1].find('[') == 0:
                end_sch_name = topology_list[idx1][-3].split(':')[-1][1:-1]
                # # print(end_sch_name)
            else:
                end_sch_name = 'NONE'
            # # print(333, end_sch_name)
            for idx2 in xrange(len(topology_list[idx1])):
                if isfloat(topology_list[idx1][idx2]):
                    total_length += float(topology_list[idx1][idx2])

                if str(topology_list[idx1][idx2]).find(':') > -1:
                    # eg: '[FRONT_USB_HEADER-2]:BOTTOM:5.1'
                    layer = None
                    for x in topology_list[idx1][idx2].split(':'):
                        if x in All_Layer_List:
                            layer = x

                    layer_next = None
                    for x in topology_list[idx1][idx2 + 2::]:
                        if str(x).find(':') > -1:
                            for x_ in x.split(':'):
                                if x_ in All_Layer_List:
                                    layer_next = x_
                                    # 换层计数
                                    if layer_next != layer:
                                        via_count += 1
                            break
            # via_count 是换层的次数，total_length 是此条信号线加下条信号线的总长度（如果有下条信号线的话）
            topology_list[idx1] = [topology_list[idx1][0], 'via_count %d' % via_count,
                                   'total_length %.3f' % total_length] + topology_list[idx1][1::]
            # # print(2, topology_list)
        # print(1, topology_list[idx1][3])
        topology_1 = copy.deepcopy(topology_list[idx1][3])
        if topology_1.find('CROSS$') > -1:
            topology_1 = topology_1[7:]
        if topology_1.find('$CROSS') > -1:
            topology_1 = topology_1[:-7]
        if topology_1.find('[') == 0:
            # 获得信号线起始芯片名并去除[]符号
            start_sch_name = topology_1.split(':')[0][1:-1]

        topology_list[idx1] = [start_sch_name, topology_list[idx1][0], end_sch_name] + topology_list[idx1][1::]
        # # print(3, topology_list)

        for idx2 in xrange(len(topology_list[idx1])):
            # 过滤topology_list中的非信号线
            if topology_list[idx1][idx2] in signal_net_list:
                net_count += 1
                if net_count > 1:
                    # 标出信号线
                    topology_list[idx1][idx2] = 'net$%s' % topology_list[idx1][idx2]
            elif topology_list[idx1][idx2] in non_signal_net_list:
                # 删除非信号线
                remove_ind = idx2
        if remove_ind != None:
            # 其实就是删除GND信号
            # # print(topology_list[idx1].pop(remove_ind))
            topology_list[idx1].pop(remove_ind)
        # # print(4, topology_list)

        topology_out_list.append(topology_list[idx1])
    # # print(topology_out_list)
    topology_out_list.sort()
    # # print(2, topology_list)
    # layout_error_diff_list = []
    # layout_error_se_list = []
    # end_time1 = time.clock()
    # # print('sm1', end_time1 - start_time)
    # Add syntax to judge the net change condition and remove the non-signal net from the end
    # signal_net_list = []
    # for n1 in All_Net_List:
    #     if n1 not in non_signal_net_list:
    #         signal_net_list.append(n1)
    #
    # for idx1 in xrange(len(topology_list)):
    #     net_count = 0
    #     remove_ind = None
    #     for idx2 in xrange(len(topology_list[idx1])):
    #         # 过滤topology_list中的非信号线
    #         if topology_list[idx1][idx2] in signal_net_list:
    #             net_count += 1
    #             if net_count > 1:
    #                 # 标出信号线
    #                 topology_list[idx1][idx2] = 'net$%s' % topology_list[idx1][idx2]
    #         elif topology_list[idx1][idx2] in non_signal_net_list:
    #             # 删除非信号线
    #             remove_ind = idx2
    #     if remove_ind != None:
    #         # 其实就是删除GND信号
    #         # # print(topology_list[idx1].pop(remove_ind))
    #         topology_list[idx1].pop(remove_ind)
    #
    # topology_list.sort()
    # # print(3, topology_list)

    # Add Start SCH, End SCH Name for each line
    # for idx in xrange(len(topology_list)):
    #     if topology_list[idx][3].find('[') == 0:
    #         # 获得信号线起始芯片名并去除[]符号
    #         start_sch_name = topology_list[idx][3].split(':')[0][1:-1]
    #         # # print(topology_list[idx][3].split(':')[0][1:-1])
    #         # 如果存在Exception;，他表示的是信号线某段的长度
    #     if str(topology_list[idx][-1]).find('Exception;') > -1:
    #         end_sch_name = topology_list[idx][-1]
    #     else:
    #         # 获得信号线最后芯片名并去除[]符号
    #         if str(topology_list[idx][-2]).split(':')[-1].find('[') == 0:
    #             end_sch_name = topology_list[idx][-2].split(':')[-1][1:-1]
    #         else:
    #             end_sch_name = None
    #
    #     topology_list[idx] = [start_sch_name, topology_list[idx][0], end_sch_name] + topology_list[idx][1::]
    # # print(3, topology_list)

    # end_time1 = time.clock()
    # # print('toposm', end_time1 - start1_time)

    # # print('topology_list', sorted(topology_list))
    return sorted(topology_list)


# xlwings function

# 清空隐藏表格内容
def ACT_sheet_content_clear():
    wb = Book(xlsm_path).caller()
    wb.sheets['ACT_diff'].range('A1').current_region.clear()
    wb.sheets['ACT_diff_count'].range('A1').current_region.clear()
    wb.sheets['ACT_diff_count'].range('E1').current_region.clear()
    wb.sheets['ACT_diff_count'].range('I1').current_region.clear()
    wb.sheets['ACT_se'].range('A1').current_region.clear()
    wb.sheets['ACT_se_count'].range('A1').current_region.clear()
    wb.sheets['ACT_se_count'].range('E1').current_region.clear()
    wb.sheets['ACT_se_count'].range('I1').current_region.clear()


def GetCellIdx(sheet, key_word, direction, offset):
    idx = None
    for cell in sheet.api.UsedRange.Cells:
        if cell.Value == key_word:
            idx = (cell.Row, cell.Column)
    if idx != None:
        for i in xrange(offset):
            if direction == 'r':
                idx = (idx[0], idx[1] + 1)
            elif direction == 'l':
                idx = (idx[0], idx[1] - 1)
            elif direction == 'u':
                idx = (idx[0] - 1, idx[1])
            elif direction == 'd':
                idx = (idx[0] + 1, idx[1])

    return idx


# 设置单元格内容的字型字体大小和字体位置
def SetCellFont(sheet, cell_ind, Font_Name, Font_Size, horizon_alignment):
    sheet.range(cell_ind).api.Font.Name = Font_Name
    sheet.range(cell_ind).api.Font.Size = Font_Size
    if horizon_alignment == 'c':
        sheet.range(cell_ind).api.HorizontalAlignment = Constants.xlCenter
    elif horizon_alignment == 'r':
        sheet.range(cell_ind).api.HorizontalAlignment = Constants.xlRight
    elif horizon_alignment == 'l':
        sheet.range(cell_ind).api.HorizontalAlignment = Constants.xlLeft


def SetCellFont_current_region(sheet, start_cell_ind, Font_Name, Font_Size, horizon_alignment):
    sheet.range(start_cell_ind).current_region.api.Font.Name = Font_Name
    sheet.range(start_cell_ind).current_region.api.Font.Size = Font_Size
    if horizon_alignment == 'c':
        sheet.range(start_cell_ind).current_region.api.HorizontalAlignment = Constants.xlCenter
    elif horizon_alignment == 'r':
        sheet.range(start_cell_ind).current_region.api.HorizontalAlignment = Constants.xlRight
    elif horizon_alignment == 'l':
        sheet.range(start_cell_ind).current_region.api.HorizontalAlignment = Constants.xlLeft


def SetCellBorder(sheet, cell_ind):
    sheet.range(cell_ind).api.Borders.LineStyle = LineStyle.xlContinuous


def SetCellNoBorder(sheet, cell_ind):
    sheet.range(cell_ind).api.Borders.LineStyle = LineStyle.xlLineStyleNone


# 设置表格边框
def SetCellBorder_current_region(sheet, start_cell_ind):
    sheet.range(start_cell_ind).current_region.api.Borders.LineStyle = LineStyle.xlContinuous


# Set1
def LoadAllegroFile():
    # Import the layout design file and extract the data
    wb = Book(xlsm_path).caller()

    active_sheet = wb.sheets.active  # Get the active sheet object

    # 按钮调用VBA的GetOpenFilename()方法：显示标准的打开对话框，并获取用户文件名，括号内内容是指定文档名与后缀名
    brd_path = wb.app.api.GetOpenFilename("Allegro PCB Design File (*.brd), *.brd")
    if brd_path:
        # 输出为 C:\Users\admin\Desktop\VGA_BOX_TBT_IO_X01_layout_20161027.brd，换为 /
        brd_path = brd_path.replace("\\", "/")

        # UsedRange属性：引用当前工作表中所有已使用的单元格区域 Rows Columns Cells
        for cell in active_sheet.api.UsedRange.Cells:

            # 获得图片名并写入
            if cell.Value == "Allegro PCB Design File Path:":
                active_sheet.range((cell.Row, cell.Column + 1)).value = brd_path
            # for cell in active_sheet.api.UsedRange.Cells:
            # 取得报告行的单元格坐标
            if cell.Value == "Allegro Report File Path:":
                report_path_idx = (cell.Row, cell.Column + 1)
    else:
        return
    #####################################
    # 与表 SymbolList 进行交互，没有用到，先删除
    # for cell in wb.sheets["SymbolList"].api.UsedRange.Cells:
    #     if cell.Value == 'Model Name':
    #         sym_list_idx = (cell.Row+1, cell.Column)

    # 清除表 SymbolList 中从 Model Name 开始的表格
    # wb.sheets["SymbolList"].range(sym_list_idx).expand('table').clear()
    #####################################

    report_list = ['eld.rpt', 'dpg.rpt', 'spn.rpt', 'x_sec.rpt', 'npr.rpt']

    command_list = ['eld', 'dpg', 'spn', 'x-section', 'npr']

    # os.path.dirname(__file__)：返回脚本的路径 C:/Users/admin/Desktop
    # os.path.basename(__file__): 返回最后的文件名
    RootPath = os.path.dirname(brd_path)

    # 返回 VGA_BOX_TBT_IO_X01_layout_20161027.brd
    BrdName = os.path.basename(brd_path)
    # # print(RootPath,BrdName)
    # C:/Users/admin/Desktop/eld.rpt
    # 合成一个总报告
    report_path_list = [os.path.join(RootPath, x) for x in report_list]

    # 生成报告的指令集，eg: report -v eld "C:/Users/admin/Desktop/VGA_BOX_TBT_IO_X01_layout_20161027.brd" "
    # C:/Users/admin/Desktop/eld.rpt"
    complete_command_list = ['report -v %s "%s" "%s"' % (command_list[idxx1], brd_path, os. \
                                                         path.join(RootPath, report_list[idxx1])) for idxx1 in
                             xrange(len(command_list))]

    # os.remove 删除缓存报告，找不到错误信息返回nil并加上错误信息
    try:
        for path in report_path_list:
            os.remove(path)

    # 系统找不到 指定文件会触发 WindowsError
    except WindowsError:
        pass

    # 要输出的报告文件绝对
    output_file_path = RootPath + '/%s_allegro_report.rpt' % BrdName

    # 删除之前的缓存报告文件
    try:
        os.remove(output_file_path)
    except WindowsError:
        pass

    # os.system 相当于在windows的cmd窗口中输入命令
    # 运行 report -v 命令生成报告
    for command in complete_command_list:
        os.system(command)

    # 读取生成报告中的数据
    data_list = [open(path, 'r').read() for path in report_path_list]
    # 将四个报告中的数据存入输出报告文件中
    f = open(output_file_path, 'w')

    # 数据整理，将有用信息提取出，两种不同的提取方法，适用于不同格式的报告
    # 加密代码
    content = ''
    for d in data_list:
        # 对走线报告进行分析
        if 'Detailed Etch Length by Layer and Width Report' in d or \
                'Detailed Trace Length by Layer and Width Report' in d:
            d = d.split('\n')
            # d_tmp_1
            d_tmp = list(map(lambda x: x.split(','), d[5:-1]))
            # test_file = open( RootPath + "/d_tmp.txt", 'w')
            # test_file.write(str(d_tmp))
            # test_file.close()
            xy_point_pattern = re.compile(r'.*?xprobe:xy:\((.*?)\).*?xprobe:xy:\((.*?)\).*$')
            # # print(xy_point_pattern)
            for idx in xrange(len(d_tmp)):
                # 返回所匹配模式的列表: [(x1,y1)],[(x2,y2)]
                xy_point = re.findall(xy_point_pattern, d_tmp[idx][-1])
                # # print(xy_point)
                d_tmp[idx][-1] = ' '.join([xy_point[0][0], xy_point[0][1]])
            # d_tmp_2
            d_tmp = list(map(lambda x: ','.join(x), d_tmp))
            # d_tmp_3
            d = '\n'.join(d[0:5] + d_tmp) + '\n'
            # # print(d)
        # 对快速报告进行分析
        elif 'Allegro Report' in d:  # or 'Diffpair Gap Report' in d:
            d = d.split('\n')
            d_tmp = d[5:-1]
            d_tmp_ = list()
            for str1 in d_tmp:
                # find方法寻找字符串，找到返回位置索引，找不到返回-1
                if str1.find('"') > -1:
                    d_tmp_.append(','.join(re.split('\s\(|\)",|,\s', str1)[1:3]))
                else:
                    d_tmp_.append(','.join(re.split(',', str1)[0:1]))
            # # print(d_tmp_)
            # 分离差分对名称
            d_tmp_ = sorted(list(set(d_tmp_)))
            d = '\n'.join(d[0:5] + d_tmp_) + '\n'
        elif 'Properties on Nets Report' in d:
            d = d.split('\n')
            d_tmp = d[5:-1]
            d_tmp_ = []
            for npr_item in d_tmp:
                npr_item_list = npr_item.split(',')
                d_tmp_.append(npr_item_list[0] + ',' + npr_item_list[1] + ',' + npr_item_list[25])

            d = '\n'.join(d[0:5] + d_tmp_) + '\n'

        content += d

    # mycrypt = MyCrypt(b'\x17\xBB\x50\xEA\x20\xA7\x4D\xE5\x2F\x7F\x29\x4C\x96\x7D\xE5\xA5')

    # content = mycrypt.myencrypt(content)
    f.write(content)
    f.close()

    # 移除五个分报告
    try:
        for path in report_path_list:
            os.remove(path)
    except WindowsError:
        pass

    # 写入报告名到单元格中
    active_sheet.range(report_path_idx).value = output_file_path

    # 清空隐藏表格内容
    ACT_sheet_content_clear()

    ClearNetList()


# setting表中获取键
def GetSetting():
    start_time = time.clock()
    # Get Setting Date Function 2
    wb = Book(xlsm_path).caller()
    setting_sheet = wb.sheets['Setting']

    # 存放生成报告的绝对路径
    allegro_report_path = None

    # Layer Type Definition Name 是不变的，所以用元组存放
    layer_type_dict = dict()
    # Start Component Name List 是每次改变的所以用列表存放
    start_sch_name_list, start_sch_name_list_ind = list(), None
    progress_ind = None

    # 有没有比 UsedRange 更高效的方法？？？
    # 存Allegro Report File Path:，Layer Type Definition:和Start Component Name List:的数据
    for cell in setting_sheet.api.UsedRange.Cells:

        # 获取报告绝对路径
        if cell.Value == 'Allegro Report File Path:':
            allegro_report_path = setting_sheet.range((cell.Row, cell.Column + 1)).value
            break
    for cell in setting_sheet.api.UsedRange.Cells:
        # start1_time = time.clock()
        # # print(1, start1_time - start_time)
        # 获取叠层名称字典
        if cell.Value == 'Layer Type Definition:':
            layer_type_table = setting_sheet.range((cell.Row + 1, cell.Column)).options(expand='table', ndim=2).value
            # # print(layer_type_table)
            try:
                layer_type_dict = dict(layer_type_table)
            except:
                layer_type_dict = dict()
            break
    for cell in setting_sheet.api.UsedRange.Cells:
        # start2_time = time.clock()
        # # print(2, start2_time - start1_time)
        # 获取元件名称列表
        if cell.Value == 'Start Component Name List:':
            start_sch_name_list_ind = (cell.Row + 1, cell.Column)
            start_sch_name_list = setting_sheet.range(start_sch_name_list_ind).options(expand='table', ndim=1).value
            break
        # start3_time = time.clock()
        # # print(3, start3_time - start2_time)
    for cell in setting_sheet.api.UsedRange.Cells:
        if cell.Value == 'Progress:':
            progress_ind = (cell.Row, cell.Column + 1)
            SetCellFont(setting_sheet, progress_ind, 'Times New Roman', 16, 'l')
            break
    #     start4_time = time.clock()
    #     # print(4, start4_time - start3_time)
    # start5_time = time.clock()
    # # print(5, start5_time - start_time)
    # set:无序不重复的元素集
    All_Layer_List = list(set(layer_type_dict.keys()))

    end_time = time.clock()
    # print('GetSetting', end_time - start_time)
    return allegro_report_path, layer_type_dict, start_sch_name_list, progress_ind, All_Layer_List


def GetSetting_DDR():
    # Get Setting Function for DDR
    wb = Book(xlsm_path).caller()
    setting_sheet = wb.sheets['Setting']

    allegro_report_path = None
    layer_type_dict = dict()
    ddr_prefix = None
    cpu_list, dram_list = list(), list()
    progress_ind = None

    for cell in setting_sheet.api.UsedRange.Cells:
        if cell.Value == 'Allegro Report File Path:':
            allegro_report_path = setting_sheet.range((cell.Row, cell.Column + 1)).value
            break
    for cell in setting_sheet.api.UsedRange.Cells:
        if cell.Value == 'Layer Type Definition:':
            layer_type_table = setting_sheet.range((cell.Row + 1, cell.Column)).options(expand='table', ndim=2).value
            try:
                layer_type_dict = dict(layer_type_table)
            except:
                layer_type_dict = dict()
            break
    for cell in setting_sheet.api.UsedRange.Cells:
        if cell.Value == 'DDR Pre-fix':
            ddr_prefix = setting_sheet.range((cell.Row + 1, cell.Column)).value
            break
    for cell in setting_sheet.api.UsedRange.Cells:
        if cell.Value == 'CPU List':
            cpu_list = setting_sheet.range((cell.Row + 1, cell.Column)).options(expand='table', ndim=1).value
            break
    for cell in setting_sheet.api.UsedRange.Cells:
        if cell.Value == 'DRAM List':
            dram_list = setting_sheet.range((cell.Row + 1, cell.Column)).options(expand='table', ndim=1).value
            break
    for cell in setting_sheet.api.UsedRange.Cells:
        if cell.Value == 'Progress:':
            progress_ind = (cell.Row, cell.Column + 1)
            SetCellFont(setting_sheet, progress_ind, 'Times New Roman', 16, 'l')
            break

    All_Layer_List = list(set(layer_type_dict.keys()))

    return allegro_report_path, layer_type_dict, ddr_prefix, cpu_list, dram_list, progress_ind, All_Layer_List


# Netlist  辨别 diff., SE., Non-Signal,并导出在NetList表中
def NetTypeDetect():
    # Classify the net type: diff., SE., Non-Signal
    global All_Net_List
    start_time = time.clock()

    wb = Book(xlsm_path).caller()

    netlist_sheet = wb.sheets['NetList']

    # 获取键值
    allegro_report_path, layer_type_dict, start_sch_name_list, progress_ind, All_Layer_List = GetSetting()
    SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data, npr_brd_data = \
        read_allegro_data(allegro_report_path)

    # 获取所有信号和差分对的名称
    # 所以all_net_list存放的是所有与pin脚相连的线名
    All_Net_List = getallnetlist(SCH_brd_data)
    diff_pair_dict, npr_diff_pair_net_list = diff_detect(diff_pair_brd_data, npr_brd_data)
    # # print('npr_diff_pair_net_list', npr_diff_pair_net_list)
    # start1_time = time.clock()
    unmachable_diff_net_list_ind = None
    # # print(1, start1_time - start_time)
    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'Differential':
            differential_ind = (cell.Row + 1, cell.Column)
            break
    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'Single-Ended':
            single_ended_ind = (cell.Row + 1, cell.Column)
            break
    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'Non-Signal Net List':
            non_signal_net_list_ind = (cell.Row + 1, cell.Column)
            break
    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'Unmatched Differential Net List':
            unmachable_diff_net_list_ind = (cell.Row + 1, cell.Column)
            break
    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'User-Defined Differential Net List':
            user_defined_differential_ind = (cell.Row + 1, cell.Column)
            break
    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value in ['User-Defined not Non-Signal Net List', 'User-Defined Single-Ended Net List']:
            user_defined_single_end_signal_list_ind = (cell.Row + 1, cell.Column)
            break
    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'User-Defined Non-Signal Net List':
            user_defined_non_signal_list_ind = (cell.Row + 1, cell.Column)
            break
    # 清除缓存的表数据
    netlist_sheet.range(differential_ind).expand('table').clear()
    netlist_sheet.range(single_ended_ind).expand('table').clear()
    netlist_sheet.range(non_signal_net_list_ind).expand('table').clear()
    if unmachable_diff_net_list_ind:
        netlist_sheet.range(unmachable_diff_net_list_ind).expand('table').clear()

    # 人工输入单根和差分
    user_defined_differential_list = netlist_sheet.range(user_defined_differential_ind).options(expand='table',
                                                                                                ndim=2).value
    user_defined_single_end_signal_list = netlist_sheet.range(user_defined_single_end_signal_list_ind).options(
        expand='table', ndim=1).value
    # 去重
    user_defined_differential_set_list = []
    for x in user_defined_differential_list:
        if x not in user_defined_differential_set_list:
            user_defined_differential_set_list.append(x)
    user_defined_differential_list = user_defined_differential_set_list
    # diff_pair_dict = diff_detect(diff_pair_brd_data)
    # # 将差分信号中的其他信号删除
    for cut_net in list(_flatten(user_defined_differential_list)) + user_defined_single_end_signal_list:
        try:
            del diff_pair_dict[cut_net]
        except:
            pass
    diff_list = []
    for x in diff_pair_dict.keys():
        diff_list.append([x, diff_pair_dict[x]])

    # # print(diff_list)
    def find_last(string, str):
        last_position = -1
        while True:
            position = string.find(str, last_position + 1)
            if position == -1:
                return last_position
            last_position = position

    # start2_time = time.clock()
    # # print(2, start2_time - start1_time)
    # temp1 = All_Net_List
    ################################################################
    # temp1中存放的是所有信号，temp2中存放的是带N的信号，temp3中存放的是其他信号
    # temp2, temp3, temp4 = [], [], []
    # for i in xrange(0, len(temp1)):
    #     # N只有一种的情况
    #     if temp1[i].count("N") == 1:
    #         temp2.append(temp1[i])
    #     elif temp1[i].count("N") > 1:
    #         temp3.append(temp1[i])
    #
    # # 判断是否是差分对
    # # 一个N时
    # for i in xrange(len(temp2)):
    #     trans2_data = temp2[i].replace("N", "P")
    #     if trans2_data in temp1:
    #         temp4.append([temp2[i], trans2_data])
    #
    # # 多个N时判断最后一个
    # for i in xrange(len(temp3)):
    #     last_ind = find_last(temp3[i], 'N')
    #     list3 = list(temp3[i])
    #     list3[last_ind] = 'P'
    #     trans3_data = ''.join(list3)
    #
    #     if trans3_data in temp1:
    #         temp4.append([temp3[i], trans3_data])
    ###########################################################
    # 判断是一个差分对（P-N），四层循环时间复杂度太高，需修改
    ##############################
    # for i in xrange(0, len(temp2)):
    #     for j in xrange(0, len(temp1)):
    #         for z in xrange(0, len(temp2[i])):
    #             for k in xrange(0, len(temp1[j])):
    #                 if (temp2[i][z] == 'N' and temp1[j][k] == 'P') and (temp2[i][0:z - 1] == temp1[j][0:k - 1])\
    #                     and (temp2[i][z + 1:]) == temp1[j][z + 1:]:
    #                     temp4.append([temp2[i], temp1[j]])
    ##############################

    # # print(diff_no_brackets_list)

    # 防止自动生成的差分报告不准确
    # for i in temp4:
    #     if [i[0]] in diff_no_brackets_list:
    #         # # print([i[0]])
    #         # # print(i)
    #         diff_list.append(i)
    #     elif [i[1]] in diff_no_brackets_list:
    #         diff_list.append(i)
    #         # # print([i[1]])

    diff_list = sorted(diff_list)
    # # print(diff_list)
    # start3_time = time.clock()
    # # print(3, start3_time - start2_time)
    # Regular Expression for quick define single ended signal net list
    # findall返回所匹配的字符串
    net_list_tmp = list()
    for net_tmp in user_defined_single_end_signal_list:
        if net_tmp != None:
            if net_tmp.find('$') == 0:
                net_list_tmp += [x for x in All_Net_List if re.findall(r'%s' % net_tmp[1::], x) != []]
    user_defined_single_end_signal_list += net_list_tmp
    # # print('user_defined_single_end_signal_list', user_defined_single_end_signal_list)
    # 验证用户输入的单根信号在所有信号中
    user_defined_single_end_signal_list = list(set(All_Net_List) & set(user_defined_single_end_signal_list))
    # # print('user_defined_single_end_signal_list', user_defined_single_end_signal_list)
    user_defined_single_end_signal_list.sort()

    # 获取用户输入的非单根信号
    user_defined_non_signal_list = netlist_sheet.range(user_defined_non_signal_list_ind).options(expand='table',
                                                                                                 ndim=1).value
    # 验证用户输入的非单根信号在所有信号中
    user_defined_non_signal_list = list(set(All_Net_List) & set(user_defined_non_signal_list))
    # # print(' user_defined_non_signal_list',  user_defined_non_signal_list)

    # Update the "diff_pair_dict" with the "user_defined_differential_list"
    if user_defined_differential_list not in [[], '', None, [[None]]]:
        user_defined_diff_list1 = []
        user_defined_diff_list2 = []
        for dp in user_defined_differential_list:
            user_defined_diff_list1.append(dp[0])
            user_defined_diff_list2.append(dp[1])

            user_defined_diff_pair_dict = dict()
        for idx_dpp in xrange(len(user_defined_diff_list1)):
            user_defined_diff_pair_dict[user_defined_diff_list1[idx_dpp]] = user_defined_diff_list2[idx_dpp]
            user_defined_diff_pair_dict[user_defined_diff_list2[idx_dpp]] = user_defined_diff_list1[idx_dpp]

        used_diff_list, user_defined_diff_pair_list = list(), list()

        for x in user_defined_diff_pair_dict.keys():
            if x in All_Net_List:
                user_defined_diff_pair_list.append([x, user_defined_diff_pair_dict[x]])

        diff_list += user_defined_diff_pair_list

        diff_list = sorted(diff_list)
    # Get the "non_signal_net_list"
    non_signal_net_list, _, _ = get_exclude_netlist(All_Net_List, npr_brd_data)
    # Update the "non_signal_net_list" with the "user_defined_non_signal_list"

    non_signal_net_list += user_defined_non_signal_list

    # Update the diff_pair_list, and non_signal_net_list from user-defined list
    # # print(All_Net_List)
    # # print(non_signal_net_list)
    non_signal_net_list = list(set(All_Net_List) & set(non_signal_net_list) -
                               set(user_defined_single_end_signal_list) - set(_flatten(user_defined_differential_list)))
    # # print('non_signal_net_list', non_signal_net_list)
    non_signal_net_list = sorted(non_signal_net_list)
    # Get the "single_ended_list"
    # 不在差分线和非信号线中即定义为单根信号
    # # print(set(All_Layer_List))
    # # print(set(diff_list))
    single_ended_half_list = set(All_Net_List) - set(_flatten(diff_list))
    if unmachable_diff_net_list_ind:
        single_ended_half_list = set(single_ended_half_list) - set(npr_diff_pair_net_list)
    # # print(_flatten(diff_list))
    # # print(_flatten(single_ended_half_list))
    single_ended_list = list(set(single_ended_half_list) - set(non_signal_net_list))
    single_ended_list = sorted(set(single_ended_list))

    # 不知道这个判断有什么意义
    #######################################
    # if user_defined_single_end_signal_list != list(): # 不为空
    #     user_defined_single_end_signal_list = [[x] for x in user_defined_single_end_signal_list]
    #     netlist_sheet.range(user_defined_single_end_signal_list_ind).expand('table').clear()
    #     netlist_sheet.range(user_defined_single_end_signal_list_ind).expand(
    #         'table').value = user_defined_single_end_signal_list
    #
    # if user_defined_non_signal_list != list():
    #     user_defined_non_signal_list = [[x] for x in user_defined_non_signal_list]
    #     netlist_sheet.range(user_defined_non_signal_list_ind).expand('table').clear()
    #     netlist_sheet.range(user_defined_non_signal_list_ind).expand('table').value = user_defined_non_signal_list
    #
    # if user_defined_differential_list not in [[], '', None, [[None]]]:
    #     netlist_sheet.range(user_defined_differential_ind).expand('table').clear()
    #     netlist_sheet.range(user_defined_differential_ind).expand('table').value = user_defined_diff_pair_list
    #########################################
    single_ended_list = [[x] for x in single_ended_list]

    netlist_sheet.range(single_ended_ind).expand('table').value = single_ended_list
    netlist_sheet.range(differential_ind).expand('table').value = diff_list

    non_signal_net_list = [[x] for x in non_signal_net_list]
    npr_diff_pair_net_list = [[x] for x in npr_diff_pair_net_list]
    netlist_sheet.range(non_signal_net_list_ind).expand('table').value = non_signal_net_list
    if unmachable_diff_net_list_ind:
        netlist_sheet.range(unmachable_diff_net_list_ind).expand('table').value = npr_diff_pair_net_list

    netlist_sheet.api.Cells.Font.Name = 'Times New Roman'
    netlist_sheet.api.Cells.Font.Size = 12
    end_time = time.clock()

    # print('NetTypeDetect', end_time - start_time)


# clear Netlist 表
def ClearNetList():
    # Clear the Net List
    # 代码可优化
    wb = Book(xlsm_path).caller()
    netlist_sheet = wb.sheets['NetList']

    # 代码可优化
    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'Differential':
            diff_idx = (cell.Row + 1, cell.Column)
        # for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'Single-Ended':
            se_idx = (cell.Row + 1, cell.Column)
        # for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'Non-Signal Net List':
            non_sig_idx = (cell.Row + 1, cell.Column)

    netlist_sheet.range(diff_idx).expand('table').clear()
    netlist_sheet.range(se_idx).expand('table').clear()
    netlist_sheet.range(non_sig_idx).expand('table').clear()


# 这个功能表格没怎么用，可以先放在一边
# def SymbolListDetect():
#     # Detect the schematic mapping condition
#     wb = Book(xlsm_path).caller()
#
#     symbollist_sheet = wb.sheets['SymbolList']
#
#     # Detect Signal and Non-Signal Net List
#     global diff_list, diff_dict, non_signal_net_list, se_list
#     netlist_sheet = wb.sheets['NetList']
#
#     for cell in netlist_sheet.api.UsedRange.Cells:
#         if cell.Value == 'Differential':
#             diff_idx = (cell.Row+1, cell.Column)
#     for cell in netlist_sheet.api.UsedRange.Cells:
#         if cell.Value == 'Single-Ended':
#             se_idx = (cell.Row+1, cell.Column)
#     for cell in netlist_sheet.api.UsedRange.Cells:
#         if cell.Value == 'Non-Signal Net List':
#             non_sig_idx = (cell.Row+1, cell.Column)
#
#     for cell in symbollist_sheet.api.UsedRange.Cells:
#         if cell.Value == 'Model Name':
#             sym_list_idx = (cell.Row+1, cell.Column)
#
#     for cell in wb.sheets['Setting'].api.UsedRange.Cells:
#         if cell.Value == 'Allegro Report File Path:':
#             report_path_idx = (cell.Row, cell.Column+1)
#
#
#     symbollist_sheet.range(sym_list_idx).expand('table').clear()
#
#     diff_list1 = []
#     diff_list2 = []
#     for dp in netlist_sheet.range(diff_idx).options(expand = 'table', ndim = 2).value:
#         diff_list1.append(dp[0])
#         diff_list2.append(dp[1])
#     diff_list = diff_list1 + diff_list2
#     # # print(diff_list)
#     diff_dict = dict()
#     for idx_dp in xrange(len(diff_list1)):
#         diff_dict[diff_list1[idx_dp]] = diff_list2[idx_dp]
#         diff_dict[diff_list2[idx_dp]] = diff_list1[idx_dp]
#
#
#     se_list = netlist_sheet.range(se_idx).options(expand = 'table', ndim = 1).value
#     se_list = [] if se_list == [None] else se_list
#     non_signal_net_list = netlist_sheet.range(non_sig_idx).options(expand = 'table', ndim = 1).value
#
#     if non_signal_net_list == [None]:
#         non_signal_net_list = []
#
#
#
#     global SCH_object_list, net_object_list, All_Net_List
#     allegro_report_path = wb.sheets['Setting'].range(report_path_idx).value
#     SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data = read_allegro_data(allegro_report_path)
#     SCH_object_list = SCH_detect(SCH_brd_data)
#
#     model_name_list = []
#     # mapped_model_name_list是自己写出的，
#     for x in SCH_object_list:
#         if len(x.GetPinList()) > 2 and x.GetModel() not in mapped_model_name_list:
#             model_name_list.append(x.GetModel())
#
#     result_list = list()
#     for model_name in model_name_list:
#         sch_list = list()
#         for sch in SCH_object_list:
#             if sch.GetModel() == model_name:
#                 sch_list.append(sch.GetName())
#         result_list.append((model_name, ';'.join(sch_list), len(get_SCH_object_by_name(sch_list[0]).GetPinList())))
#     # # print(result_list)
#     # sorted 按key的权值排序
#     result_list = sorted(result_list, key= lambda x: x[2], reverse=True)
#     # # print(result_list)
#     symbollist_sheet.range(sym_list_idx).expand('table').value = result_list
#
#
#     SetCellFont_current_region(symbollist_sheet, sym_list_idx, 'Times New Roman', 12, 'l')
#     SetCellBorder_current_region(symbollist_sheet, sym_list_idx)
# 获得从表中获取 diff., SE., Non-Signal 的数值，并返回
def GetNetList():
    start_time = time.clock()
    # Detect Signal and Non-Signal Net List
    wb = Book(xlsm_path).caller()
    netlist_sheet = wb.sheets['NetList']

    global diff_list, diff_dict, se_list, non_signal_net_list

    diff_list, se_list, non_signal_net_list = list(), list(), list()
    # # print(111)
    for cell in netlist_sheet.api.UsedRange.Cells:
        # # print(2222)
        if cell.Value == 'Differential':
            # # print(333)
            diff_list1 = []
            diff_list2 = []

            # 个人认为，多余操作，后来证明不多余
            #####################################
            for dp in netlist_sheet.range((cell.Row + 1, cell.Column)).options(expand='table', ndim=2).value:
                diff_list1.append(dp[0])
                diff_list2.append(dp[1])
            # [1] + [2] => [1,2]
            diff_list = diff_list1 + diff_list2
            # # print('diff_list1', diff_list1)
            # # print('diff_list2', diff_list2)

            diff_dict = dict()
            for idx_dp in xrange(len(diff_list1)):
                diff_dict[diff_list1[idx_dp]] = diff_list2[idx_dp]
                diff_dict[diff_list2[idx_dp]] = diff_list1[idx_dp]
            # ######################################
            # # print(len(diff_dict.keys()))
            break
    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'Single-Ended':
            se_list = netlist_sheet.range((cell.Row + 1, cell.Column)).options(expand='table', ndim=1).value
            break
    for cell in netlist_sheet.api.UsedRange.Cells:
        if cell.Value == 'Non-Signal Net List':
            non_signal_net_list = netlist_sheet.range((cell.Row + 1, cell.Column)).options(expand='table', ndim=1).value
            break

    end_time = time.clock()
    # print('GetNetList', end_time - start_time)
    # # print(diff_dict)
    return diff_list, diff_dict, se_list, non_signal_net_list


# set0
def RunSignalTopology():
    start_time = time.clock()
    # Main Topology Extraction Function
    ###########################################
    # global layout_error_diff_list, layout_error_se_list
    # layout_error_diff_list, layout_error_se_list = list(), list()
    ##########################################

    # 抓取 dif re ns 的数据并输出到 NetList 表中
    NetTypeDetect()

    wb = Book(xlsm_path).caller()

    # 基本没用SymbolList
    ##############################################
    # 取得 SymbolList 中 User-Defined Termination {format:SCH Name-Pin}表的坐标值
    # for cell in wb.sheets("SymbolList").api.UsedRange.Cells:
    #     if cell.Value == "User-Defined Termination {format:SCH Name-Pin}":
    #         user_defined_termination_idx = (cell.Row+1, cell.Column)
    ##############################################
    # 获取 Signal Type 的值: diff or se or all
    for cell in wb.sheets("Setting").api.UsedRange.Cells:
        if cell.Value == "Signal Type:":
            signal_type_ = wb.sheets("Setting").range((cell.Row, cell.Column + 1)).value

    # global user_defined_termination_list

    # Get User_Defined_Termination_list
    # 获取 SymbolList 表中用户输入值
    # user_defined_termination_list = wb.sheets['SymbolList'].range(user_defined_termination_idx).options(expand='table', ndim=1).value
    # if user_defined_termination_list == [None]:
    #     user_defined_termination_list = list()

    # 为不同的 signal type 创建不同表
    # Get Run Signal Type
    if signal_type_ == 'Differential':
        signal_type_list = ['Differential']
        ACT_sheet_name_list = ['ACT_diff']
        ACT_count_sheet_name_list = ['ACT_diff_count']
    elif signal_type_ == 'Single-Ended':
        signal_type_list = ['Single-Ended']
        ACT_sheet_name_list = ['ACT_se']
        ACT_count_sheet_name_list = ['ACT_se_count']
    # elif signal_type_ == 'All':
    #     signal_type_list = ['Differential', 'Single-Ended']
    #     ACT_sheet_name_list = ['ACT_diff', 'ACT_se']
    #     ACT_count_sheet_name_list = ['ACT_diff_count', 'ACT_se_count']

    # Get Setting
    global allegro_report_path, layer_type_dict, start_sch_name_list, All_Layer_List

    allegro_report_path, layer_type_dict, start_sch_name_list, progress_ind, All_Layer_List = GetSetting()
    # 多余代码
    # All_Layer_List = list(set(layer_type_dict.keys()))

    setting_sheet = wb.sheets['Setting']
    setting_sheet.range(progress_ind).value = ''

    setting_sheet.range(progress_ind).value = 'Pre-Processing...'

    # 获取 dif re ns 的值
    # Detect Signal and Non-Signal Net List
    global diff_list, diff_dict, se_list, non_signal_net_list
    diff_list, diff_dict, se_list, non_signal_net_list = GetNetList()

    # # print(diff_list)
    # 获取值
    # Read Allegro Report Data
    global SCH_object_list, net_object_list, All_Net_List
    SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data, npr_brd_data = read_allegro_data(
        allegro_report_path)
    # # print(SCH_brd_data.GetData())
    SCH_object_list = SCH_detect(SCH_brd_data)
    net_object_list = net_detect(Net_brd_data, SCH_object_list)
    All_Net_List = getallnetlist(SCH_brd_data)

    # ---------------------------------------------------------------------------------------------
    ok_count = 0
    total_net_number = 0
    # Start to Extract Topology
    for idxxxx in xrange(len(signal_type_list)):
        signal_type = signal_type_list[idxxxx]
        ACT_sheet_name = ACT_sheet_name_list[idxxxx]
        ACT_count_sheet_name = ACT_count_sheet_name_list[idxxxx]

        topology_dict = dict()
        ok_check_net_list, fail_check_net_list = list(), list()
        check_sch_ok_net_dict = dict()
        check_sch_fail_net_dict = dict()

        total_check_net_list = []

        # x 在 setting 中的 Start Component Name List 中
        # 分离需check的diff 与 se信号
        for x in start_sch_name_list:
            # 和pin脚相连的线且在net_object_list中的线
            total_check_net_list.append(get_connected_net_list_by_SCH_name(x))

        if signal_type == 'Differential':
            total_check_net_list = [y for x in total_check_net_list for y in x if y in diff_list]
            # # print(total_check_net_list)
        elif signal_type == 'Single-Ended':
            total_check_net_list = [y for x in total_check_net_list for y in x if y in se_list]

        total_net_number += len(total_check_net_list)
        # # print(total_net_number)

        setting_sheet.range(progress_ind).value = 'Extracting...0/%d' % total_net_number

        # start_sch_name_list是从setting表中得出的数据，pin脚个数符合的芯片的名称
        # start_sch_name_list = ['U4']
        for check_sch_name in start_sch_name_list:
            check_net_list = get_connected_net_list_by_SCH_name(check_sch_name)
            check_net_list = [x for x in check_net_list if x in total_check_net_list]
            # check_net_list = ['J42_PRSNT2#']
            for check_net_name in check_net_list:
                try:
                    # 遍历每个器件连接的所有线名
                    # check_sch_name是符合最小pin脚个数的芯片，check_net_name是与芯片相接的信号线列表的遍历
                    # if check_net_name == 'SPI_CLK_PCH_R':
                    # print(1, check_net_name, check_sch_name)
                    # print('')
                    topology_list = topology_extract2(check_net_name, check_sch_name)
                    topology_output_list = []
                    # print('topology_list', topology_list)
                    # print('')
                    for x in topology_list:
                        if x not in topology_output_list:
                            topology_output_list.append(x)
                    # print('topology_output_list', topology_output_list)
                    # print('')
                    # print('topology_list', topology_list)
                    topology_1_list = topology_list_format_simplified(topology_output_list)
                    topology_1_output_list = []
                    # print('topology_1_list', topology_1_list)
                    for x in topology_1_list:
                        if x not in topology_1_output_list:
                            topology_1_output_list.append(x)
                    # print(check_sch_name, check_net_name)
                    # print('topology_1_output_list', topology_1_output_list)
                    # print('')
                    old_topology_list = topology_dict.get((check_sch_name, check_net_name, 'all'), [])
                    # print('old_topology_list', old_topology_list)
                    for x in topology_1_output_list:
                        if x not in old_topology_list:
                            old_topology_list += [x]
                    topology_dict[(check_sch_name, check_net_name, 'all')] = old_topology_list
                    # print('old_topology_list_out', old_topology_list)
                    # topology_dict[(check_sch_name, check_net_name, 'all')] = topology_1_output_list
                    if check_sch_ok_net_dict.get(check_sch_name) == None:
                        check_sch_ok_net_dict[check_sch_name] = [check_net_name]
                    else:
                        check_sch_ok_net_dict[check_sch_name] += [check_net_name]
                    ok_check_net_list.append(check_net_name)
                    ok_count += 1
                    setting_sheet.range(progress_ind).value = 'Extracting...%d/%d' % (ok_count, total_net_number)
                except:# Exception as e:
                    # print(check_net_name)
                    # print(e)
                    # print('')
                    if check_sch_fail_net_dict.get(check_sch_name) == None:
                        check_sch_fail_net_dict[check_sch_name] = [check_net_name]
                    else:
                        check_sch_fail_net_dict[check_sch_name] += [check_net_name]
                    fail_check_net_list.append(check_net_name)
        # # print(check_sch_ok_net_dict)
        # # print(topology_dict)
        # # print(topology_value_list)
        # # print(fail_check_net_list)
        setting_sheet.range(progress_ind).value = 'Writing to ACT Sheet...'

        # Write to ACT count Sheet
        ACT_count_sheet = wb.sheets[ACT_count_sheet_name]
        ACT_count_sheet.clear()
        row_count = 1
        ACT_count_sheet.range((row_count, 1)).value = 'ok_SCH_name'
        ACT_count_sheet.range((row_count, 2)).value = 'ok_Net_name'
        ACT_count_sheet.range((row_count, 3)).value = 'ok_Count:%d' % len(ok_check_net_list)

        for check_sch_name in start_sch_name_list:
            if check_sch_ok_net_dict.get(check_sch_name) != None:
                for check_net_name in check_sch_ok_net_dict[check_sch_name]:
                    row_count += 1
                    ACT_count_sheet.range((row_count, 1)).value = check_sch_name
                    ACT_count_sheet.range((row_count, 2)).value = check_net_name
        row_count = 1
        ACT_count_sheet.range((row_count, 5)).value = 'fail_SCH_name'
        ACT_count_sheet.range((row_count, 6)).value = 'fail_Net_name'
        ACT_count_sheet.range((row_count, 7)).value = 'fail_Count:%d' % len(fail_check_net_list)

        for check_sch_name in start_sch_name_list:
            if check_sch_fail_net_dict.get(check_sch_name) != None:
                for check_net_name in check_sch_fail_net_dict[check_sch_name]:
                    row_count += 1
                    ACT_count_sheet.range((row_count, 5)).value = check_sch_name
                    ACT_count_sheet.range((row_count, 6)).value = check_net_name

        # row_count = 1
        # # print(layout_error_diff_list)
        # # print(layout_error_se_list)

        # 整个程序找不到 layout_error_diff_list 与 layout_error_se_list， 应该是业务逻辑没有写全
        ######################################################
        # layout_error_diff_list = sorted(list(set(layout_error_diff_list)))
        # layout_error_se_list = sorted(list(set(layout_error_se_list)))
        #
        # if signal_type == 'Differential':
        #     layout_error_list = layout_error_diff_list
        # elif signal_type == 'Single-Ended':
        #     layout_error_list = layout_error_se_list
        #
        # ACT_count_sheet.range((row_count, 9)).value = 'Layout Error Net Name'
        # ACT_count_sheet.range((row_count, 10)).value = 'Error Count:%d' % len(layout_error_list)
        # for error_net_name in layout_error_list:
        #     row_count += 1
        #     ACT_count_sheet.range((row_count, 9)).value = error_net_name
        ######################################################
        # 自动调整列宽以适应文字
        ACT_count_sheet.api.UsedRange.Cells.EntireColumn.AutoFit()

        # Write to ACT Sheet
        ACT_sheet = wb.sheets[ACT_sheet_name]
        ACT_sheet.clear()

        row_count = 1
        ACT_sheet.range((row_count, 1)).value = 'ACT Results for %s Topology' % signal_type
        row_count = 2

        all_result_list = []
        # print(topology_dict)
        for content in topology_dict.values():
            for line in content:
                # if line[0] == 'CPU-C21':
                #     # print(line)
                all_result_list.append(line)
        all_result_list.sort()
        # # print(all_result_list)
        # # print(1, all_result_list)
        if all_result_list != []:
            max_len = max([len(x) for x in all_result_list])
            for idx in xrange(len(all_result_list)):
                len_tmp = len(all_result_list[idx])
                for len_diff in xrange(abs(max_len - len_tmp)):
                    # 比最长的项少的部分
                    all_result_list[idx].append(None)
            ACT_sheet.range((row_count, 1)).expand('table').value = all_result_list
        setting_sheet.range(progress_ind).value = 'Done!: success:%d, fail:%d' % (ok_count, total_net_number - ok_count)

        ACT_sheet.api.UsedRange.Cells.EntireColumn.AutoFit()
        ACT_sheet.range((row_count, 1)).current_region.api.HorizontalAlignment = Constants.xlLeft

        end_time = time.clock()

        # print('SET0', end_time - start_time)


# Detect Start & End Component
def LoadStartEndComponent():
    # Load the TX List by the specified min pin count
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    if selection_range.value.find('Topology') > -1:
        start_ind = (selection_range.row, selection_range.column)
        signal_type_ind = (start_ind[0] + 1, start_ind[1] + 1)
        # ---------------------------------------------------------------------------------------------------------------------#
        # 使用者在Topology后使用 : 进对数据初始sch进行筛选
        se_start_sch_list = []
        selected_start_sch_list = selection_range.value.split(':')[1::]

        for i in xrange(len(selected_start_sch_list)):
            selected_start_sch_half_list = str(selected_start_sch_list[i]).upper()
            se_start_sch_list.append(selected_start_sch_half_list)
        # # print(se_start_sch_list)

        # show出所写的初始芯片名
        selection_range.value = 'Topology'

        # 使用者在Signal Type后使用 : 进对数据结尾sch进行筛选
        selected_end_sch_list = active_sheet.range((start_ind[0] + 1, start_ind[1])).value.split(':')[1::]
        for i in xrange(len(selected_end_sch_list)):
            selected_end_sch_list = selected_end_sch_list[i].upper()
        active_sheet.range((start_ind[0] + 1, start_ind[1])).value = 'Signal Type'
        # ---------------------------------------------------------------------------------------------------------------------#
        if active_sheet.range(start_ind[0] + 11, start_ind[1] + 2).value == 'Net Name':
            start_sch_ind = (selection_range.row + 12, selection_range.column)
            end_sch_ind = (start_sch_ind[0], start_sch_ind[1] + 1)
            start_net_ind = (end_sch_ind[0], end_sch_ind[1] + 1)
            result_ind = (start_net_ind[0], start_net_ind[1] + 1)

        active_sheet.range(result_ind).expand('table').clear()
        net_list = active_sheet.range(start_net_ind).options(expand='table', ndim=1).value
        signal_type = active_sheet.range(signal_type_ind).value
        temp = []
        # 其实就是实现有序集合
        for net in net_list:
            if net not in temp:
                temp.append(net)
        net_list = temp

        if signal_type == 'Differential':
            act_sheet = wb.sheets['ACT_diff']

        elif signal_type == 'Single-ended':
            act_sheet = wb.sheets['ACT_se']

        output_list = list()

        act_all = act_sheet.range('A1').current_region.value

        for net in net_list:
            line_list = []
            if se_start_sch_list == [] and selected_end_sch_list == []:
                for x in act_all:
                    if x[1] == net:
                        line_list.append(x)
            elif se_start_sch_list != [] and selected_end_sch_list == []:
                for x in act_all:
                    if x[1] == net and x[0].split('-')[0] and x[0].split('-')[0] in se_start_sch_list:
                        # # print(11,se_start_sch_list)
                        line_list.append(x)
            elif se_start_sch_list == [] and selected_end_sch_list != []:
                for x in act_all:
                    if x[1] == net and x[2].split('-')[0] in selected_end_sch_list:
                        line_list.append(x)
            elif se_start_sch_list != [] and selected_end_sch_list != []:
                for x in act_all:
                    if x[1] == net and x[0].split('-')[0] in se_start_sch_list and x[2].split('-')[
                        0] in selected_end_sch_list:
                        line_list.append(x)

            if line_list != []:
                for line in line_list:
                    # 起始芯片，终止芯片，信号线
                    output_list.append((line[0], line[2], line[1]))
            else:
                output_list.append(('NA', 'NA', net))
        active_sheet.range(start_net_ind).expand('table').clear()
        active_sheet.range(start_sch_ind).expand('table').clear()

        active_sheet.range(start_sch_ind).expand('table').value = output_list

        active_sheet.api.Cells.Font.Name = 'Times New Roman'
        active_sheet.api.Cells.Font.Size = 12


def ShowResults(color_ind_list, specified_range=None):
    # Check the Present Result
    wb = Book(xlsm_path).caller()
    if specified_range is None:
        active_sheet = wb.sheets.active  # Get the active sheet object
        selection_range = wb.app.selection
    else:
        active_sheet = specified_range.sheet
        # # print(active_sheet)
        selection_range = specified_range

    if selection_range.value == 'Topology':

        start_ind = (selection_range.row, selection_range.column)
        # signal_type_ind = (start_ind[0]+1, start_ind[1]+1)

        for idx1 in xrange(len(selection_range.current_region.value)):
            line = selection_range.current_region.value[idx1]
            for idx2 in xrange(len(line)):
                if selection_range.current_region.value[idx1][idx2] == 'Net Name':
                    spec_table_length = idx1 + 1
                    break

        start_sch_ind = (start_ind[0] + spec_table_length, start_ind[1])
        end_sch_ind = (start_sch_ind[0], start_sch_ind[1] + 1)
        start_net_ind = (end_sch_ind[0], end_sch_ind[1] + 1)
        result_ind1 = (start_net_ind[0], start_net_ind[1] - 2)
        result_ind2 = (start_net_ind[0], start_net_ind[1] + 1)

        # Extract the Topology data
        topology_table = selection_range.current_region.value

        for idxx in xrange(len(topology_table)):
            row_tmp = topology_table[idxx]
            if 'Start Segment Name' in row_tmp:
                segment_list = topology_table[idxx][3::]
                segment_list = [x for x in segment_list if x not in ['', None]]
                result_idx = segment_list.index('Result')
            elif 'Min' in topology_table[idxx]:
                spec_min_list = [str(x) for x in topology_table[idxx][3::]]
                spec_min_list = spec_min_list[0:result_idx + 1]
            elif 'Max' in row_tmp:
                spec_max_list = [str(x) for x in topology_table[idxx][3::]]
                spec_max_list = spec_max_list[0:result_idx + 1]

        len1 = len(active_sheet.range(result_ind1).options(expand='table', ndim=2).value)

    SetCellFont_current_region(active_sheet, start_ind, 'Times New Roman', 12, 'l')
    SetCellBorder_current_region(active_sheet, start_ind)
    active_sheet.autofit('c')

    # Determine Pass or Fail:
    for idx1 in xrange(len1):
        result_cell_idx = (result_ind2[0] + idx1, result_ind2[1] + result_idx)
        if active_sheet.range(result_cell_idx).value != 'Something Wrong':
            active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 4
            active_sheet.range(result_cell_idx).value = 'Pass'
        elif active_sheet.range(result_cell_idx).value == 'Something Wrong':
            active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
    for idx1 in xrange(int(len1 / 2)):
        result_cell_idx1 = (result_ind2[0] + idx1 * 2, result_ind2[1] + result_idx)
        result_cell_idx2 = (result_ind2[0] + idx1 * 2 + 1, result_ind2[1] + result_idx)
        if active_sheet.range(result_cell_idx1).value != 'Something Wrong' and \
                active_sheet.range(result_cell_idx2).value != 'Something Wrong':

            if color_ind_list != []:
                if color_ind_list[idx1] == 1:
                    active_sheet.range(result_cell_idx1).value = 'Fail'
                    active_sheet.range(result_cell_idx2).value = 'Fail'
                    active_sheet.range(result_cell_idx1).api.Interior.ColorIndex = 3
                    active_sheet.range(result_cell_idx2).api.Interior.ColorIndex = 3

    for idx1 in xrange(len(segment_list)):
        for idx2 in xrange(len1):
            spec_min = spec_min_list[idx1]
            spec_max = spec_max_list[idx1]
            result_cell_idx = (result_ind2[0] + idx2, result_ind2[1] + result_idx)
            # # print('spec_min', spec_min)
            # # print('spec_max', spec_max)
            # # print('result_cell_idx', result_cell_idx)
            if active_sheet.range(result_cell_idx).value != 'Something Wrong':

                if isfloat(active_sheet.range((result_ind2[0] + idx2, result_ind2[1] + idx1)).value):
                    if spec_min == 'NA' and isfloat(spec_max):
                        if float(active_sheet.range((result_ind2[0] + idx2, result_ind2[1] + idx1)).value) <= float(
                                spec_max):
                            active_sheet.range(
                                (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 4
                        else:
                            active_sheet.range(
                                (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 3
                            active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
                            active_sheet.range(result_cell_idx).value = 'Fail'
                    elif isfloat(spec_min) and isfloat(spec_max):
                        if float(spec_min) <= float(
                                active_sheet.range((result_ind2[0] + idx2, result_ind2[1] + idx1)).value) <= float(
                            spec_max):
                            active_sheet.range(
                                (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 4
                        else:
                            active_sheet.range(
                                (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 3
                            active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
                            active_sheet.range(result_cell_idx).value = 'Fail'
                    elif isfloat(spec_min) and spec_max == 'NA':
                        if float(spec_min) <= float(
                                active_sheet.range((result_ind2[0] + idx2, result_ind2[1] + idx1)).value):
                            active_sheet.range(
                                (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 4
                        else:
                            active_sheet.range(
                                (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 3
                            active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
                            active_sheet.range(result_cell_idx).value = 'Fail'

    # Organize the table format
    table_height = len(selection_range.current_region.value)
    table_width = len(selection_range.current_region.value[0])
    result_idx = selection_range.current_region.value[2].index('Result')

    start_ind = (selection_range.row + 2, selection_range.column + result_idx + 1)
    end_ind = (
        selection_range.row + table_height - 1,
        selection_range.column + result_idx + 1 + table_width - (result_idx + 1))
    active_sheet.range(start_ind, end_ind).clear()

    start_ind = (selection_range.row, selection_range.column + 2)
    end_ind = (selection_range.row + 1, selection_range.column + table_width - 1)
    active_sheet.range(start_ind, end_ind).clear()

    end_ind = (selection_range.row + 1, selection_range.column + result_idx)
    active_sheet.range(start_ind, end_ind).api.MergeCells = True

    for idx in xrange(result_idx - 1):
        SetCellBorder(active_sheet, (start_ind[0], start_ind[1] + idx))
        SetCellBorder(active_sheet, (start_ind[0] + 1, start_ind[1] + idx))


#  清空原先生成数据表格
def ClearCheckResults(specified_range=None):
    # Clear the checked result of topology table
    wb = Book(xlsm_path).caller()
    if specified_range == None:
        active_sheet = wb.sheets.active  # Get the active sheet object
        selection_range = wb.app.selection
    else:
        active_sheet = specified_range.sheet
        # # print(active_sheet)
        selection_range = specified_range

    if selection_range.value == 'Topology':
        # Extract the Topology data
        start_ind = (selection_range.row, selection_range.column)
        # # print(start_ind)
        # signal_type_ind = (start_ind[0]+1, start_ind[1]+1)
        for idx1 in xrange(len(selection_range.current_region.value)):
            line = selection_range.current_region.value[idx1]
            for idx2 in xrange(len(line)):
                if line[idx2] == 'Net Name':
                    spec_table_length = idx1 + 1
                    break

        # # print(spec_table_length)
        start_sch_ind = (start_ind[0] + spec_table_length, start_ind[1])
        end_sch_ind = (start_sch_ind[0], start_sch_ind[1] + 1)
        start_net_ind = (end_sch_ind[0], end_sch_ind[1] + 1)
        result_ind2 = (start_net_ind[0], start_net_ind[1] + 1)

        active_sheet.range(result_ind2).expand('table').clear()


def getPreTable(num=1):
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)
    pre_ind = None
    topology_num = 0
    # haowuliaoa
    for idx in xrange(1, start_ind[0]):
        if active_sheet.range((start_ind[0] - idx, start_ind[1])).value == 'Topology':
            topology_num += 1
            if topology_num == num:
                pre_ind = (start_ind[0] - idx, start_ind[1])
                return pre_ind
            else:
                pass


def getFirstTable():
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)
    fir_ind = None

    for idx in xrange(1, start_ind[0]):
        if active_sheet.range((idx, start_ind[1])).value == 'Topology':
            fir_ind = (idx, start_ind[1])
            break
    return fir_ind


def CheckTopology(specified_range=None):
    global segment_list, trace_width_list, layer_list

    # Fill the topology table by the specified spec parameter
    def LayerMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict):
        # # print(start_sch_list)
        # # print(start_net_list)
        # # print(end_sch_list)
        # # print(act_data_dict)
        if len(start_net_list) >= 2 and len(start_net_list) % 2 == 0:
            for idx in xrange(int(len(start_net_list) / 2)):
                start_sch1, start_net1, end_sch1 = start_sch_list[2 * idx], start_net_list[2 * idx], end_sch_list[
                    2 * idx]
                start_sch2, start_net2, end_sch2 = start_sch_list[2 * idx + 1], start_net_list[2 * idx + 1], \
                                                   end_sch_list[2 * idx + 1]

                act_data1 = act_data_dict[(start_sch1, start_net1, end_sch1)]
                act_data2 = act_data_dict[(start_sch2, start_net2, end_sch2)]

                layer_len_list1 = list()
                for idx1 in xrange(len(act_data1)):
                    seg_1 = act_data1[idx1]
                    if str(seg_1).find(':') > -1:
                        for x in act_data1[idx1].split(':'):
                            if x in All_Layer_List:
                                layer1 = x
                        len1 = act_data1[idx1 + 1]
                        layer_len_list1.append((layer1, len1))
                layer_len_list2 = list()
                for idx2 in xrange(len(act_data2)):
                    if str(act_data2[idx2]).find(':') > -1:
                        for x in act_data2[idx2].split(':'):
                            if x in All_Layer_List:
                                layer2 = x
                        len2 = act_data2[idx2 + 1]
                        layer_len_list2.append((layer2, len2))

                end_tmp = False
                while end_tmp is False:
                    end_tmp = True
                    for idx1 in xrange(len(layer_len_list1) - 1):
                        if layer_len_list1[idx1][0] == layer_len_list1[idx1 + 1][0]:
                            layer_len_list1[idx1] = (layer_len_list1[idx1][0], float(layer_len_list1[idx1][1]) + float(
                                layer_len_list1[idx1 + 1][1]))
                            layer_len_list1.pop(idx1 + 1)
                            end_tmp = False
                            break
                end_tmp = False
                while end_tmp is False:
                    end_tmp = True
                    for idx2 in xrange(len(layer_len_list2) - 1):
                        if layer_len_list2[idx2][0] == layer_len_list2[idx2 + 1][0]:
                            layer_len_list2[idx2] = (layer_len_list2[idx2][0], float(layer_len_list2[idx2][1]) + float(
                                layer_len_list2[idx2 + 1][1]))
                            layer_len_list2.pop(idx2 + 1)
                            end_tmp = False
                            break

                max_mismatch = 0
                if len(layer_len_list1) == len(layer_len_list2):
                    for seg_idx in xrange(len(layer_len_list1)):
                        if layer_len_list1[seg_idx][0] == layer_len_list2[seg_idx][0]:
                            if abs(layer_len_list1[seg_idx][1] - layer_len_list2[seg_idx][1]) > max_mismatch:
                                max_mismatch = abs(layer_len_list1[seg_idx][1] - layer_len_list2[seg_idx][1])
                else:
                    max_mismatch = -1

                act_data_dict[((start_sch1, start_net1, end_sch1), 'layer_mismatch')] = max_mismatch
                act_data_dict[((start_sch2, start_net2, end_sch2), 'layer_mismatch')] = max_mismatch
        else:
            for idx in xrange(len(start_net_list)):
                act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'layer_mismatch')] = 'NA'

        return act_data_dict

    def SegmentMismatch(start_sch_list, start_net_list, end_sch_list, result_dict, all_width_list, all_layer_list,
                        all_result_list, segment_out_name_list):
        # # print('all_width_list', all_width_list)
        # # print('all_layer_list', all_layer_list)
        # # print('all_result_list', all_result_list)

        # # print('segment_out_name_list', segment_out_name_list)
        if result_dict:
            if len(start_net_list) >= 2 and len(start_net_list) % 2 == 0:

                for idx in xrange(int(len(start_net_list) / 2)):
                    start_sch1, start_net1, end_sch1 = start_sch_list[2 * idx], start_net_list[2 * idx], end_sch_list[
                        2 * idx]
                    start_sch2, start_net2, end_sch2 = start_sch_list[2 * idx + 1], start_net_list[2 * idx + 1], \
                                                       end_sch_list[2 * idx + 1]

                    result_list1 = result_dict[(start_sch1, start_net1, end_sch1)]
                    result_list2 = result_dict[(start_sch2, start_net2, end_sch2)]

                    # # print(1, result_list1)
                    # # print(2, result_list2)
                    # # print(3, segment_list)
                    # 找到segment位置
                    com_del_ind_list = []
                    for se_ind in xrange(len(segment_list)):
                        if segment_list[se_ind] == 'Segment Mismatch':
                            seg_table_ind = se_ind
                        if segment_list[se_ind] == 'Total Length':
                            seg_name_ind = se_ind

                    # 找出segment_name
                    segment_name_list = list(segment_list[:seg_name_ind])
                    # # print(segment_name_list)
                    # # print('segment_name_list', segment_name_list)
                    all_width_list_1 = all_width_list[2 * idx]
                    all_width_list_2 = all_width_list[2 * idx + 1]
                    all_layer_list_1 = all_layer_list[2 * idx]
                    all_layer_list_2 = all_layer_list[2 * idx + 1]
                    all_result_list_1 = all_result_list[2 * idx]
                    all_result_list_2 = all_result_list[2 * idx + 1]

                    def addComponent(param):
                        for x in xrange(len(param)):
                            if str(param[x]).find('net$') > -1:
                                param.insert(x + 1, 'Component Name')

                    # def delNet(param):
                    #     param = [param[x] for x in xrange(len(param)) if str(param[x]).find('net$') == -1]
                    #     return param
                    addComponent(all_width_list_1)
                    addComponent(all_width_list_2)
                    addComponent(all_layer_list_1)
                    addComponent(all_layer_list_2)
                    addComponent(all_result_list_1)
                    addComponent(all_result_list_2)
                    # # print('all_width_list_1', all_width_list_1)
                    # # print('all_width_list_2', all_width_list_2)
                    # # print('all_layer_list_1', all_layer_list_1)
                    # # print('all_layer_list_2', all_layer_list_2)
                    # # print('all_result_list_1', all_result_list_1)
                    # # print('all_result_list_2', all_result_list_2)

                    # 找出segement的值
                    for x in xrange(len(all_width_list_1)):
                        if all_width_list_1[x].find('net$') > -1:
                            segment_out_name_list.insert(x, '0')
                            segment_out_name_list.insert(x + 1, '0')
                    # # print(segment_out_name_list)

                    count = 0
                    seg_name_dict = {}
                    segment_out_name_number_list = []
                    for x in xrange(len(segment_out_name_list)):
                        segment_value_list = []
                        for y in xrange(len(segment_out_name_list[x])):
                            segment_out_name_number_list.append(count)
                            segment_value_list.append(count)
                            count += 1
                        if len(segment_value_list) > 1:
                            for z in xrange(1, len(segment_value_list) + 1):
                                seg_name_dict[segment_value_list[z - 1]] = segment_name_list[x] + '_' + str(z)
                        else:
                            seg_name_dict[segment_value_list[0]] = segment_name_list[x]
                    # # print(segment_out_name_number_list)
                    # # print(seg_name_dict)
                    # 判断是否是segment
                    seg_com_ind_list = []
                    seg_wid_ind_list = []
                    seg_lay_ind_list = []

                    for wid_ind in xrange(len(all_width_list_1) - 1):
                        if all_width_list_1[wid_ind] == 'Component Name':
                            seg_com_ind_list.append(wid_ind)

                        if all_width_list_1[wid_ind] != 'Component Name' and all_width_list_1[wid_ind + 1] \
                                != 'Component Name' and all_width_list_1[wid_ind].find('net$') == -1 and \
                                all_width_list_1[wid_ind + 1].find('net$') == -1:
                            if all_width_list_1[wid_ind] != all_width_list_1[wid_ind + 1]:
                                seg_wid_ind_list.append(wid_ind)

                    for layer_ind in xrange(len(all_layer_list_1) - 1):
                        if all_layer_list_1[layer_ind] != 'Component Name' and all_layer_list_1[layer_ind + 1] \
                                != 'Component Name' and all_layer_list_1[layer_ind].find('net$') == -1 and \
                                all_layer_list_1[layer_ind + 1].find('net$') == -1:
                            if all_layer_list_1[layer_ind] != all_layer_list_1[layer_ind + 1]:
                                seg_lay_ind_list.append(layer_ind)

                    # # print(6, seg_com_ind_list)
                    # # print(7, seg_wid_ind_list)
                    # # print(8, seg_lay_ind_list)

                    seg_name_list = []
                    seg_value_list = []
                    # 计算segment mismatch
                    for ind1 in seg_wid_ind_list:
                        seg_name_list.append(seg_name_dict.get(ind1))
                        seg_name_list.append(seg_name_dict.get(ind1 + 1))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind1]) - float(all_result_list_2[ind1])), 2))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind1 + 1]) - float(all_result_list_2[ind1 + 1])), 2))
                    for ind2 in seg_com_ind_list:
                        seg_name_list.append(seg_name_dict.get(ind2 - 2))
                        seg_name_list.append(seg_name_dict.get(ind2 + 1))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind2 - 2]) - float(all_result_list_2[ind2 - 2])), 2))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind2 + 1]) - float(all_result_list_2[ind2 + 1])), 2))
                    for ind3 in seg_lay_ind_list:
                        seg_name_list.append(seg_name_dict.get(ind3))
                        seg_name_list.append(seg_name_dict.get(ind3 + 1))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind3]) - float(all_result_list_2[ind3])), 2))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind3 + 1]) - float(all_result_list_2[ind3 + 1])), 2))

                    seg_name_out_list = []
                    seg_del_ind_list = []
                    seg_value_out_list = []

                    for x in xrange(len(seg_name_list)):
                        if seg_name_list[x] not in seg_name_out_list:
                            seg_name_out_list.append(seg_name_list[x])
                        else:
                            seg_del_ind_list.append(x)

                    seg_value_out_list = [seg_value_list[x] for x in xrange(len(seg_name_list)) if
                                          x not in seg_del_ind_list]

                    seg_out_ind = []
                    seg_max_ind = None

                    # # print('seg_name_list', seg_name_list)
                    # # print('seg_value_list', seg_value_list)
                    ######################################################
                    # 获得管控值
                    max_seg = spec_max_list[seg_table_ind]
                    min_seg = spec_min_list[seg_table_ind]

                    # 检测有没有超过管控
                    for seg_ind in xrange(len(seg_value_out_list)):
                        # # print(seg_ind, seg_value_list[seg_ind])
                        if seg_value_out_list[seg_ind] >= float(max_seg) or seg_value_out_list[seg_ind] < float(
                                min_seg):
                            seg_out_ind.append(seg_ind)
                    if seg_out_ind == []:
                        seg_max_ind = seg_value_out_list.index(max(seg_value_out_list))

                    # 输出数值并改变颜色
                    seg_out_value = ''

                    # # print(seg_out_ind, seg_max_ind)

                    if seg_out_ind:
                        color_ind_list.append(1)
                        for x in seg_out_ind:
                            seg_out_value += seg_name_out_list[x] + ':' + str(seg_value_out_list[x]) + '; '

                    # # print(seg_value_list)
                    # # print(seg_max_ind)
                    if seg_max_ind != None:
                        # # print('ok')
                        color_ind_list.append(0)
                        seg_out_value = seg_value_list[seg_max_ind]
                    # # print(seg_out_value)
                    result_list1[seg_table_ind] = seg_out_value
                    result_list2[seg_table_ind] = seg_out_value

                    result_dict[(start_sch1, start_net1, end_sch1)] = result_list1
                    result_dict[(start_sch2, start_net2, end_sch2)] = result_list2
            # # print(color_ind_list)
            return result_dict, color_ind_list
        else:
            for idx in xrange(len(start_net_list)):
                act_data_dict[((start_sch_list[idx], start_net_list[idx],
                                end_sch_list[idx]), 'segment_mismatch')] = 'NA'
            # # print(act_data_dict)
            # # print('color_ind_list', color_ind_list)
            # # print(act_data_dict)
            return act_data_dict

    def TotalMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict):
        if len(start_net_list) >= 2 and len(start_net_list) % 2 == 0:
            for idx in xrange(int(len(start_net_list) / 2)):
                start_sch1, start_net1, end_sch1 = start_sch_list[2 * idx], start_net_list[2 * idx], end_sch_list[
                    2 * idx]
                start_sch2, start_net2, end_sch2 = start_sch_list[2 * idx + 1], start_net_list[2 * idx + 1], \
                                                   end_sch_list[2 * idx + 1]

                total_length1 = act_data_dict[((start_sch1, start_net1, end_sch1), 'total_length')]
                total_length2 = act_data_dict[((start_sch2, start_net2, end_sch2), 'total_length')]

                len1 = float(total_length1)
                len2 = float(total_length2)

                act_data_dict[((start_sch1, start_net1, end_sch1), 'total_mismatch')] = abs(len1 - len2)
                act_data_dict[((start_sch2, start_net2, end_sch2), 'total_mismatch')] = abs(len1 - len2)
        else:
            for idx in xrange(len(start_net_list)):
                act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'total_mismatch')] = 'NA'

        return act_data_dict

    def Skew2BundleMismatch(unit_bundle_number, start_sch_list, start_net_list, end_sch_list, act_data_dict):

        if len(start_net_list) >= unit_bundle_number and len(start_net_list) % unit_bundle_number == 0:
            for idx1 in xrange(len(start_net_list)):
                if idx1 % unit_bundle_number == 0:
                    len_list_tmp = list()
                    bundle_max_mismatch = 0
                    for idx2 in xrange(unit_bundle_number):
                        len_tmp_ = float(act_data_dict[(
                            (start_sch_list[idx1 + idx2], start_net_list[idx1 + idx2], end_sch_list[idx1 + idx2]),
                            'total_length')])
                        len_list_tmp.append(len_tmp_)
                    bundle_max_mismatch = max(len_list_tmp) - min(len_list_tmp)

                    for idx2 in xrange(unit_bundle_number):
                        act_data_dict[(
                            (start_sch_list[idx1 + idx2], start_net_list[idx1 + idx2], end_sch_list[idx1 + idx2]),
                            'bundle_mismatch')] = bundle_max_mismatch
        else:
            for idx1 in xrange(len(start_net_list)):
                act_data_dict[
                    ((start_sch_list[idx1], start_net_list[idx1], end_sch_list[idx1]), 'bundle_mismatch')] = 'NA'

        return act_data_dict

    def Skew2TargetMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict, target_list_all, act_content,
                            key_title='skew2target'):

        if target_list_all != []:
            for idx_t in xrange(len(target_list_all)):
                target_list = target_list_all[idx_t]
                if key_title.find('skew2target') > -1:
                    key_title = 'skew2target%d' % idx_t

                act_content_se = wb.sheets['ACT_se'].range('A1').current_region.value
                act_content_diff = wb.sheets['ACT_diff'].range('A1').current_region.value

                if act_content_diff != None and act_content_se != None:
                    act_content_all = act_content_se + act_content_diff
                elif act_content_diff != None and act_content_se == None:
                    act_content_all = act_content_diff
                elif act_content_diff == None and act_content_se != None:
                    act_content_all = act_content_se
                else:
                    act_content_all = []

                for start_sch, end_sch, start_net in target_list_all[idx_t]:
                    for line in act_content_all:
                        if start_sch == line[0] and end_sch == line[2] and start_net == line[1]:
                            if act_data_dict.get((start_sch, start_net, end_sch)) == None:
                                act_data_dict[((start_sch, start_net, end_sch), 'via')] = line[3]
                                act_data_dict[((start_sch, start_net, end_sch), 'total_length')] = line[4].split()[1]

                                act_data_dict[(start_sch, start_net, end_sch)] = [x for x in line[5::] if
                                                                                  x not in ['', None]]
                                break

                target_len_list = list()

                for target in target_list:
                    start_sch, end_sch, start_net = target[0], target[1], target[2]

                    target_len = float(act_data_dict[((start_sch, start_net, end_sch), 'total_length')])

                    target_len_list.append(target_len)

                for i22 in xrange(len(start_sch_list)):
                    item_len = float(
                        act_data_dict[((start_sch_list[i22], start_net_list[i22], end_sch_list[i22]), 'total_length')])
                    act_data_dict[((start_sch_list[i22], start_net_list[i22], end_sch_list[i22]), key_title)] = max(
                        [abs(item_len - x) for x in target_len_list])
        else:
            for i33 in xrange(len(start_sch_list)):
                act_data_dict[((start_sch_list[i33], start_net_list[i33], end_sch_list[i33]), key_title)] = 'NA'

        return act_data_dict

    def GroupMismatch(topology_name_list_all, start_sch_list, start_net_list, end_sch_list, act_data_dict, act_content):

        wb = Book(xlsm_path).caller()
        active_sheet = wb.sheets.active  # Get the active sheet object

        for idx_ in xrange(len(topology_name_list_all)):
            data_list = list()
            for topology_name in topology_name_list_all[idx_]:
                for cell in active_sheet.api.UsedRange.Cells:
                    if cell.Value == 'Topology' and active_sheet.range(
                            (cell.Row, cell.Column + 1)).value == topology_name:
                        data_list.append(active_sheet.range((cell.Row, cell.Column)).current_region.value)

            target_list = list()
            for data in data_list:
                for x in data:
                    if x[2] == 'Net Name':
                        data_idx = data.index(x) + 1
                target_list += [(x[0], x[1], x[2].split('!')[0]) for x in data[data_idx::]]

            target_list_all = [target_list]
            act_data_dict = Skew2TargetMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict,
                                                target_list_all, act_content, key_title='group_mismatch%d' % idx_)

        return act_data_dict

    def DQSDLLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict, DQS_flag=True):

        for idx in xrange(int(len(start_net_list) / 2)):
            start_sch1, start_net1, end_sch1 = start_sch_list[2 * idx], start_net_list[2 * idx], end_sch_list[2 * idx]
            start_sch2, start_net2, end_sch2 = start_sch_list[2 * idx + 1], start_net_list[2 * idx + 1], end_sch_list[
                2 * idx + 1]

            act_data1 = act_data_dict[(start_sch1, start_net1, end_sch1), 'total_length']
            act_data2 = act_data_dict[(start_sch2, start_net2, end_sch2), 'total_length']
            if DQS_flag:
                act_data_dict[(start_sch1, start_net1, end_sch1), 'DQS_TO_DQS'] = \
                    round(abs(float(act_data1) - float(act_data2)), 2)
                act_data_dict[(start_sch2, start_net2, end_sch2), 'DQS_TO_DQS'] = \
                    round(abs(float(act_data1) - float(act_data2)), 2)
            else:
                act_data_dict[(start_sch1, start_net1, end_sch1), 'DLL_Group'] = \
                    round(abs(float(act_data1) - float(act_data2)), 2)
                act_data_dict[(start_sch2, start_net2, end_sch2), 'DLL_Group'] = \
                    round(abs(float(act_data1) - float(act_data2)), 2)

        return act_data_dict

        # # print(act_data_dict)

    def DQTODQMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict):

        DQ_value_list = [float(act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]),
                                              'total_length')]) for idx in xrange(len(start_net_list))]

        # dq_max = max(DQ_value_list)
        # dq_min = min(DQ_value_list)
        # DQ_real_list = []
        # for i in xrange(len(DQ_value_list)):
        #     lst = round(dq_max-DQ_value_list[i], 2)
        #     lst1 = round(DQ_value_list[i] - dq_min, 2)
        #     DQ_real_list.append(max(lst, lst1))
        # for idx in xrange(len(start_net_list)):
        # act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'DQ_TO_DQS')] = DQ_real_list[idx]

        for idx in xrange(len(start_net_list)):
            act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'DQ_TO_DQ')] = \
                round(max(abs(max(DQ_value_list) - DQ_value_list[idx]), abs(DQ_value_list[idx] -
                                                                            min(DQ_value_list))), 2)

        return act_data_dict

    def DQTODQSMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict):

        DQ_value_list = [float(act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]),
                                              'total_length')]) for idx in xrange(len(start_net_list))]

        DQS_table_ind = getPreTable()
        DQS_value_list = []

        if DQS_table_ind:
            topology_table = active_sheet.range(DQS_table_ind).current_region.value
            for ind in xrange(len(topology_table[2])):
                if topology_table[2][ind] == 'Total Length':
                    len_ind = ind
            # 获取DQS长度数据
            for value in topology_table[12:]:
                DQS_value_list.append(value[len_ind])

        # dqs_max = max(DQS_value_list)
        # dqs_min = min(DQS_value_list)
        # DQS_real_list = []
        # for i in xrange(len(DQ_value_list)):
        #     lst = round(dqs_max-DQ_value_list[i],2)
        #     lst1 = round(DQ_value_list[i]-dqs_min,2)
        #     DQS_real_list.append(max(lst,lst1))
        # for idx in xrange(len(start_net_list)):
        #     act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'DQ_TO_DQS')] = DQS_real_list[idx]
        for idx in xrange(len(start_net_list)):
            act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'DQ_TO_DQS')] = \
                round(max(abs(max(DQS_value_list) - DQ_value_list[idx]), abs(DQ_value_list[idx] -
                                                                             min(DQS_value_list))), 2)
        return act_data_dict

    def CMDCTLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict, CMD_flag=True):

        # # print(start_sch_list, start_net_list, end_sch_list, act_data_dict, CMD_flag)
        CMD_value_list = [float(act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]),
                                               'total_length')]) for idx in xrange(len(start_net_list))]
        # # print(1, CMD_value_list)
        CLK_table_ind = getFirstTable()
        CLK_value_list = []
        # # print(2, CLK_table_ind)
        if CLK_table_ind:
            topology_table = active_sheet.range(CLK_table_ind).current_region.value
            for ind in xrange(len(topology_table[2])):
                if topology_table[2][ind] == 'Total Length':
                    len_ind = ind
            # 获取DQS长度数据
            for value in topology_table[12:]:
                CLK_value_list.append(value[len_ind])

        if CMD_flag:
            for idx in xrange(len(start_net_list)):
                act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'CMD_TO_CLK')] = \
                    round(max(abs(max(CLK_value_list) - CMD_value_list[idx]),
                              abs(CMD_value_list[idx] - min(CLK_value_list))), 2)
        else:
            for idx in xrange(len(start_net_list)):
                act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'CTL_TO_CLK')] = \
                    round(max(abs(max(CLK_value_list) - CMD_value_list[idx]), abs(CMD_value_list[idx] -
                                                                                  min(CLK_value_list))), 2)
        return act_data_dict

    # def ForAMD_LOther(start_sch_list, start_net_list, end_sch_list, act_data_dict, CAP_type = None, Length_Type = 'Center'):
    #
    #     if CAP_type == None:
    #         dict_kw = 'AMD_LOther_CAP'
    #     elif CAP_type == 'PTH':
    #         dict_kw = 'AMD_LOther_Via'
    #
    #
    #     global SCH_object_list
    #     def GetConnectedCAPList(topology_list, CAP_type):
    #         tmp_list = []
    #         for seg in topology_list:
    #             if re.findall('\[(.*?)\]', str(seg)):
    #                 tmp_list.append(seg)
    #
    #         connected_sch_list_ = []
    #         for seg in tmp_list:
    #             for sch in re.findall('\[(.*?)\]', seg):
    #                 connected_sch_list_.append(sch.split('-'))
    #
    #         connected_cap_list = list()
    #         for sch, pin in connected_sch_list_:
    #             sch_obj = get_SCH_object_by_name(sch)
    #             if re.findall('^CAP_', sch_obj.GetModel()) != []:
    #                 connected_cap_list.append((sch, pin))
    #
    #         return connected_cap_list
    #     def TwoPointDistance(xy1, xy2):
    #         return ((float(xy1[0])-float(xy2[0]))**2 + (float(xy1[1])-float(xy2[1]))**2)**0.5
    #
    #     diff_list, diff_dict, se_list, non_signal_net_list = GetNetList()
    #     allegro_report_path, layer_type_dict, start_sch_name_list, progress_ind, All_Layer_List = GetSetting()
    #     SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data = read_allegro_data(allegro_report_path)
    #     SCH_object_list = SCH_detect(SCH_brd_data)
    #
    #     if len(start_net_list) >= 2 and len(start_net_list)%2 == 0:
    #         for idx in xrange(len(start_net_list)):
    #             check_net = start_net_list[idx]
    #             topology_list_check = act_data_dict[(start_sch_list[idx], start_net_list[idx], end_sch_list[idx])]
    #             connected_cap_list_check = GetConnectedCAPList(topology_list_check, CAP_type)
    #
    #             start_net_idx_list_target = []
    #             for idx_ in xrange(len(start_net_list)):
    #                 if start_net_list[idx_] not in [check_net, diff_dict[check_net]]:
    #                     start_net_idx_list_target.append(idx_)
    #
    #             start_net_list_target = []
    #             start_sch_list_target = []
    #             end_sch_list_target = []
    #             for idx_ in start_net_idx_list_target:
    #                 start_net_list_target.append(start_net_list[idx_])
    #                 start_sch_list_target.append(start_sch_list[idx_])
    #                 end_sch_list_target.append(end_sch_list[idx_])
    #
    #
    #             connected_cap_list_all_target = list()
    #             for i12 in xrange(len(start_sch_list_target)):
    #                 topology_list = act_data_dict[(start_sch_list_target[i12], start_net_list_target[i12], end_sch_list_target[i12])]
    #                 connected_cap_list = GetConnectedCAPList(topology_list, CAP_type)
    #                 connected_cap_list_all_target += connected_cap_list
    #
    #             dis_list = list()
    #
    #             if Length_Type == 'Center':
    #                 for sch, pin in connected_cap_list_check:
    #                     cap_obj1 = get_SCH_object_by_name(sch)
    #                     if len(cap_obj1.GetPinList()) == 2:
    #                         xy_list = []
    #                         for pin in cap_obj1.GetPinList():
    #                             xy_list.append(cap_obj1.GetXY(pin))
    #                         xy_center1 = ((float(xy_list[0][0])+float(xy_list[1][0]))/2, (float(xy_list[0][1])+float(xy_list[1][1]))/2)
    #                         for sch_, pin_ in connected_cap_list_all_target:
    #                             cap_obj2 = get_SCH_object_by_name(sch_)
    #                             if len(cap_obj2.GetPinList()) == 2:
    #                                 xy_list = []
    #                                 for pin in cap_obj2.GetPinList():
    #                                     xy_list.append(cap_obj2.GetXY(pin))
    #                                 xy_center2 = ((float(xy_list[0][0])+float(xy_list[1][0]))/2, (float(xy_list[0][1])+float(xy_list[1][1]))/2)
    #
    #                                 dis_list.append(TwoPointDistance(xy_center1, xy_center2))
    #             elif Length_Type == 'Pin2Pin':
    #                 for sch, pin in connected_cap_list_check:
    #                     cap_obj1 = get_SCH_object_by_name(sch)
    #                     if len(cap_obj1.GetPinList()) == 2:
    #                         xy_list1 = []
    #                         for pin in cap_obj1.GetPinList():
    #                             xy_list1.append(cap_obj1.GetXY(pin))
    #                         for sch_, pin_ in connected_cap_list_all_target:
    #                             cap_obj2 = get_SCH_object_by_name(sch_)
    #                             if len(cap_obj2.GetPinList()) == 2:
    #                                 xy_list2 = []
    #                                 for pin in cap_obj2.GetPinList():
    #                                     xy_list2.append(cap_obj2.GetXY(pin))
    #                                 dis_list += [TwoPointDistance(xy1, xy2) for xy1 in xy_list1 for xy2 in xy_list2]
    #
    #             if dis_list == []:
    #                 act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]),dict_kw)] = 'NA'
    #             else:
    #                 act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]),dict_kw)] = '%.3f'%min(dis_list)
    #
    #
    #     else:
    #         for idx in xrange(len(start_net_list)):
    #             act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]),dict_kw)] = 'NA'
    #
    #     return act_data_dict
    ########################  Run
    # 清空原先生成数据表格

    ClearCheckResults(specified_range=specified_range)

    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)
    topology_table = active_sheet.range(start_ind).current_region.value

    color_ind_list = []

    # for DDR 按钮
    # 如果是特定的DDR表格则指定输入特定位置的值
    if active_sheet.range(start_ind[0], start_ind[1] + 1).value == 'For DDR':

        topology_width_table = topology_table[7][3::]

        topology_width_table = [x for x in topology_width_table if x not in ['NA', 'N']]

        if topology_width_table[-1] in [4.3, 6.75]:
            # MS
            active_sheet.range(start_ind[0] + 3, start_ind[1] + 2 + len(topology_width_table)).value = 'MS'
            active_sheet.range(start_ind[0] + 9, start_ind[1] + 3 + len(topology_width_table)).value = 800
            active_sheet.range(start_ind[0] + 10, start_ind[1] + 3 + len(topology_width_table)).value = 3200
        elif topology_width_table[-1] in [3.5, 5.7, 4, 6]:
            # SL
            active_sheet.range(start_ind[0] + 3, start_ind[1] + 2 + len(topology_width_table)).value = 'SL'
            active_sheet.range(start_ind[0] + 9, start_ind[1] + 3 + len(topology_width_table)).value = 800
            active_sheet.range(start_ind[0] + 10, start_ind[1] + 3 + len(topology_width_table)).value = 4700
        else:
            pass

        active_sheet.range(start_ind[0], start_ind[1] + 1).value = ''

    if specified_range == None:
        active_sheet = wb.sheets.active  # Get the active sheet object
        selection_range = wb.app.selection
    else:
        active_sheet = specified_range.sheet
        selection_range = specified_range

    allegro_report_path, layer_type_dict, start_sch_name_list, progress_ind, All_Layer_List = GetSetting()
    All_Layer_List = list(set(layer_type_dict.keys()))
    # # print(layer_type_dict)

    # 获取每层的厚度以及总的厚度
    SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data, npr_brd_data = read_allegro_data(
        allegro_report_path)
    layerLength_list = []
    data = stackup_brd_data.GetData()
    data = data[1:-3]
    # # print(data)
    # lst = []
    for idx in range(len(data)):
        data[idx] = data[idx].split(',')
        layerLength_list.append({data[idx][0]: data[idx][3]})  # 存储所有层的厚度

    # 找到topology的表格
    if selection_range.value == 'Topology':

        # 计算Segment的行数
        for se_ind in xrange(len(topology_table[2])):
            # 获取Segment Mismatch的管控条件
            if topology_table[2][se_ind] and topology_table[2][se_ind].find('Segment Mismatch') > -1:
                se1_ind = se_ind

        start_ind = (selection_range.row, selection_range.column)

        for idx1 in xrange(len(selection_range.current_region.value)):
            for idx2 in xrange(len(selection_range.current_region.value[idx1])):
                if selection_range.current_region.value[idx1][idx2] == 'Net Name':
                    spec_table_length = idx1 + 1
                    break
        # 获取坐标
        start_sch_ind = (start_ind[0] + spec_table_length, start_ind[1])
        end_sch_ind = (start_sch_ind[0], start_sch_ind[1] + 1)
        start_net_ind = (end_sch_ind[0], end_sch_ind[1] + 1)
        # result_ind1 不是和 start_sch_ind 重合了吗
        result_ind1 = (start_net_ind[0], start_net_ind[1] - 2)
        result_ind2 = (start_net_ind[0], start_net_ind[1] + 1)
        # # print(start_sch_ind, result_ind1)

        # Extract the Topology data
        # 获取topology表格数据，没有数据的用None表示
        topology_table = selection_range.current_region.value

        # 分类存放表格数据
        for idxx in xrange(len(topology_table)):
            row_tmp = topology_table[idxx]
            if 'Start Segment Name' in row_tmp:
                segment_list = topology_table[idxx][3::]
                segment_list = [x for x in segment_list if x not in ['', None]]
                # # print(segment_list)
                # 从B0开始为0，一直到Result的索引
                result_idx = segment_list.index('Result')
                # # print(result_idx)
            elif 'Layer' in topology_table[idxx]:
                layer_list = []
                for x in topology_table[idxx][3::]:
                    layer_list.append(str(x))
                layer_list = layer_list[0:result_idx + 1]
                layer_list = [x.split('/') for x in layer_list]
                # # print(layer_list)
            elif 'Layer Change' in topology_table[idxx]:
                layer_change_list = []
                for x in topology_table[idxx][3::]:
                    layer_change_list.append(str(x))
                layer_change_list = layer_change_list[0:result_idx + 1]
            elif 'Connect Component' in row_tmp:
                connect_sch_list = []
                for x in topology_table[idxx][3::]:
                    connect_sch_list.append(str(x))
                connect_sch_list = connect_sch_list[0:result_idx + 1]
            elif 'Cross Over' in row_tmp:
                cross_over_list = []
                for x in topology_table[idxx][3::]:
                    cross_over_list.append(str(x))
                cross_over_list = cross_over_list[0:result_idx + 1]
            elif 'Trace Width' in topology_table[idxx]:
                trace_width_list = []
                for x in topology_table[idxx][3::]:
                    trace_width_list.append(str(x))
                trace_width_list = trace_width_list[0:result_idx + 1]
                for idx in xrange(len(trace_width_list)):
                    # 得到的数都为浮点数
                    tw = trace_width_list[idx]
                    # 数据为浮点数时，如3.5，或者有 / 时（差分信号）
                    if isfloat(trace_width_list[idx]) or tw.find('/') > -1:
                        trace_width_list[idx] = ['%.3f' % float(x) for x in trace_width_list[idx].split('/')]
            elif 'Space' in row_tmp:
                space_list = []
                for x in topology_table[idxx][3::]:
                    space_list.append(str(x))
                space_list = space_list[0:result_idx + 1]
            elif 'Min' in row_tmp:
                spec_min_list = []
                for x in topology_table[idxx][3::]:
                    spec_min_list.append(str(x))
                spec_min_list = spec_min_list[0:result_idx + 1]
            elif 'Max' in row_tmp:
                spec_max_list = []
                for x in topology_table[idxx][3::]:
                    spec_max_list.append(str(x))
                spec_max_list = spec_max_list[0:result_idx + 1]
        print('connect_sch_list', connect_sch_list)
        print('segment_list', segment_list)
        print('trace_width_list', trace_width_list)
        print('layer_change_list', layer_change_list)
        # 获取表格类型：differential or single
        signal_type = topology_table[1][1]
        # 数据已经被清空过了，此为多余代码
        # active_sheet.range(result_ind2).expand('table').clear()
        # 包含 Start Component Name-Pin Number，End Component Name-Pin Number,Net Name三个值的表格值
        table1 = active_sheet.range(result_ind1).expand('table').value

        # 无用代码
        # 最后判断为有用代码
        ##################################
        if type(table1[0]) == type(u''):
            table1 = [table1]
        ##################################

        start_sch_list = [tt[0] for tt in table1]
        end_sch_list = [tt[1] for tt in table1]
        start_net_list = [tt[2] for tt in table1]

        # Check ignore syntax to skip specified segment check
        ignore_segment_dict = dict()
        start_net_list = list(start_net_list)
        for idx in xrange(len(start_net_list)):
            net_name = start_net_list[idx]
            # !代表什么
            if net_name.find('!') > -1:
                ignore_segment_list = start_net_list[idx].split('!')[1::]
                start_net_list[idx] = start_net_list[idx].split('!')[0]
                ignore_segment_dict[(start_sch_list[idx], start_net_list[idx], end_sch_list[idx])] = ignore_segment_list
            else:
                ignore_segment_dict[(start_sch_list[idx], start_net_list[idx], end_sch_list[idx])] = []
        start_net_list = tuple(start_net_list)

        # # print(1, start_net_list)
        layer_mismatch, segment_mismatch, bundle_mismatch, total_mismatch, DQS_TO_DQS_mismatch \
            = False, False, False, False, False
        DQ_TO_DQ_mismatch, DQ_TO_DQS_mismatch, CMD_mismatch, CTL_mismatch, DLL_mismatch \
            = False, False, False, False, False
        # amd_lother_cap, amd_lother_via = False, False

        # Extract the ACT data
        if signal_type == 'Differential':
            act_sheet = wb.sheets['ACT_diff']
            if 'Layer Mismatch' in segment_list:
                layer_mismatch = True
            if 'Segment Mismatch' in segment_list:
                segment_mismatch = True
            if 'Skew to Bundle' in segment_list:
                bundle_mismatch = True
            if 'Total Mismatch' in segment_list:
                total_mismatch = True
            # 下面两个是什么,暂时没有用到
            # 少了 skew to Target 与 group mismatch
            ##############################
            # if 'AMD_LOther_CAP' in segment_list:
            #     amd_lother_cap = True
            # if 'AMD_LOther_Via' in segment_list:
            #     amd_lother_via = True
            ##############################

            if 'Relative Length Spec(DQS to DQS)' in segment_list:
                DQS_TO_DQS_mismatch = True
        elif signal_type == 'Single-ended':
            act_sheet = wb.sheets['ACT_se']
            if 'Relative Length Spec(DQ to DQ)' in segment_list:
                DQ_TO_DQ_mismatch = True
            if 'Relative Length Spec(DQ to DQS)' in segment_list:
                DQ_TO_DQS_mismatch = True
            if 'CMD or ADD to CLK Length Matching' in segment_list:
                CMD_mismatch = True
            if 'CTL to CLK Length Matching' in segment_list:
                CTL_mismatch = True
            if 'DLL Group Length Matching' in segment_list:
                DLL_mismatch = True

        # 获取ACT_diff或ACT_se的数据
        act_content = act_sheet.range('A1').current_region.value

        act_data_dict = dict()
        # 将表格中的start_sch_name,start_net_name,end_sch_name的数据与ATC_diff,ATC_se比较，看是否存在数据
        # 并将数据赋值给act_data_dict，包含via,total length与分段的详细数据，因为存成dict所以数据顺序不定
        stub_value = []
        for idx123 in xrange(len(start_sch_list)):
            for line in act_content:

                # # print(line[0],line[2],line[1])
                # # print(start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123])
                # 如果符合
                if start_sch_list[idx123] == line[0] and end_sch_list[idx123] == line[2] and start_net_list[idx123] == \
                        line[1]:
                    if act_data_dict.get(
                            (start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123])) == None:
                        # # print((start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123]))
                        act_data_dict[((start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123]), 'via')] = \
                            line[3]
                        act_data_dict[
                            ((start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123]), 'total_length')] = \
                            line[4].split()[1]
                        act_data_dict[(start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123])] = [x for x
                                                                                                                 in
                                                                                                                 line[
                                                                                                                 5::] if
                                                                                                                 x not in [
                                                                                                                     '',
                                                                                                                     None]]

                        # 通过过孔via数目判断是否有stub
                        # 对data数据进行处理
                        net_data = act_data_dict[(start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123])]

                    if topology_table[0][1] and topology_table[0][1].upper() in ['LOW LOSS', 'MID LOSS']:
                        if int(line[3].split()[1]) > 0:
                            # # print(topology_table[0][1])
                            stub_layer_list = []
                            for idx in range(len(net_data)):
                                item = net_data[idx]
                                if str(item).find(":") > -1:
                                    stub_layer_list += [x for x in item.split(':') if x in All_Layer_List]
                            index_list = []

                            for i in range(len(stub_layer_list)):
                                for j in range(len(layerLength_list)):
                                    if stub_layer_list[i] == layerLength_list[j].keys()[0]:
                                        # # print(j)  #具体走线层匹配的索引位置
                                        # havingLength_layer_list.append({stub_layer_list[i]:layerLength_list[j].values()[0]}) #给对应层匹配相应的长度
                                        index_list.append(j)
                            # # print(index_list) # index_list储存经过层的对应索引
                            new_index_list = []
                            new_index_list = sorted(set(index_list))
                            # # print(new_index_list) #new_index_list储存 去重并排序后的曾经过的索引
                            # # print(layerLength_list)
                            # # print(new_index_list)
                            top_stub = bottom_stub = 0

                            try:
                                for top_idx in range(1, new_index_list[-2]):
                                    # # print(layerLength_list[top_idx].values()[0])
                                    top_stub += float(layerLength_list[top_idx].values()[0])
                            except:
                                pass

                            try:
                                for bot_idx in range(new_index_list[1] + 1, len(layerLength_list) - 1):
                                    bottom_stub += float(layerLength_list[bot_idx].values()[0])
                            except:
                                pass

                            final_stub = round(max(top_stub, bottom_stub), 2)
                            stub_value.append(final_stub)
                        else:
                            pass

        if topology_table[0][1] and topology_table[0][1].upper() in ['LOW LOSS', 'MID LOSS']:
            # 获得当前的sheet名称
            active_sheet = wb.sheets.active
            PCIE_sheet = wb.sheets[active_sheet]
            for cell in PCIE_sheet.api.UsedRange.Cells:
                if cell.Value == 'Stub':
                    Stub_ind = (cell.Row, cell.Column)
                    break
            stub_table = active_sheet.range(Stub_ind).current_region.value
            # # print(stub_table)
            for idx in range(len(topology_table[2])):
                if topology_table[2][idx] == 'Total Length':
                    len_idx = idx
                    break
            if topology_table[0][1].upper() == 'LOW LOSS':

                if max(stub_value) > stub_table[1][0]:

                    active_sheet.range(start_ind[0] + 10, start_ind[1] + len_idx).value = stub_table[2][2]
                else:
                    active_sheet.range(start_ind[0] + 10, start_ind[1] + len_idx).value = stub_table[1][2]

            elif topology_table[0][1].upper() == 'MID LOSS':
                if max(stub_value) > stub_table[1][0] and max(stub_value) < stub_table[2][0]:
                    active_sheet.range(start_ind[0] + 10, start_ind[1] + len_idx).value = stub_table[2][1]
                else:
                    active_sheet.range(start_ind[0] + 10, start_ind[1] + len_idx).value = stub_table[1][1]
        else:
            pass

        # Get the list of "Skew to Target"
        # # print(segment_list)
        skew_to_target_mismatch = False
        segment_list = list(segment_list)
        target_list_all = list()
        # 对Skew to Target与Group Mismatch的情况进行讨论
        for idx in xrange(len(segment_list)):
            seg = segment_list[idx]
            if seg.find('Skew to Target') > -1:
                skew_to_target_mismatch = True
                target_list_all.append(segment_list[idx].split(':')[1::])
                target_list_all[-1] = [x[1:-1].split(',') for x in target_list_all[-1]]

                segment_list[idx] = 'Skew to Target'
        segment_list = tuple(segment_list)

        # Get the list of "Group Mismatch Target"
        group_mismatch = False
        topology_name_list_all = list()
        segment_list = list(segment_list)
        for idx in xrange(len(segment_list)):
            if segment_list[idx].find('Group Mismatch') > -1:
                group_mismatch = True
                topology_name_list_all.append(segment_list[idx].split(':')[1::])
                # # print(topology_name_list_all)
                segment_list[idx] = 'Group Mismatch'
        segment_list = tuple(segment_list)

        act_value_list = []
        all_result_list = []
        all_width_list = []
        all_layer_list = []

        for ind1 in xrange(len(start_sch_list)):
            if act_data_dict.get((start_sch_list[ind1], start_net_list[ind1], end_sch_list[ind1])):
                act_value_list.append(act_data_dict[(start_sch_list[ind1], start_net_list[ind1], end_sch_list[ind1])])

        for ind1 in xrange(len(act_value_list)):
            half_result_list = []
            half_width_list = []
            half_layer_list = []
            del_ind_list = []
            act_value = act_value_list[ind1]

            for ind2 in xrange(len(act_value)):
                if str(act_value[ind2]).find('net$') > -1:  # f7684584
                    half_width_list.append(str(act_value[ind2]).split(':')[-1])
                # if str(act_value[ind2]).find(':') > -1 and isfloat(str(act_value[ind2]).split(':')[-1]) \
                #         or str(act_value[ind2]).find('net$') > -1:
                #     half_width_list.append(str(act_value[ind2]).split(':')[-1])

                if str(act_value[ind2]).find(':') > -1:
                    for x in act_value[ind2].split(':'):
                        if x in All_Layer_List:
                            half_layer_list.append(x)

                if str(act_value[ind2]).find('net$') > -1:
                    half_layer_list.append(act_value[ind2])

                if str(act_value[ind2]).find(':') > -1:  # f7684584
                    for x in act_value[ind2].split(':'):
                        if isfloat(x):
                            half_width_list.append(x)
                # if str(act_value[ind2]).find(':') > -1 and isfloat(str(act_value[ind2]).split(':')[-2]):
                #     half_width_list.append(str(act_value[ind2]).split(':')[-2])

                if str(act_value[ind2]).find(':') == -1:
                    half_result_list.append(act_value[ind2])

            # # print(1, half_result_list)
            # # print(2, half_layer_list)
            # # print(3, half_width_list)

            all_result_list.append(half_result_list)
            all_width_list.append(half_width_list)
            all_layer_list.append(half_layer_list)

        # print(1, all_result_list)
        # print(2, all_width_list)
        # print(3, all_layer_list)

        # Calculate Mismatch for Differential Topology
        if signal_type == 'Differential':
            if total_mismatch:
                act_data_dict = TotalMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
            if layer_mismatch:
                act_data_dict = LayerMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
            # if segment_mismatch:
            #     act_data_dict = SegmentMismatch(start_sch_list, start_net_list, end_sch_list)
            if bundle_mismatch:
                act_data_dict = Skew2BundleMismatch(4, start_sch_list, start_net_list, end_sch_list, act_data_dict)
            if DQS_TO_DQS_mismatch:
                act_data_dict = DQSDLLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
            # if amd_lother_cap:
            #     act_data_dict = ForAMD_LOther(start_sch_list, start_net_list, end_sch_list, act_data_dict)
            # if amd_lother_via:
            #     act_data_dict = ForAMD_LOther(start_sch_list, start_net_list, end_sch_list, act_data_dict, CAP_type = 'PTH', Length_Type = 'Pin2Pin')
        else:
            if DQ_TO_DQ_mismatch:
                act_data_dict = DQTODQMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
            if DQ_TO_DQS_mismatch:
                act_data_dict = DQTODQSMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
            if CMD_mismatch:
                act_data_dict = CMDCTLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
            if CTL_mismatch:
                act_data_dict = CMDCTLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict, False)
            if DLL_mismatch:
                act_data_dict = DQSDLLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict, False)

        if skew_to_target_mismatch:
            act_data_dict = Skew2TargetMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict,
                                                target_list_all, act_content)
        if group_mismatch:
            # 若无特殊情况 topology_name_list_all为空[]
            act_data_dict = GroupMismatch(topology_name_list_all, start_sch_list, start_net_list, end_sch_list,
                                          act_data_dict, act_content)

        # # print(act_data_dict)
        act_data_layer_list = []
        result_dict = dict()
        check_length_seg_list = list()
        idx_all = -1

        segment_result_dict = {}

        real_segment_list = []

        # # print(trace_width_list)
        # # print(segment_list)

        # 分段索引
        segment_org_name_list = [str(x) for x in xrange(1, len(all_width_list[0]))]
        # print('segment_name_list', segment_org_name_list)

        for id1 in xrange(len(start_sch_list)):
            # print('start_sch_list', start_sch_list[id1])
            idx_all += 1
            # print(act_data_dict.get((start_sch_list[id1], start_net_list[id1], end_sch_list[id1])))
            # print(act_data_dict)
            if act_data_dict.get((start_sch_list[id1], start_net_list[id1], end_sch_list[id1])) != None:
                result_list = list()
                segment_out_name_list = []
                act_data = act_data_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]
                act_data_layer_list.append(act_data)

                # # print(11111, act_data)

                length_dict = dict()
                count1 = 0
                ignore_key1 = '$CN%d' % count1

                group_mismatch_count = 0
                skew2target_count = 0
                # print('act_data', act_data)
                for idx1 in xrange(len(segment_list)):
                    # print('segment_list', segment_list[idx1])
                    seg = segment_list[idx1]

                    layer_change_count = 0
                    cross_over_count = 0

                    # Check Order trace width, connect sch, change layer
                    check_length = False
                    check_connect_sch = False
                    check_layer_change = False
                    check_layer_change_num = 1
                    check_cross_over_num = 1
                    check_cross_over = False
                    # # print('trace_width_list', trace_width_list[idx1])

                    # 如果有trace width才去check
                    # # print('trace_width_list[idx1]', trace_width_list[idx1])
                    connect_sch_spec = ''
                    #
                    if type(trace_width_list[idx1]) == type([]):
                        check_length = True
                        # # print('trace_width_list[idx1]', trace_width_list[idx1])
                        spec_trace_width_list = trace_width_list[idx1]
                        spec_layer_type_list = layer_list[idx1]

                        try:
                            # # print('connect_sch_list[idx1]', connect_sch_list[idx1])
                            if connect_sch_list[idx1] not in ['N', 'NA', 'None', 'All', 'Connect Component']:
                                check_connect_sch = True
                                connect_sch_spec = connect_sch_list[idx1]
                        except:
                            pass

                        try:
                            if layer_change_list[idx1] == 'Y':
                                check_layer_change = True
                            elif re.findall('^Y\d$', layer_change_list[idx1]) != []:
                                check_layer_change = True
                                check_layer_change_num = int(re.findall('^Y(\d)$', layer_change_list[idx1])[0])
                        except:
                            pass

                        try:  # f7684584
                            if cross_over_list[idx1] == 'Y':
                                check_cross_over = True
                            elif re.findall('^Y\d$', cross_over_list[idx1]) != []:
                                check_cross_over = True
                                check_cross_over_num = int(re.findall('^Y(\d)$', cross_over_list[idx1])[0])
                        except:
                            pass

                    if check_length:
                        # # print('check_connect_sch', check_connect_sch)
                        # # print('ok1')
                        check_length_seg_list.append(segment_list[idx1])
                        # # print(check_length_seg_list)
                        length = 0
                        segment_part_name = ''
                        end = False
                        check_layer_change_wrong = False
                        check_cross_over_wrong = False
                        while end is False:
                            # # print(segment_list[idx1])
                            # # print(ignore_segment_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])])
                            # 没有发现！符号，都为空
                            if segment_list[idx1] in ignore_segment_dict[
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                result_list.append(0)
                                break
                            # 不知ignore_key1有何作用
                            if ignore_key1 in ignore_segment_dict[
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] or act_data == []:
                                result_list.append('NA')
                                break
                            # # print(act_data)
                            for idx2 in xrange(len(act_data)):
                                # # print('ok2')
                                # # print(1, act_data)
                                # # print('act_data', act_data)
                                act_seg = act_data[idx2]
                                # # print('act_seg', act_seg)
                                if act_data[idx2].find(':') > -1:
                                    # # print('ok3')
                                    # # print(act_data[idx2])

                                    # 实际的 trace width 和实际的 layer
                                    trace_width = ['%.3f' % float(x) for x in act_data[idx2].split(':') if isfloat(x)][
                                        0]
                                    # print('trace_width', trace_width)
                                    layer = [x for x in act_data[idx2].split(':') if x in All_Layer_List][0]
                                    layer_type = layer_type_dict[layer]
                                    # print('layer', layer, layer_type)
                                    # print(layer_type)
                                    # # print(spec_trace_width_list, spec_layer_type_list)

                                    # # print(1, act_data[idx2])
                                    # # print(11111111111111, spec_trace_width_list, trace_width)
                                    # # print(22222222222222, spec_layer_type_list, layer_type)
                                    if trace_width in spec_trace_width_list and layer_type in spec_layer_type_list:
                                        length += float('%.3f' % act_data[idx2 + 1])
                                        try:
                                            segment_part_name += segment_org_name_list[idx2]
                                            # segment_org_name_list = segment_org_name_list[idx2 + 1:]
                                        except:
                                            pass

                                        act_data = act_data[idx2 + 2::]
                                        # # print('act_dataAAAAAAAAAAAAAAAAA', act_data)
                                        # # print(check_connect_sch)
                                        # if connect_sch_spec:
                                        # # print(connect_sch_spec)
                                        # # print(1, segment_part_name)
                                        # # print(5, length)
                                        # # print(length)

                                        if check_connect_sch:
                                            if connect_sch_spec == 'Y':  # check any connect sch
                                                # # print('act_dataYYYYYYYYYYYYYY', act_data)
                                                # # print('act_data[idx2]', act_data[idx2])
                                                if act_data[idx2].split(':')[-1].find('[') == 0:
                                                    end = True
                                                    connect_sch_spec = ''
                                                    break
                                            else:  # check specific connect sch with the keyword of sch name
                                                if connect_sch_spec.find('/') > -1:
                                                    connect_sch_spec_list = connect_sch_spec.split('/')
                                                    for sch_tmp1 in connect_sch_spec_list:
                                                        if sch_tmp1.find('-') > -1:
                                                            if '[%s]' % sch_tmp1 == act_data[idx2].split(':')[-1]:
                                                                end = True
                                                                break
                                                        else:
                                                            if act_seg.split(':')[-1].find(sch_tmp1) > -1:
                                                                end = True
                                                                break
                                                else:
                                                    if connect_sch_spec.find('-') > -1:
                                                        if '[%s]' % connect_sch_spec == act_data[idx2].split(':')[-1]:
                                                            end = True
                                                            break
                                                    else:
                                                        if act_seg.split(':')[-1].find(connect_sch_spec) > -1:
                                                            end = True
                                                            break

                                        if check_layer_change:

                                            try:
                                                act_seg_next = act_data[0]

                                                if str(act_seg_next).find('net$') > -1:
                                                    end = True
                                                    break

                                                layer_next = \
                                                    [x for x in act_seg_next.split(':') if x in All_Layer_List][0]
                                                if layer != layer_next:
                                                    layer_change_count += 1
                                                    if layer_change_count == check_layer_change_num:
                                                        end = True
                                                        break
                                            except:
                                                if act_data != []:
                                                    # 表示抓到下一段net(net name改變), 需示警
                                                    result_list.append('Warning!')
                                                    check_layer_change_wrong = True
                                                    end = True
                                                    break

                                        # if check_cross_over:
                                        #     try:
                                        #         act_seg_next = act_data[0]
                                        #
                                        #         if str(act_seg_next).find('net$') > -1:
                                        #             end = True
                                        #             break
                                        #
                                        #     except:
                                        #         if act_data != []:
                                        #             result_list.append('Warning!')
                                        #             check_cross_over_wrong = True
                                        #             end = True
                                        #             break

                                        if check_cross_over:
                                            try:
                                                act_seg_next = act_data[0]
                                                if 'CROSS$' in act_seg_next.split(':'):  # f7684584
                                                    cross_over_count += 1
                                                    if cross_over_count == check_cross_over_num:
                                                        end = True
                                                        break

                                            except:
                                                if act_data != []:
                                                    result_list.append('Warning!')
                                                    check_cross_over_wrong = True
                                                    end = True
                                                    break

                                    else:  # trace width not match, jump to next spec of topology segment
                                        end = True

                                    break
                                elif act_seg.find('net$') > -1:
                                    end = True
                                    break
                            if act_data == []:
                                end = True
                            if end and not check_layer_change_wrong and not check_cross_over_wrong:
                                length_dict[seg] = float('%.3f' % length)
                                result_list.append(float('%.3f' % length))
                                # # print(2, segment_part_name)
                                if segment_part_name != '':
                                    segment_out_name_list.append(segment_part_name)

                            # # print('segment_out_name_list1', segment_out_name_list)
                            # # print('result_list', result_list)
                    else:

                        # # print('act_data', act_data)
                        if segment_list[idx1] == 'Component Name':
                            count1 += 1
                            ignore_key1 = '$CN%d' % count1

                            if len(act_data) > 0 and ignore_key1 not in ignore_segment_dict[
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] and act_data[0].find(
                                'net$') > -1:
                                # print(act_data)
                                act_data_1 = copy.deepcopy(act_data[1])
                                # 防止取到cross
                                if act_data_1.find(u'CROSS$') > -1:
                                    act_data_1 = act_data_1[7:]
                                result_list.append(act_data_1.split(':')[0][1:-1])
                                # print(1, result_list)
                                result_list.append(act_data[0][4::])
                                # print(2, result_list)
                                act_data = act_data[1::]

                            else:
                                result_list.append('NA')
                                result_list.append('NA')
                            # # print(3, result_list)

                            # segment mismatch

                        elif seg == 'Total Length':
                            result_list.append(act_data_dict[(
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]), 'total_length')])

                        elif segment_list[idx1] == 'Via Count':
                            result_list.append(act_data_dict[(
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]), 'via')].split()[1])
                        elif re.findall(r'\+|\-|\*|\/', seg) != []:

                            if seg not in ignore_segment_dict[
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                sub_seg_list = re.split(r'\+|\-|\*|\/|\(|\)', segment_list[idx1])
                                sub_seg_list = [x for x in sub_seg_list if x != '']
                                seg_tmp = str(seg)

                                sub_seg_len_list = []
                                for x in sub_seg_list:
                                    sub_seg_len_list.append(length_dict.get(x, 0))

                                for id_sub in xrange(len(sub_seg_list)):
                                    try:
                                        float(sub_seg_list[id_sub])
                                    except ValueError:
                                        seg_tmp = seg_tmp.replace(sub_seg_list[id_sub],
                                                                  str(float(sub_seg_len_list[id_sub])))

                                result_list.append(eval(seg_tmp))

                            else:

                                result_list.append('NA')

                        elif seg == 'Total Mismatch':
                            result_list.append(act_data_dict[(
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]), 'total_mismatch')])
                        elif segment_list[idx1] == 'Layer Mismatch':
                            result_list.append(act_data_dict[(
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]), 'layer_mismatch')])
                        #
                        # elif seg == 'Segment Mismatch':
                        #     act_data_dict = SegmentMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict,
                        #                                     all_width_list, all_layer_list, segment_out_name_list)
                        #     # # print(result_list)
                        #     result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                        #                                        end_sch_list[id1]), 'segment_mismatch')])
                        elif segment_list[idx1] == 'Skew to Bundle':
                            result_list.append(act_data_dict[(
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]), 'bundle_mismatch')])
                        elif seg == 'Skew to Target':

                            if '$SK2T' not in ignore_segment_dict[
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] and '$SK2T%d' % (
                                    skew2target_count + 1) not in ignore_segment_dict[
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                if target_list_all != []:
                                    result_list.append(act_data_dict[(
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                        'skew2target%d' % skew2target_count)])
                                else:
                                    result_list.append('NA')
                            else:
                                result_list.append('NA')

                            skew2target_count += 1

                        elif segment_list[idx1] == 'Group Mismatch':
                            if '$GM' not in ignore_segment_dict[
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] and '$GM%d' % (
                                    group_mismatch_count + 1) not in ignore_segment_dict[
                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                if topology_name_list_all != []:
                                    result_list.append(act_data_dict[(
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                        'group_mismatch%d' % group_mismatch_count)])
                                else:
                                    result_list.append('NA')
                            else:
                                result_list.append('NA')
                            group_mismatch_count += 1

                        elif segment_list[idx1] == 'Relative Length Spec(DQS to DQS)':
                            result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                               end_sch_list[id1]), 'DQS_TO_DQS')])
                        elif segment_list[idx1] == 'Relative Length Spec(DQ to DQ)':
                            result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                               end_sch_list[id1]), 'DQ_TO_DQ')])
                        elif segment_list[idx1] == 'Relative Length Spec(DQ to DQS)':
                            result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                               end_sch_list[id1]), 'DQ_TO_DQS')])
                        elif segment_list[idx1] == 'CMD or ADD to CLK Length Matching':
                            result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                               end_sch_list[id1]), 'CMD_TO_CLK')])
                        elif segment_list[idx1] == 'CTL to CLK Length Matching':
                            result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                               end_sch_list[id1]), 'CTL_TO_CLK')])
                        elif segment_list[idx1] == 'DLL Group Length Matching':
                            result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                               end_sch_list[id1]), 'DLL_Group')])


                        # elif segment_list[idx1] == 'AMD_LOther_CAP':
                        #     result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),'AMD_LOther_CAP')])
                        # elif seg == 'AMD_LOther_Via':
                        #     result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),'AMD_LOther_Via')])
                        elif segment_list[idx1] == 'Result':

                            cal_title_list = ['Total Mismatch', 'Layer Mismatch', 'Total Length', 'Via Count', 'Result',
                                              'Segment Mismatch', 'Skew to Bundle', 'Skew to Target', 'Group Mismatch',
                                              'Relative Length Spec(DQS to DQS)', 'Relative Length Spec(DQ to DQ)',
                                              'Relative Length Spec(DQ to DQS)', 'CMD or ADD to CLK Length Matching',
                                              'CTL to CLK Length Matching', 'DLL Group Length Matching']

                            total_length_tmp = [result_list[idx_tmp] for idx_tmp in xrange(len(segment_list)) if
                                                segment_list[idx_tmp].find('+') == -1 and segment_list[
                                                    idx_tmp] not in cal_title_list]
                            total_length_tmp = sum([float(x) for x in total_length_tmp if isfloat(x)])

                            if abs(float(act_data_dict[((start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                                        'total_length')]) - total_length_tmp) < 0.01:
                                result_list.append('Setting OK')
                            else:
                                result_list.append('Something Wrong')
                        elif seg in ['Segment Name', 'Start Segment Name']:
                            pass
                        else:
                            result_list.append('NA')

                result_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] = result_list

            else:
                result_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] = ['NA'] * len(segment_list)

        # # print(segment_list)
        # # print(trace_width_list)
        # # print(result_dict)
        color_ind_list = []
        if segment_mismatch:
            result_dict, color_ind_list = SegmentMismatch(start_sch_list, start_net_list, end_sch_list, result_dict,
                                                          all_width_list, all_layer_list, all_result_list,
                                                          segment_out_name_list)

        idx1 = -1
        for idb in xrange(len(start_sch_list)):
            idx1 += 1
            SetCellBorder(active_sheet, (result_ind2[0] + idx1, result_ind2[1] - 3))
            SetCellBorder(active_sheet, (result_ind2[0] + idx1, result_ind2[1] - 2))
            SetCellBorder(active_sheet, (result_ind2[0] + idx1, result_ind2[1] - 1))
            if result_dict.get((start_sch_list[idb], start_net_list[idb], end_sch_list[idb])) != None:
                for idx2 in xrange(len(result_dict[(start_sch_list[idb], start_net_list[idb], end_sch_list[idb])])):
                    active_sheet.range((result_ind2[0] + idx1, result_ind2[1] + idx2)).value = \
                        result_dict[(start_sch_list[idb], start_net_list[idb], end_sch_list[idb])][idx2]

        # 对segment项进行合并
        # # print(color_wid_list)
        # # # print(color_com_list)
        if segment_mismatch:
            for ind in xrange(len(start_sch_list) / 2):
                # 判断颜色
                try:
                    if color_ind_list[ind] == 0:
                        active_sheet.range(
                            (start_ind[0] + 12 + ind * 2, start_ind[1] + se1_ind)).api.Interior.ColorIndex = 4
                        active_sheet.range(
                            (start_ind[0] + 13 + ind * 2, start_ind[1] + se1_ind)).api.Interior.ColorIndex = 4
                    elif color_ind_list[ind] == 1:
                        active_sheet.range(
                            (start_ind[0] + 12 + ind * 2, start_ind[1] + se1_ind)).api.Interior.ColorIndex = 3
                        active_sheet.range(
                            (start_ind[0] + 13 + ind * 2, start_ind[1] + se1_ind)).api.Interior.ColorIndex = 3
                except:
                    pass

    SetCellFont_current_region(active_sheet, start_ind, 'Times New Roman', 12, 'l')
    SetCellBorder_current_region(active_sheet, start_ind)
    active_sheet.autofit('c')

    ShowResults(color_ind_list, specified_range=specified_range)


def BatchUpdate_Topology(specified_range=None):
    """
    批量处理Topology的更新
    :param specified_range:
    :return:
    """
    global segment_list, trace_width_list, layer_list

    # Fill the topology table by the specified spec parameter
    def LayerMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict):
        # # ##print(start_sch_list)
        # # ##print(start_net_list)
        # # ##print(end_sch_list)
        # # ##print(act_data_dict)
        if len(start_net_list) >= 2 and len(start_net_list) % 2 == 0:
            for idx in xrange(int(len(start_net_list) / 2)):
                start_sch1, start_net1, end_sch1 = start_sch_list[2 * idx], start_net_list[2 * idx], end_sch_list[
                    2 * idx]
                start_sch2, start_net2, end_sch2 = start_sch_list[2 * idx + 1], start_net_list[2 * idx + 1], \
                                                   end_sch_list[2 * idx + 1]

                act_data1 = act_data_dict[(start_sch1, start_net1, end_sch1)]
                act_data2 = act_data_dict[(start_sch2, start_net2, end_sch2)]

                layer_len_list1 = list()
                for idx1 in xrange(len(act_data1)):
                    seg_1 = act_data1[idx1]
                    if str(seg_1).find(':') > -1:
                        for x in act_data1[idx1].split(':'):
                            if x in All_Layer_List:
                                layer1 = x
                        len1 = act_data1[idx1 + 1]
                        layer_len_list1.append((layer1, len1))
                layer_len_list2 = list()
                for idx2 in xrange(len(act_data2)):
                    if str(act_data2[idx2]).find(':') > -1:
                        for x in act_data2[idx2].split(':'):
                            if x in All_Layer_List:
                                layer2 = x
                        len2 = act_data2[idx2 + 1]
                        layer_len_list2.append((layer2, len2))

                end_tmp = False
                while end_tmp is False:
                    end_tmp = True
                    for idx1 in xrange(len(layer_len_list1) - 1):
                        if layer_len_list1[idx1][0] == layer_len_list1[idx1 + 1][0]:
                            layer_len_list1[idx1] = (layer_len_list1[idx1][0], float(layer_len_list1[idx1][1]) + float(
                                layer_len_list1[idx1 + 1][1]))
                            layer_len_list1.pop(idx1 + 1)
                            end_tmp = False
                            break
                end_tmp = False
                while end_tmp is False:
                    end_tmp = True
                    for idx2 in xrange(len(layer_len_list2) - 1):
                        if layer_len_list2[idx2][0] == layer_len_list2[idx2 + 1][0]:
                            layer_len_list2[idx2] = (layer_len_list2[idx2][0], float(layer_len_list2[idx2][1]) + float(
                                layer_len_list2[idx2 + 1][1]))
                            layer_len_list2.pop(idx2 + 1)
                            end_tmp = False
                            break

                max_mismatch = 0
                if len(layer_len_list1) == len(layer_len_list2):
                    for seg_idx in xrange(len(layer_len_list1)):
                        if layer_len_list1[seg_idx][0] == layer_len_list2[seg_idx][0]:
                            if abs(layer_len_list1[seg_idx][1] - layer_len_list2[seg_idx][1]) > max_mismatch:
                                max_mismatch = abs(layer_len_list1[seg_idx][1] - layer_len_list2[seg_idx][1])
                else:
                    max_mismatch = -1

                act_data_dict[((start_sch1, start_net1, end_sch1), 'layer_mismatch')] = max_mismatch
                act_data_dict[((start_sch2, start_net2, end_sch2), 'layer_mismatch')] = max_mismatch
        else:
            for idx in xrange(len(start_net_list)):
                act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'layer_mismatch')] = 'NA'

        return act_data_dict

    def SegmentMismatch(start_sch_list, start_net_list, end_sch_list, result_dict, all_width_list, all_layer_list,
                        all_result_list, segment_out_name_list):
        # # ##print('all_width_list', all_width_list)
        # # ##print('all_layer_list', all_layer_list)
        # # ##print('all_result_list', all_result_list)

        # # ##print('segment_out_name_list', segment_out_name_list)
        if result_dict:
            if len(start_net_list) >= 2 and len(start_net_list) % 2 == 0:

                for idx in xrange(int(len(start_net_list) / 2)):
                    start_sch1, start_net1, end_sch1 = start_sch_list[2 * idx], start_net_list[2 * idx], end_sch_list[
                        2 * idx]
                    start_sch2, start_net2, end_sch2 = start_sch_list[2 * idx + 1], start_net_list[2 * idx + 1], \
                                                       end_sch_list[2 * idx + 1]

                    result_list1 = result_dict[(start_sch1, start_net1, end_sch1)]
                    result_list2 = result_dict[(start_sch2, start_net2, end_sch2)]

                    # # ##print(1, result_list1)
                    # # ##print(2, result_list2)
                    # # ##print(3, segment_list)
                    # 找到segment位置
                    com_del_ind_list = []
                    for se_ind in xrange(len(segment_list)):
                        if segment_list[se_ind] == 'Segment Mismatch':
                            seg_table_ind = se_ind
                        if segment_list[se_ind] == 'Total Length':
                            seg_name_ind = se_ind

                    # 找出segment_name
                    segment_name_list = list(segment_list[:seg_name_ind])
                    # # ##print(segment_name_list)
                    # # ##print('segment_name_list', segment_name_list)
                    all_width_list_1 = all_width_list[2 * idx]
                    all_width_list_2 = all_width_list[2 * idx + 1]
                    all_layer_list_1 = all_layer_list[2 * idx]
                    all_layer_list_2 = all_layer_list[2 * idx + 1]
                    all_result_list_1 = all_result_list[2 * idx]
                    all_result_list_2 = all_result_list[2 * idx + 1]

                    def addComponent(param):
                        for x in xrange(len(param)):
                            if str(param[x]).find('net$') > -1:
                                param.insert(x + 1, 'Component Name')

                    addComponent(all_width_list_1)
                    addComponent(all_width_list_2)
                    addComponent(all_layer_list_1)
                    addComponent(all_layer_list_2)
                    addComponent(all_result_list_1)
                    addComponent(all_result_list_2)

                    # 找出segement的值
                    for x in xrange(len(all_width_list_1)):
                        if all_width_list_1[x].find('net$') > -1:
                            segment_out_name_list.insert(x, '0')
                            segment_out_name_list.insert(x + 1, '0')
                    # # ##print(segment_out_name_list)

                    count = 0
                    seg_name_dict = {}
                    segment_out_name_number_list = []
                    for x in xrange(len(segment_out_name_list)):
                        segment_value_list = []
                        for y in xrange(len(segment_out_name_list[x])):
                            segment_out_name_number_list.append(count)
                            segment_value_list.append(count)
                            count += 1
                        if len(segment_value_list) > 1:
                            for z in xrange(1, len(segment_value_list) + 1):
                                seg_name_dict[segment_value_list[z - 1]] = segment_name_list[x] + '_' + str(z)
                        else:
                            seg_name_dict[segment_value_list[0]] = segment_name_list[x]
                    # # ##print(segment_out_name_number_list)
                    # # ##print(seg_name_dict)
                    # 判断是否是segment
                    seg_com_ind_list = []
                    seg_wid_ind_list = []
                    seg_lay_ind_list = []

                    for wid_ind in xrange(len(all_width_list_1) - 1):
                        if all_width_list_1[wid_ind] == 'Component Name':
                            seg_com_ind_list.append(wid_ind)

                        if all_width_list_1[wid_ind] != 'Component Name' and all_width_list_1[wid_ind + 1] \
                                != 'Component Name' and all_width_list_1[wid_ind].find('net$') == -1 and \
                                all_width_list_1[wid_ind + 1].find('net$') == -1:
                            if all_width_list_1[wid_ind] != all_width_list_1[wid_ind + 1]:
                                seg_wid_ind_list.append(wid_ind)

                    for layer_ind in xrange(len(all_layer_list_1) - 1):
                        if all_layer_list_1[layer_ind] != 'Component Name' and all_layer_list_1[layer_ind + 1] \
                                != 'Component Name' and all_layer_list_1[layer_ind].find('net$') == -1 and \
                                all_layer_list_1[layer_ind + 1].find('net$') == -1:
                            if all_layer_list_1[layer_ind] != all_layer_list_1[layer_ind + 1]:
                                seg_lay_ind_list.append(layer_ind)

                    # # ##print(6, seg_com_ind_list)
                    # # ##print(7, seg_wid_ind_list)
                    # # ##print(8, seg_lay_ind_list)

                    seg_name_list = []
                    seg_value_list = []
                    # 计算segment mismatch
                    for ind1 in seg_wid_ind_list:
                        seg_name_list.append(seg_name_dict.get(ind1))
                        seg_name_list.append(seg_name_dict.get(ind1 + 1))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind1]) - float(all_result_list_2[ind1])), 2))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind1 + 1]) - float(all_result_list_2[ind1 + 1])), 2))
                    for ind2 in seg_com_ind_list:
                        seg_name_list.append(seg_name_dict.get(ind2 - 2))
                        seg_name_list.append(seg_name_dict.get(ind2 + 1))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind2 - 2]) - float(all_result_list_2[ind2 - 2])), 2))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind2 + 1]) - float(all_result_list_2[ind2 + 1])), 2))
                    for ind3 in seg_lay_ind_list:
                        seg_name_list.append(seg_name_dict.get(ind3))
                        seg_name_list.append(seg_name_dict.get(ind3 + 1))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind3]) - float(all_result_list_2[ind3])), 2))
                        seg_value_list.append(
                            round(abs(float(all_result_list_1[ind3 + 1]) - float(all_result_list_2[ind3 + 1])), 2))

                    seg_name_out_list = []
                    seg_del_ind_list = []
                    seg_value_out_list = []

                    for x in xrange(len(seg_name_list)):
                        if seg_name_list[x] not in seg_name_out_list:
                            seg_name_out_list.append(seg_name_list[x])
                        else:
                            seg_del_ind_list.append(x)

                    seg_value_out_list = [seg_value_list[x] for x in xrange(len(seg_name_list)) if
                                          x not in seg_del_ind_list]

                    seg_out_ind = []
                    seg_max_ind = None

                    # # ##print('seg_name_list', seg_name_list)
                    # # ##print('seg_value_list', seg_value_list)
                    ######################################################
                    # 获得管控值
                    max_seg = spec_max_list[seg_table_ind]
                    min_seg = spec_min_list[seg_table_ind]

                    # 检测有没有超过管控
                    for seg_ind in xrange(len(seg_value_out_list)):
                        # # ##print(seg_ind, seg_value_list[seg_ind])
                        if seg_value_out_list[seg_ind] >= float(max_seg) or seg_value_out_list[seg_ind] < float(
                                min_seg):
                            seg_out_ind.append(seg_ind)
                    if seg_out_ind == []:
                        seg_max_ind = seg_value_out_list.index(max(seg_value_out_list))

                    # 输出数值并改变颜色
                    seg_out_value = ''

                    # # ##print(seg_out_ind, seg_max_ind)

                    if seg_out_ind:
                        color_ind_list.append(1)
                        for x in seg_out_ind:
                            seg_out_value += seg_name_out_list[x] + ':' + str(seg_value_out_list[x]) + '; '

                    # # ##print(seg_value_list)
                    # # ##print(seg_max_ind)
                    if seg_max_ind != None:
                        # # ##print('ok')
                        color_ind_list.append(0)
                        seg_out_value = seg_value_list[seg_max_ind]
                    # # ##print(seg_out_value)
                    result_list1[seg_table_ind] = seg_out_value
                    result_list2[seg_table_ind] = seg_out_value

                    result_dict[(start_sch1, start_net1, end_sch1)] = result_list1
                    result_dict[(start_sch2, start_net2, end_sch2)] = result_list2
            # # ##print(color_ind_list)
            return result_dict, color_ind_list
        else:
            for idx in xrange(len(start_net_list)):
                act_data_dict[((start_sch_list[idx], start_net_list[idx],
                                end_sch_list[idx]), 'segment_mismatch')] = 'NA'
            # # ##print(act_data_dict)
            # # ##print('color_ind_list', color_ind_list)
            # # ##print(act_data_dict)
            return act_data_dict

    def TotalMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict):
        if len(start_net_list) >= 2 and len(start_net_list) % 2 == 0:
            for idx in xrange(int(len(start_net_list) / 2)):
                start_sch1, start_net1, end_sch1 = start_sch_list[2 * idx], start_net_list[2 * idx], end_sch_list[
                    2 * idx]
                start_sch2, start_net2, end_sch2 = start_sch_list[2 * idx + 1], start_net_list[2 * idx + 1], \
                                                   end_sch_list[2 * idx + 1]

                total_length1 = act_data_dict[((start_sch1, start_net1, end_sch1), 'total_length')]
                total_length2 = act_data_dict[((start_sch2, start_net2, end_sch2), 'total_length')]

                len1 = float(total_length1)
                len2 = float(total_length2)

                act_data_dict[((start_sch1, start_net1, end_sch1), 'total_mismatch')] = abs(len1 - len2)
                act_data_dict[((start_sch2, start_net2, end_sch2), 'total_mismatch')] = abs(len1 - len2)
        else:
            for idx in xrange(len(start_net_list)):
                act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'total_mismatch')] = 'NA'

        return act_data_dict

    def Skew2BundleMismatch(unit_bundle_number, start_sch_list, start_net_list, end_sch_list, act_data_dict):

        if len(start_net_list) >= unit_bundle_number and len(start_net_list) % unit_bundle_number == 0:
            for idx1 in xrange(len(start_net_list)):
                if idx1 % unit_bundle_number == 0:
                    len_list_tmp = list()
                    bundle_max_mismatch = 0
                    for idx2 in xrange(unit_bundle_number):
                        len_tmp_ = float(act_data_dict[(
                            (start_sch_list[idx1 + idx2], start_net_list[idx1 + idx2], end_sch_list[idx1 + idx2]),
                            'total_length')])
                        len_list_tmp.append(len_tmp_)
                    bundle_max_mismatch = max(len_list_tmp) - min(len_list_tmp)

                    for idx2 in xrange(unit_bundle_number):
                        act_data_dict[(
                            (start_sch_list[idx1 + idx2], start_net_list[idx1 + idx2], end_sch_list[idx1 + idx2]),
                            'bundle_mismatch')] = bundle_max_mismatch
        else:
            for idx1 in xrange(len(start_net_list)):
                act_data_dict[
                    ((start_sch_list[idx1], start_net_list[idx1], end_sch_list[idx1]), 'bundle_mismatch')] = 'NA'

        return act_data_dict

    def Skew2TargetMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict, target_list_all, act_content,
                            key_title='skew2target'):

        if target_list_all != []:
            for idx_t in xrange(len(target_list_all)):
                target_list = target_list_all[idx_t]
                if key_title.find('skew2target') > -1:
                    key_title = 'skew2target%d' % idx_t

                act_content_se = wb.sheets['ACT_se'].range('A1').current_region.value
                act_content_diff = wb.sheets['ACT_diff'].range('A1').current_region.value

                if act_content_diff != None and act_content_se != None:
                    act_content_all = act_content_se + act_content_diff
                elif act_content_diff != None and act_content_se == None:
                    act_content_all = act_content_diff
                elif act_content_diff == None and act_content_se != None:
                    act_content_all = act_content_se
                else:
                    act_content_all = []

                for start_sch, end_sch, start_net in target_list_all[idx_t]:
                    for line in act_content_all:
                        if start_sch == line[0] and end_sch == line[2] and start_net == line[1]:
                            if act_data_dict.get((start_sch, start_net, end_sch)) == None:
                                act_data_dict[((start_sch, start_net, end_sch), 'via')] = line[3]
                                act_data_dict[((start_sch, start_net, end_sch), 'total_length')] = line[4].split()[1]

                                act_data_dict[(start_sch, start_net, end_sch)] = [x for x in line[5::] if
                                                                                  x not in ['', None]]
                                break

                target_len_list = list()

                for target in target_list:
                    start_sch, end_sch, start_net = target[0], target[1], target[2]

                    target_len = float(act_data_dict[((start_sch, start_net, end_sch), 'total_length')])

                    target_len_list.append(target_len)

                for i22 in xrange(len(start_sch_list)):
                    item_len = float(
                        act_data_dict[((start_sch_list[i22], start_net_list[i22], end_sch_list[i22]), 'total_length')])
                    act_data_dict[((start_sch_list[i22], start_net_list[i22], end_sch_list[i22]), key_title)] = max(
                        [abs(item_len - x) for x in target_len_list])
        else:
            for i33 in xrange(len(start_sch_list)):
                act_data_dict[((start_sch_list[i33], start_net_list[i33], end_sch_list[i33]), key_title)] = 'NA'

        return act_data_dict

    def GroupMismatch(topology_name_list_all, start_sch_list, start_net_list, end_sch_list, act_data_dict, act_content):

        wb = Book(xlsm_path).caller()
        active_sheet = wb.sheets.active  # Get the active sheet object

        for idx_ in xrange(len(topology_name_list_all)):
            data_list = list()
            for topology_name in topology_name_list_all[idx_]:
                for cell in active_sheet.api.UsedRange.Cells:
                    if cell.Value == 'Topology' and active_sheet.range(
                            (cell.Row, cell.Column + 1)).value == topology_name:
                        data_list.append(active_sheet.range((cell.Row, cell.Column)).current_region.value)

            target_list = list()
            for data in data_list:
                for x in data:
                    if x[2] == 'Net Name':
                        data_idx = data.index(x) + 1
                target_list += [(x[0], x[1], x[2].split('!')[0]) for x in data[data_idx::]]

            target_list_all = [target_list]
            act_data_dict = Skew2TargetMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict,
                                                target_list_all, act_content, key_title='group_mismatch%d' % idx_)

        return act_data_dict

    def DQSDLLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict, DQS_flag=True):

        for idx in xrange(int(len(start_net_list) / 2)):
            start_sch1, start_net1, end_sch1 = start_sch_list[2 * idx], start_net_list[2 * idx], end_sch_list[2 * idx]
            start_sch2, start_net2, end_sch2 = start_sch_list[2 * idx + 1], start_net_list[2 * idx + 1], end_sch_list[
                2 * idx + 1]

            act_data1 = act_data_dict[(start_sch1, start_net1, end_sch1), 'total_length']
            act_data2 = act_data_dict[(start_sch2, start_net2, end_sch2), 'total_length']
            if DQS_flag:
                act_data_dict[(start_sch1, start_net1, end_sch1), 'DQS_TO_DQS'] = \
                    round(abs(float(act_data1) - float(act_data2)), 2)
                act_data_dict[(start_sch2, start_net2, end_sch2), 'DQS_TO_DQS'] = \
                    round(abs(float(act_data1) - float(act_data2)), 2)
            else:
                act_data_dict[(start_sch1, start_net1, end_sch1), 'DLL_Group'] = \
                    round(abs(float(act_data1) - float(act_data2)), 2)
                act_data_dict[(start_sch2, start_net2, end_sch2), 'DLL_Group'] = \
                    round(abs(float(act_data1) - float(act_data2)), 2)

        return act_data_dict

    def DQTODQMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict):

        DQ_value_list = [float(act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]),
                                              'total_length')]) for idx in xrange(len(start_net_list))]
        for idx in xrange(len(start_net_list)):
            act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'DQ_TO_DQ')] = \
                round(max(abs(max(DQ_value_list) - DQ_value_list[idx]), abs(DQ_value_list[idx] -
                                                                            min(DQ_value_list))), 2)

        return act_data_dict

    def DQTODQSMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict):

        DQ_value_list = [float(act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]),
                                              'total_length')]) for idx in xrange(len(start_net_list))]

        DQS_table_ind = getPreTable()
        DQS_value_list = []

        if DQS_table_ind:
            topology_table = active_sheet.range(DQS_table_ind).current_region.value
            for ind in xrange(len(topology_table[2])):
                if topology_table[2][ind] == 'Total Length':
                    len_ind = ind
            # 获取DQS长度数据
            for value in topology_table[12:]:
                DQS_value_list.append(value[len_ind])

        # dqs_max = max(DQS_value_list)
        # dqs_min = min(DQS_value_list)
        # DQS_real_list = []
        # for i in xrange(len(DQ_value_list)):
        #     lst = round(dqs_max-DQ_value_list[i],2)
        #     lst1 = round(DQ_value_list[i]-dqs_min,2)
        #     DQS_real_list.append(max(lst,lst1))
        # for idx in xrange(len(start_net_list)):
        #     act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'DQ_TO_DQS')] = DQS_real_list[idx]
        for idx in xrange(len(start_net_list)):
            act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'DQ_TO_DQS')] = \
                round(max(abs(max(DQS_value_list) - DQ_value_list[idx]), abs(DQ_value_list[idx] -
                                                                             min(DQS_value_list))), 2)
        return act_data_dict

    def CMDCTLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict, CMD_flag=True):

        # # ##print(start_sch_list, start_net_list, end_sch_list, act_data_dict, CMD_flag)
        CMD_value_list = [float(act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]),
                                               'total_length')]) for idx in xrange(len(start_net_list))]
        # # ##print(1, CMD_value_list)
        CLK_table_ind = getFirstTable()
        CLK_value_list = []
        # # ##print(2, CLK_table_ind)
        if CLK_table_ind:
            topology_table = active_sheet.range(CLK_table_ind).current_region.value
            for ind in xrange(len(topology_table[2])):
                if topology_table[2][ind] == 'Total Length':
                    len_ind = ind
            # 获取DQS长度数据
            for value in topology_table[12:]:
                CLK_value_list.append(value[len_ind])

        if CMD_flag:
            for idx in xrange(len(start_net_list)):
                act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'CMD_TO_CLK')] = \
                    round(max(abs(max(CLK_value_list) - CMD_value_list[idx]),
                              abs(CMD_value_list[idx] - min(CLK_value_list))), 2)
        else:
            for idx in xrange(len(start_net_list)):
                act_data_dict[((start_sch_list[idx], start_net_list[idx], end_sch_list[idx]), 'CTL_TO_CLK')] = \
                    round(max(abs(max(CLK_value_list) - CMD_value_list[idx]), abs(CMD_value_list[idx] -
                                                                                  min(CLK_value_list))), 2)
        return act_data_dict

    ClearCheckResults(specified_range=specified_range)
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active
    color_ind_list = []
    allegro_report_path, layer_type_dict, start_sch_name_list, progress_ind, All_Layer_List = GetSetting()
    All_Layer_List = list(set(layer_type_dict.keys()))

    # 获取每层的厚度以及总的厚度
    # 计算stub的时候需要每层厚度
    SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data, npr_brd_data = read_allegro_data(
        allegro_report_path)
    layerLength_list = []
    data = stackup_brd_data.GetData()
    data = data[1:-3]

    selection_range = wb.app.selection
    selected_flag = False
    if selection_range.value == 'Topology':
        selected_flag = True
        selected_idx = (selection_range.row, selection_range.column)

    # ##print(data)
    for idx in range(len(data)):
        data[idx] = data[idx].split(',')
        layerLength_list.append({data[idx][0]: data[idx][3]})  # 存储所有层的厚度

    for cell in active_sheet.api.UsedRange.Cells:
        if cell.Value == 'Topology':
            cell_idx = (cell.Row, cell.Column)
            # 从选中处开始
            if selected_flag:
                if cell_idx[0] >= selected_idx[0]:
                    topology_table = active_sheet.range(cell_idx).current_region.value
                    for idx1 in xrange(len(topology_table)):
                        line = topology_table[idx1]
                        # print(line)
                        for idx2 in xrange(len(line)):
                            if line[idx2] == 'Net Name':
                                spec_table_length = idx1 + 1
                                break
                    start_sch_ind = (cell_idx[0] + spec_table_length, cell_idx[1])
                    end_sch_ind = (start_sch_ind[0], start_sch_ind[1] + 1)
                    start_net_ind = (end_sch_ind[0], end_sch_ind[1] + 1)
                    result_ind2 = (start_net_ind[0], start_net_ind[1] + 1)
                    active_sheet.range(result_ind2).expand('table').clear()

                    for se_ind in xrange(len(topology_table[2])):
                        # 获取Segment Mismatch的管控条件
                        if topology_table[2][se_ind] and topology_table[2][se_ind].find('Segment Mismatch') > -1:
                            se1_ind = se_ind
                    for idx1 in xrange(len(topology_table)):
                        line = topology_table[idx1]
                        for idx2 in xrange(len(line)):
                            if line[idx2] == 'Net Name':
                                spec_table_length = idx1 + 1
                                break
                    # ##print('spec_table_length',spec_table_length)

                    # 获取坐标
                    start_sch_ind = (cell_idx[0] + spec_table_length, cell_idx[1])
                    end_sch_ind = (start_sch_ind[0], start_sch_ind[1] + 1)
                    start_net_ind = (end_sch_ind[0], end_sch_ind[1] + 1)
                    result_ind1 = (start_net_ind[0], start_net_ind[1] - 2)
                    result_ind2 = (start_net_ind[0], start_net_ind[1] + 1)

                    for idxx in xrange(len(topology_table)):
                        # print(topology_table)
                        row_tmp = topology_table[idxx]
                        if 'Start Segment Name' in row_tmp:
                            segment_list = topology_table[idxx][3::]
                            segment_list = [x for x in segment_list if x not in ['', None]]
                            result_idx = segment_list.index('Result')
                            # ##print(result_idx)
                        elif 'Layer' in topology_table[idxx]:
                            layer_list = []
                            for x in topology_table[idxx][3::]:
                                layer_list.append(str(x))
                            layer_list = layer_list[0:result_idx + 1]
                            layer_list = [x.split('/') for x in layer_list]
                            # ##print(layer_list)
                        elif 'Layer Change' in topology_table[idxx]:
                            layer_change_list = []
                            for x in topology_table[idxx][3::]:
                                layer_change_list.append(str(x))
                            layer_change_list = layer_change_list[0:result_idx + 1]
                        elif 'Connect Component' in row_tmp:
                            connect_sch_list = []
                            for x in topology_table[idxx][3::]:
                                connect_sch_list.append(str(x))
                            connect_sch_list = connect_sch_list[0:result_idx + 1]
                        elif 'Cross Over' in row_tmp:
                            cross_over_list = []
                            for x in topology_table[idxx][3::]:
                                cross_over_list.append(str(x))
                            cross_over_list = cross_over_list[0:result_idx + 1]
                        elif 'Trace Width' in topology_table[idxx]:
                            trace_width_list = []
                            for x in topology_table[idxx][3::]:
                                trace_width_list.append(str(x))
                            trace_width_list = trace_width_list[0:result_idx + 1]
                            for idx in xrange(len(trace_width_list)):
                                # 得到的数都为浮点数
                                tw = trace_width_list[idx]
                                # 数据为浮点数时，如3.5，或者有 / 时（差分信号）
                                if isfloat(trace_width_list[idx]) or tw.find('/') > -1:
                                    trace_width_list[idx] = ['%.3f' % float(x) for x in
                                                             trace_width_list[idx].split('/')]
                        elif 'Space' in row_tmp:
                            space_list = []
                            for x in topology_table[idxx][3::]:
                                space_list.append(str(x))
                            space_list = space_list[0:result_idx + 1]
                        elif 'Min' in row_tmp:
                            spec_min_list = []
                            for x in topology_table[idxx][3::]:
                                spec_min_list.append(str(x))
                            spec_min_list = spec_min_list[0:result_idx + 1]
                        elif 'Max' in row_tmp:
                            spec_max_list = []
                            for x in topology_table[idxx][3::]:
                                spec_max_list.append(str(x))
                            spec_max_list = spec_max_list[0:result_idx + 1]

                    signal_type = topology_table[1][1]
                    #####################################################
                    table1 = active_sheet.range(result_ind1).expand('table').value
                    if type(table1[0]) == type(u''):
                        table1 = [table1]
                    ########################################################
                    start_sch_list = [tt[0] for tt in table1]
                    end_sch_list = [tt[1] for tt in table1]
                    start_net_list = [tt[2] for tt in table1]

                    # Check ignore syntax to skip specified segment check
                    ignore_segment_dict = dict()
                    start_net_list = list(start_net_list)
                    for idx in xrange(len(start_net_list)):
                        net_name = start_net_list[idx]
                        # !代表什么
                        if net_name.find('!') > -1:
                            ignore_segment_list = start_net_list[idx].split('!')[1::]
                            start_net_list[idx] = start_net_list[idx].split('!')[0]
                            ignore_segment_dict[
                                (start_sch_list[idx], start_net_list[idx], end_sch_list[idx])] = ignore_segment_list
                        else:
                            ignore_segment_dict[(start_sch_list[idx], start_net_list[idx], end_sch_list[idx])] = []

                    start_net_list = tuple(start_net_list)

                    layer_mismatch, segment_mismatch, bundle_mismatch, total_mismatch, DQS_TO_DQS_mismatch \
                        = False, False, False, False, False
                    DQ_TO_DQ_mismatch, DQ_TO_DQS_mismatch, CMD_mismatch, CTL_mismatch, DLL_mismatch \
                        = False, False, False, False, False
                    if signal_type == 'Differential':
                        act_sheet = wb.sheets['ACT_diff']
                        if 'Layer Mismatch' in segment_list:
                            layer_mismatch = True
                        if 'Segment Mismatch' in segment_list:
                            segment_mismatch = True
                        if 'Skew to Bundle' in segment_list:
                            bundle_mismatch = True
                        if 'Total Mismatch' in segment_list:
                            total_mismatch = True
                        if 'Relative Length Spec(DQS to DQS)' in segment_list:
                            DQS_TO_DQS_mismatch = True
                    elif signal_type == 'Single-ended':
                        act_sheet = wb.sheets['ACT_se']
                        if 'Relative Length Spec(DQ to DQ)' in segment_list:
                            DQ_TO_DQ_mismatch = True
                        if 'Relative Length Spec(DQ to DQS)' in segment_list:
                            DQ_TO_DQS_mismatch = True
                        if 'CMD or ADD to CLK Length Matching' in segment_list:
                            CMD_mismatch = True
                        if 'CTL to CLK Length Matching' in segment_list:
                            CTL_mismatch = True
                        if 'DLL Group Length Matching' in segment_list:
                            DLL_mismatch = True

                    act_content = act_sheet.range('A1').current_region.value  # net name 所有信息
                    # print(act_content)
                    act_data_dict = dict()
                    stub_value = []
                    for idx123 in xrange(len(start_sch_list)):
                        for line in act_content:
                            if start_sch_list[idx123] == line[0] and end_sch_list[idx123] == line[2] and start_net_list[
                                idx123] == line[1]:
                                if act_data_dict.get(
                                        (start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123])) == None:
                                    # ##print((start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123]))
                                    act_data_dict[
                                        ((start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123]),
                                         'via')] = \
                                        line[3]
                                    act_data_dict[
                                        ((start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123]),
                                         'total_length')] = line[4].split()[1]
                                    act_data_dict[
                                        (start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123])] = [x for
                                                                                                                   x in
                                                                                                                   line[
                                                                                                                   5::]
                                                                                                                   if
                                                                                                                   x not in [
                                                                                                                       '',
                                                                                                                       None]]

                                    # 通过过孔via数目判断是否有stub
                                    # 对data数据进行处理
                                    net_data = act_data_dict[
                                        (start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123])]

                                if topology_table[0][1] and topology_table[0][1].upper() in ['LOW LOSS', 'MID LOSS']:
                                    if int(line[3].split()[1]) > 0:
                                        # # ##print(topology_table[0][1])
                                        stub_layer_list = []
                                        for idx in range(len(net_data)):
                                            item = net_data[idx]
                                            if str(item).find(":") > -1:
                                                stub_layer_list += [x for x in item.split(':') if x in All_Layer_List]
                                        index_list = []
                                        for i in range(len(stub_layer_list)):
                                            for j in range(len(layerLength_list)):
                                                if stub_layer_list[i] == layerLength_list[j].keys()[0]:
                                                    # # ##print(j)  #具体走线层匹配的索引位置
                                                    # havingLength_layer_list.append({stub_layer_list[i]:layerLength_list[j].values()[0]}) #给对应层匹配相应的长度
                                                    index_list.append(j)
                                        new_index_list = []
                                        new_index_list = sorted(set(index_list))
                                        top_stub = bottom_stub = 0

                                        try:
                                            for top_idx in range(1, new_index_list[-2]):
                                                # print(layerLength_list[top_idx].values()[0])
                                                top_stub += float(layerLength_list[top_idx].values()[0])
                                        except:
                                            pass

                                        try:
                                            for bot_idx in range(new_index_list[1] + 1, len(layerLength_list) - 1):
                                                bottom_stub += float(layerLength_list[bot_idx].values()[0])
                                        except:
                                            pass

                                        final_stub = round(max(top_stub, bottom_stub), 2)
                                        stub_value.append(final_stub)
                                    else:
                                        pass
                    # print('act_data_dict',act_data_dict)
                    if topology_table[0][1] and topology_table[0][1].upper() in ['LOW LOSS', 'MID LOSS']:
                        # 获得当前的sheet名称
                        active_sheet = wb.sheets.active
                        PCIE_sheet = wb.sheets[active_sheet]
                        for cell in PCIE_sheet.api.UsedRange.Cells:
                            if cell.Value == 'Stub':
                                Stub_ind = (cell.Row, cell.Column)
                                break
                        stub_table = active_sheet.range(Stub_ind).current_region.value
                        # # ##print(stub_table)
                        for idx in range(len(topology_table[2])):
                            if topology_table[2][idx] == 'Total Length':
                                len_idx = idx
                                break
                        if topology_table[0][1].upper() == 'LOW LOSS':
                            if max(stub_value) > stub_table[1][0]:
                                active_sheet.range(cell.Row[0] + 10, cell.Column[1] + len_idx).value = stub_table[2][2]
                            else:
                                active_sheet.range(cell.Row[0] + 10, cell.Column[1] + len_idx).value = stub_table[1][2]

                        elif topology_table[0][1].upper() == 'MID LOSS':
                            if max(stub_value) > stub_table[1][0] and max(stub_value) < stub_table[2][0]:
                                active_sheet.range(cell.Row[0] + 10, cell.Column[1] + len_idx).value = stub_table[2][1]
                            else:
                                active_sheet.range(cell.Row[0] + 10, cell.Column[1] + len_idx).value = stub_table[1][1]
                    else:
                        pass

                    skew_to_target_mismatch = False
                    segment_list = list(segment_list)
                    target_list_all = list()
                    # 对Skew to Target与Group Mismatch的情况进行讨论
                    for idx in xrange(len(segment_list)):
                        seg = segment_list[idx]
                        if seg.find('Skew to Target') > -1:
                            skew_to_target_mismatch = True
                            target_list_all.append(segment_list[idx].split(':')[1::])
                            target_list_all[-1] = [x[1:-1].split(',') for x in target_list_all[-1]]
                            segment_list[idx] = 'Skew to Target'
                    segment_list = tuple(segment_list)
                    # Get the list of "Group Mismatch Target"
                    group_mismatch = False
                    topology_name_list_all = list()
                    segment_list = list(segment_list)
                    for idx in xrange(len(segment_list)):
                        if segment_list[idx].find('Group Mismatch') > -1:
                            group_mismatch = True
                            topology_name_list_all.append(segment_list[idx].split(':')[1::])
                            # # ##print(topology_name_list_all)
                            segment_list[idx] = 'Group Mismatch'
                    segment_list = tuple(segment_list)
                    # ##print(segment_list)

                    act_value_list = []
                    all_result_list = []
                    all_width_list = []
                    all_layer_list = []
                    for ind1 in xrange(len(start_sch_list)):
                        if act_data_dict.get((start_sch_list[ind1], start_net_list[ind1], end_sch_list[ind1])):
                            act_value_list.append(
                                act_data_dict[(start_sch_list[ind1], start_net_list[ind1], end_sch_list[ind1])])
                    for ind1 in xrange(len(act_value_list)):
                        half_result_list = []
                        half_width_list = []
                        half_layer_list = []
                        del_ind_list = []
                        act_value = act_value_list[ind1]

                        for ind2 in xrange(len(act_value)):
                            if str(act_value[ind2]).find('net$') > -1:  # f7684584
                                half_width_list.append(str(act_value[ind2]).split(':')[-1])
                            # if str(act_value[ind2]).find(':') > -1 and isfloat(str(act_value[ind2]).split(':')[-1]) \
                            #         or str(act_value[ind2]).find('net$') > -1:
                            #     half_width_list.append(str(act_value[ind2]).split(':')[-1])

                            if str(act_value[ind2]).find(':') > -1:
                                for x in act_value[ind2].split(':'):
                                    if x in All_Layer_List:
                                        half_layer_list.append(x)

                            if str(act_value[ind2]).find('net$') > -1:
                                half_layer_list.append(act_value[ind2])

                            if str(act_value[ind2]).find(':') > -1:  # f7684584
                                for x in act_value[ind2].split(':'):
                                    if isfloat(x):
                                        half_width_list.append(x)
                            # if str(act_value[ind2]).find(':') > -1 and isfloat(str(act_value[ind2]).split(':')[-2]):
                            #     half_width_list.append(str(act_value[ind2]).split(':')[-2])

                            if str(act_value[ind2]).find(':') == -1:
                                half_result_list.append(act_value[ind2])

                        all_result_list.append(half_result_list)
                        all_width_list.append(half_width_list)
                        all_layer_list.append(half_layer_list)

                    if signal_type == 'Differential':
                        if total_mismatch:
                            act_data_dict = TotalMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                        if layer_mismatch:
                            act_data_dict = LayerMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                        if bundle_mismatch:
                            act_data_dict = Skew2BundleMismatch(4, start_sch_list, start_net_list, end_sch_list,
                                                                act_data_dict)
                        if DQS_TO_DQS_mismatch:
                            act_data_dict = DQSDLLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                        # if amd_lother_cap:
                        #     act_data_dict = ForAMD_LOther(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                        # if amd_lother_via:
                        #     act_data_dict = ForAMD_LOther(start_sch_list, start_net_list, end_sch_list, act_data_dict, CAP_type = 'PTH', Length_Type = 'Pin2Pin')
                    else:
                        if DQ_TO_DQ_mismatch:
                            act_data_dict = DQTODQMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                        if DQ_TO_DQS_mismatch:
                            act_data_dict = DQTODQSMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                        if CMD_mismatch:
                            act_data_dict = CMDCTLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                        if CTL_mismatch:
                            act_data_dict = CMDCTLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict,
                                                           False)
                        if DLL_mismatch:
                            act_data_dict = DQSDLLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict,
                                                           False)
                    if skew_to_target_mismatch:
                        act_data_dict = Skew2TargetMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict,
                                                            target_list_all, act_content)
                    if group_mismatch:
                        # 若无特殊情况 topology_name_list_all为空[]
                        act_data_dict = GroupMismatch(topology_name_list_all, start_sch_list, start_net_list,
                                                      end_sch_list,
                                                      act_data_dict, act_content)

                    act_data_layer_list = []
                    result_dict = dict()
                    check_length_seg_list = list()
                    idx_all = -1
                    segment_result_dict = {}
                    real_segment_list = []
                    # 分段索引
                    segment_org_name_list = [str(x) for x in xrange(1, len(all_width_list[0]))]

                    for id1 in xrange(len(start_sch_list)):
                        idx_all += 1
                        if act_data_dict.get((start_sch_list[id1], start_net_list[id1], end_sch_list[id1])) != None:
                            result_list = list()
                            segment_out_name_list = []
                            act_data = act_data_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]
                            act_data_layer_list.append(act_data)
                            length_dict = dict()
                            count1 = 0
                            ignore_key1 = '$CN%d' % count1

                            group_mismatch_count = 0
                            skew2target_count = 0
                            # # ##print('act_data', act_data)
                            for idx1 in xrange(len(segment_list)):
                                # # ##print('segment_list', segment_list[idx1])
                                seg = segment_list[idx1]

                                layer_change_count = 0
                                cross_over_count = 0

                                # Check Order trace width, connect sch, change layer
                                check_length = False
                                check_connect_sch = False
                                check_layer_change = False
                                check_layer_change_num = 1
                                check_cross_over_num = 1
                                check_cross_over = False
                                # # ##print('trace_width_list', trace_width_list[idx1])

                                # 如果有trace width才去check
                                # # ##print('trace_width_list[idx1]', trace_width_list[idx1])
                                connect_sch_spec = ''
                                #
                                if type(trace_width_list[idx1]) == type([]):
                                    check_length = True
                                    # # ##print('trace_width_list[idx1]', trace_width_list[idx1])
                                    spec_trace_width_list = trace_width_list[idx1]
                                    spec_layer_type_list = layer_list[idx1]

                                    try:
                                        # # ##print('connect_sch_list[idx1]', connect_sch_list[idx1])
                                        if connect_sch_list[idx1] not in ['N', 'NA', 'None', 'All',
                                                                          'Connect Component']:
                                            check_connect_sch = True
                                            connect_sch_spec = connect_sch_list[idx1]
                                    except:
                                        pass

                                    try:  # f7684584
                                        if cross_over_list[idx1] == 'Y':
                                            check_cross_over = True
                                        elif re.findall('^Y\d$', cross_over_list[idx1]) != []:
                                            check_cross_over = True
                                            check_cross_over_num = int(re.findall('^Y(\d)$', cross_over_list[idx1])[0])
                                    except:
                                        pass

                                    try:
                                        if layer_change_list[idx1] == 'Y':
                                            check_layer_change = True
                                        elif re.findall('^Y\d$', layer_change_list[idx1]) != []:
                                            check_layer_change = True
                                            check_layer_change_num = int(
                                                re.findall('^Y(\d)$', layer_change_list[idx1])[0])
                                    except:
                                        pass

                                if check_length:
                                    # # ##print('check_connect_sch', check_connect_sch)
                                    # # ##print('ok1')
                                    check_length_seg_list.append(segment_list[idx1])
                                    # # ##print(check_length_seg_list)
                                    length = 0
                                    segment_part_name = ''
                                    end = False
                                    check_layer_change_wrong = False
                                    check_cross_over_wrong = False
                                    while end is False:
                                        # # ##print(segment_list[idx1])
                                        # # ##print(ignore_segment_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])])
                                        # 没有发现！符号，都为空
                                        if segment_list[idx1] in ignore_segment_dict[
                                            (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                            result_list.append(0)
                                            break
                                        # 不知ignore_key1有何作用
                                        if ignore_key1 in ignore_segment_dict[
                                            (start_sch_list[id1], start_net_list[id1],
                                             end_sch_list[id1])] or act_data == []:
                                            result_list.append('NA')
                                            break
                                        # # ##print(act_data)
                                        for idx2 in xrange(len(act_data)):
                                            # # ##print('ok2')
                                            # # ##print(1, act_data)
                                            # # ##print('act_data', act_data)
                                            act_seg = act_data[idx2]
                                            # # ##print('act_seg', act_seg)
                                            if act_data[idx2].find(':') > -1:
                                                # # ##print('ok3')
                                                # # ##print(act_data[idx2])

                                                # 实际的 trace width 和实际的 layer
                                                trace_width = \
                                                    ['%.3f' % float(x) for x in act_data[idx2].split(':') if
                                                     isfloat(x)][
                                                        0]
                                                # # ##print(4, trace_width)
                                                layer = [x for x in act_data[idx2].split(':') if x in All_Layer_List][0]
                                                layer_type = layer_type_dict[layer]

                                                if trace_width in spec_trace_width_list and layer_type in spec_layer_type_list:
                                                    length += float('%.3f' % act_data[idx2 + 1])
                                                    try:
                                                        segment_part_name += segment_org_name_list[idx2]
                                                        # segment_org_name_list = segment_org_name_list[idx2 + 1:]
                                                    except:
                                                        pass
                                                    act_data = act_data[idx2 + 2::]

                                                    if check_connect_sch:
                                                        if connect_sch_spec == 'Y':  # check any connect sch
                                                            if act_data[idx2].split(':')[-1].find('[') == 0:
                                                                end = True
                                                                connect_sch_spec = ''
                                                                break
                                                        else:  # check specific connect sch with the keyword of sch name
                                                            if connect_sch_spec.find('/') > -1:
                                                                connect_sch_spec_list = connect_sch_spec.split('/')
                                                                for sch_tmp1 in connect_sch_spec_list:
                                                                    if sch_tmp1.find('-') > -1:
                                                                        if '[%s]' % sch_tmp1 == \
                                                                                act_data[idx2].split(':')[-1]:
                                                                            end = True
                                                                            break
                                                                    else:
                                                                        if act_seg.split(':')[-1].find(sch_tmp1) > -1:
                                                                            end = True
                                                                            break
                                                            else:
                                                                if connect_sch_spec.find('-') > -1:
                                                                    if '[%s]' % connect_sch_spec == \
                                                                            act_data[idx2].split(':')[
                                                                                -1]:
                                                                        end = True
                                                                        break
                                                                else:
                                                                    if act_seg.split(':')[-1].find(
                                                                            connect_sch_spec) > -1:
                                                                        end = True
                                                                        break

                                                    if check_layer_change:

                                                        try:
                                                            act_seg_next = act_data[0]

                                                            if str(act_seg_next).find('net$') > -1:
                                                                end = True
                                                                break

                                                            layer_next = \
                                                                [x for x in act_seg_next.split(':') if
                                                                 x in All_Layer_List][0]
                                                            if layer != layer_next:
                                                                layer_change_count += 1
                                                                if layer_change_count == check_layer_change_num:
                                                                    end = True
                                                                    break
                                                        except:
                                                            if act_data != []:
                                                                # 表示抓到下一段net(net name改變), 需示警
                                                                result_list.append('Warning!')
                                                                check_layer_change_wrong = True
                                                                end = True
                                                                break

                                                    if check_cross_over:
                                                        try:
                                                            act_seg_next = act_data[0]
                                                            if 'CROSS$' in act_seg_next.split(':'):  # f7684584
                                                                cross_over_count += 1
                                                                if cross_over_count == check_cross_over_num:
                                                                    end = True
                                                                    break

                                                        except:
                                                            if act_data != []:
                                                                result_list.append('Warning!')
                                                                check_cross_over_wrong = True
                                                                end = True
                                                                break

                                                else:  # trace width not match, jump to next spec of topology segment
                                                    end = True

                                                break
                                            elif act_seg.find('net$') > -1:
                                                end = True
                                                break
                                        if act_data == []:
                                            end = True
                                        if end and not check_layer_change_wrong and not check_cross_over_wrong:
                                            length_dict[seg] = float('%.3f' % length)
                                            result_list.append(float('%.3f' % length))
                                            if segment_part_name != '':
                                                segment_out_name_list.append(segment_part_name)
                                else:

                                    if segment_list[idx1] == 'Component Name':
                                        count1 += 1
                                        ignore_key1 = '$CN%d' % count1

                                        if len(act_data) > 0 and ignore_key1 not in ignore_segment_dict[
                                            (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] and act_data[
                                            0].find(
                                            'net$') > -1:

                                            act_data_1 = copy.deepcopy(act_data[1])
                                            # 防止取到cross
                                            if act_data_1.find(u'CROSS$') > -1:
                                                act_data_1 = act_data_1[7:]
                                            result_list.append(act_data_1.split(':')[0][1:-1])
                                            # # ##print(1, result_list)
                                            result_list.append(act_data[0][4::])
                                            # # ##print(2, result_list)
                                            act_data = act_data[1::]
                                        else:
                                            result_list.append('NA')
                                            result_list.append('NA')

                                    elif seg == 'Total Length':
                                        result_list.append(act_data_dict[(
                                            (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                            'total_length')])

                                    elif segment_list[idx1] == 'Via Count':
                                        result_list.append(act_data_dict[(
                                            (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                            'via')].split()[1])
                                    elif re.findall(r'\+|\-|\*|\/', seg) != []:
                                        if seg not in ignore_segment_dict[
                                            (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                            sub_seg_list = re.split(r'\+|\-|\*|\/|\(|\)', segment_list[idx1])
                                            sub_seg_list = [x for x in sub_seg_list if x != '']
                                            seg_tmp = str(seg)

                                            sub_seg_len_list = []
                                            for x in sub_seg_list:
                                                sub_seg_len_list.append(length_dict.get(x, 0))

                                            for id_sub in xrange(len(sub_seg_list)):
                                                try:
                                                    float(sub_seg_list[id_sub])
                                                except ValueError:
                                                    seg_tmp = seg_tmp.replace(sub_seg_list[id_sub],
                                                                              str(float(sub_seg_len_list[id_sub])))

                                            result_list.append(eval(seg_tmp))

                                        else:
                                            result_list.append('NA')
                                    elif seg == 'Total Mismatch':
                                        result_list.append(act_data_dict[(
                                            (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                            'total_mismatch')])
                                    elif segment_list[idx1] == 'Layer Mismatch':
                                        result_list.append(act_data_dict[(
                                            (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                            'layer_mismatch')])
                                    elif segment_list[idx1] == 'Skew to Bundle':
                                        result_list.append(act_data_dict[(
                                            (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                            'bundle_mismatch')])
                                    elif seg == 'Skew to Target':

                                        if '$SK2T' not in ignore_segment_dict[
                                            (start_sch_list[id1], start_net_list[id1],
                                             end_sch_list[id1])] and '$SK2T%d' % (
                                                skew2target_count + 1) not in ignore_segment_dict[
                                            (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                            if target_list_all != []:
                                                result_list.append(act_data_dict[(
                                                    (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                                    'skew2target%d' % skew2target_count)])
                                            else:
                                                result_list.append('NA')
                                        else:
                                            result_list.append('NA')

                                        skew2target_count += 1

                                    elif segment_list[idx1] == 'Group Mismatch':
                                        if '$GM' not in ignore_segment_dict[
                                            (start_sch_list[id1], start_net_list[id1],
                                             end_sch_list[id1])] and '$GM%d' % (
                                                group_mismatch_count + 1) not in ignore_segment_dict[
                                            (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                            if topology_name_list_all != []:
                                                result_list.append(act_data_dict[(
                                                    (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                                    'group_mismatch%d' % group_mismatch_count)])
                                            else:
                                                result_list.append('NA')
                                        else:
                                            result_list.append('NA')
                                        group_mismatch_count += 1

                                    elif segment_list[idx1] == 'Relative Length Spec(DQS to DQS)':
                                        result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                           end_sch_list[id1]), 'DQS_TO_DQS')])
                                    elif segment_list[idx1] == 'Relative Length Spec(DQ to DQ)':
                                        result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                           end_sch_list[id1]), 'DQ_TO_DQ')])
                                    elif segment_list[idx1] == 'Relative Length Spec(DQ to DQS)':
                                        result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                           end_sch_list[id1]), 'DQ_TO_DQS')])
                                    elif segment_list[idx1] == 'CMD or ADD to CLK Length Matching':
                                        result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                           end_sch_list[id1]), 'CMD_TO_CLK')])
                                    elif segment_list[idx1] == 'CTL to CLK Length Matching':
                                        result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                           end_sch_list[id1]), 'CTL_TO_CLK')])
                                    elif segment_list[idx1] == 'DLL Group Length Matching':
                                        result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                           end_sch_list[id1]), 'DLL_Group')])
                                    elif segment_list[idx1] == 'Result':
                                        cal_title_list = ['Total Mismatch', 'Layer Mismatch', 'Total Length',
                                                          'Via Count',
                                                          'Result',
                                                          'Segment Mismatch', 'Skew to Bundle', 'Skew to Target',
                                                          'Group Mismatch',
                                                          'Relative Length Spec(DQS to DQS)',
                                                          'Relative Length Spec(DQ to DQ)',
                                                          'Relative Length Spec(DQ to DQS)',
                                                          'CMD or ADD to CLK Length Matching',
                                                          'CTL to CLK Length Matching', 'DLL Group Length Matching']

                                        total_length_tmp = [result_list[idx_tmp] for idx_tmp in
                                                            xrange(len(segment_list)) if
                                                            segment_list[idx_tmp].find('+') == -1 and segment_list[
                                                                idx_tmp] not in cal_title_list]
                                        total_length_tmp = sum([float(x) for x in total_length_tmp if isfloat(x)])

                                        if abs(float(
                                                act_data_dict[
                                                    ((start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                                     'total_length')]) - total_length_tmp) < 0.01:
                                            result_list.append('Setting OK')
                                        else:
                                            result_list.append('Something Wrong')
                                    elif seg in ['Segment Name', 'Start Segment Name']:
                                        pass
                                    else:
                                        result_list.append('NA')
                            result_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] = result_list
                        else:
                            result_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] = ['NA'] * len(
                                segment_list)

                    if segment_mismatch:
                        result_dict, color_ind_list = SegmentMismatch(start_sch_list, start_net_list, end_sch_list,
                                                                      result_dict,
                                                                      all_width_list, all_layer_list, all_result_list,
                                                                      segment_out_name_list)

                    idx1 = -1

                    for idb in xrange(len(start_sch_list)):
                        idx1 += 1
                        SetCellBorder(active_sheet, (result_ind2[0] + idx1, result_ind2[1] - 3))
                        SetCellBorder(active_sheet, (result_ind2[0] + idx1, result_ind2[1] - 2))
                        SetCellBorder(active_sheet, (result_ind2[0] + idx1, result_ind2[1] - 1))
                        if result_dict.get((start_sch_list[idb], start_net_list[idb], end_sch_list[idb])) != None:
                            for idx2 in xrange(
                                    len(result_dict[(start_sch_list[idb], start_net_list[idb], end_sch_list[idb])])):
                                active_sheet.range((result_ind2[0] + idx1, result_ind2[1] + idx2)).value = \
                                    result_dict[(start_sch_list[idb], start_net_list[idb], end_sch_list[idb])][idx2]
                    if segment_mismatch:
                        for ind in xrange(len(start_sch_list) / 2):
                            # 判断颜色
                            try:
                                if color_ind_list[ind] == 0:
                                    active_sheet.range(
                                        (start_ind[0] + 12 + ind * 2,
                                         start_ind[1] + se1_ind)).api.Interior.ColorIndex = 4
                                    active_sheet.range(
                                        (start_ind[0] + 13 + ind * 2,
                                         start_ind[1] + se1_ind)).api.Interior.ColorIndex = 4
                                elif color_ind_list[ind] == 1:
                                    active_sheet.range(
                                        (start_ind[0] + 12 + ind * 2,
                                         start_ind[1] + se1_ind)).api.Interior.ColorIndex = 3
                                    active_sheet.range(
                                        (start_ind[0] + 13 + ind * 2,
                                         start_ind[1] + se1_ind)).api.Interior.ColorIndex = 3
                            except:
                                pass

                    for idxx in xrange(len(topology_table)):
                        row_tmp = topology_table[idxx]
                        if 'Start Segment Name' in row_tmp:
                            segment_list = topology_table[idxx][3::]
                            segment_list = [x for x in segment_list if x not in ['', None]]
                            result_idx = segment_list.index('Result')
                        elif 'Min' in topology_table[idxx]:
                            spec_min_list = [str(x) for x in topology_table[idxx][3::]]
                            spec_min_list = spec_min_list[0:result_idx + 1]
                        elif 'Max' in row_tmp:
                            spec_max_list = [str(x) for x in topology_table[idxx][3::]]
                            spec_max_list = spec_max_list[0:result_idx + 1]
                    len1 = len(active_sheet.range(result_ind1).options(expand='table', ndim=2).value)

                    for idx1 in xrange(len1):
                        result_cell_idx = (result_ind2[0] + idx1, result_ind2[1] + result_idx)
                        if active_sheet.range(result_cell_idx).value != 'Something Wrong':
                            active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 4
                            active_sheet.range(result_cell_idx).value = 'Pass'
                        elif active_sheet.range(result_cell_idx).value == 'Something Wrong':
                            active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
                    # for idx1 in xrange(int(len1 / 2)):
                    #     result_cell_idx1 = (result_ind2[0] + idx1 * 2, result_ind2[1] + result_idx)
                    #     result_cell_idx2 = (result_ind2[0] + idx1 * 2 + 1, result_ind2[1] + result_idx)
                    #     if active_sheet.range(result_cell_idx1).value != 'Something Wrong' and \
                    #             active_sheet.range(result_cell_idx2).value != 'Something Wrong':
                    #         if color_ind_list != []:
                    #             if color_ind_list[idx1] == 1:
                    #                 active_sheet.range(result_cell_idx1).value = 'Fail'
                    #                 active_sheet.range(result_cell_idx2).value = 'Fail'
                    #                 active_sheet.range(result_cell_idx1).api.Interior.ColorIndex = 3
                    #                 active_sheet.range(result_cell_idx2).api.Interior.ColorIndex = 3

                    for idx1 in xrange(len(segment_list)):
                        for idx2 in xrange(len1):
                            spec_min = spec_min_list[idx1]
                            spec_max = spec_max_list[idx1]
                            result_cell_idx = (result_ind2[0] + idx2, result_ind2[1] + result_idx)
                            if active_sheet.range(result_cell_idx).value != 'Something Wrong':
                                if isfloat(active_sheet.range((result_ind2[0] + idx2, result_ind2[1] + idx1)).value):
                                    if spec_min == 'NA' and isfloat(spec_max):
                                        if float(active_sheet.range(
                                                (result_ind2[0] + idx2, result_ind2[1] + idx1)).value) <= float(
                                            spec_max):
                                            active_sheet.range(
                                                (result_ind2[0] + idx2,
                                                 result_ind2[1] + idx1)).api.Interior.ColorIndex = 4
                                        else:
                                            active_sheet.range(
                                                (result_ind2[0] + idx2,
                                                 result_ind2[1] + idx1)).api.Interior.ColorIndex = 3
                                            active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
                                            active_sheet.range(result_cell_idx).value = 'Fail'
                                    elif isfloat(spec_min) and isfloat(spec_max):
                                        if float(spec_min) <= float(
                                                active_sheet.range(
                                                    (result_ind2[0] + idx2, result_ind2[1] + idx1)).value) <= float(
                                            spec_max):
                                            active_sheet.range(
                                                (result_ind2[0] + idx2,
                                                 result_ind2[1] + idx1)).api.Interior.ColorIndex = 4
                                        else:
                                            active_sheet.range(
                                                (result_ind2[0] + idx2,
                                                 result_ind2[1] + idx1)).api.Interior.ColorIndex = 3
                                            active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
                                            active_sheet.range(result_cell_idx).value = 'Fail'
                                    elif isfloat(spec_min) and spec_max == 'NA':
                                        if float(spec_min) <= float(
                                                active_sheet.range(
                                                    (result_ind2[0] + idx2, result_ind2[1] + idx1)).value):
                                            active_sheet.range(
                                                (result_ind2[0] + idx2,
                                                 result_ind2[1] + idx1)).api.Interior.ColorIndex = 4
                                        else:
                                            active_sheet.range(
                                                (result_ind2[0] + idx2,
                                                 result_ind2[1] + idx1)).api.Interior.ColorIndex = 3
                                            active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
                                            active_sheet.range(result_cell_idx).value = 'Fail'

                    SetCellFont_current_region(active_sheet, cell_idx, 'Times New Roman', 12, 'l')
                    SetCellBorder_current_region(active_sheet, cell_idx)
                    active_sheet.autofit('c')
            else:
                topology_table = active_sheet.range(cell_idx).current_region.value
                for idx1 in xrange(len(topology_table)):
                    line = topology_table[idx1]
                    # print(line)
                    for idx2 in xrange(len(line)):
                        if line[idx2] == 'Net Name':
                            spec_table_length = idx1 + 1
                            break
                start_sch_ind = (cell_idx[0] + spec_table_length, cell_idx[1])
                end_sch_ind = (start_sch_ind[0], start_sch_ind[1] + 1)
                start_net_ind = (end_sch_ind[0], end_sch_ind[1] + 1)
                result_ind2 = (start_net_ind[0], start_net_ind[1] + 1)
                active_sheet.range(result_ind2).expand('table').clear()

                for se_ind in xrange(len(topology_table[2])):
                    # 获取Segment Mismatch的管控条件
                    if topology_table[2][se_ind] and topology_table[2][se_ind].find('Segment Mismatch') > -1:
                        se1_ind = se_ind
                for idx1 in xrange(len(topology_table)):
                    line = topology_table[idx1]
                    for idx2 in xrange(len(line)):
                        if line[idx2] == 'Net Name':
                            spec_table_length = idx1 + 1
                            break
                # ##print('spec_table_length',spec_table_length)

                # 获取坐标
                start_sch_ind = (cell_idx[0] + spec_table_length, cell_idx[1])
                end_sch_ind = (start_sch_ind[0], start_sch_ind[1] + 1)
                start_net_ind = (end_sch_ind[0], end_sch_ind[1] + 1)
                result_ind1 = (start_net_ind[0], start_net_ind[1] - 2)
                result_ind2 = (start_net_ind[0], start_net_ind[1] + 1)

                for idxx in xrange(len(topology_table)):
                    # print(topology_table)
                    row_tmp = topology_table[idxx]
                    if 'Start Segment Name' in row_tmp:
                        segment_list = topology_table[idxx][3::]
                        segment_list = [x for x in segment_list if x not in ['', None]]
                        result_idx = segment_list.index('Result')
                        # ##print(result_idx)
                    elif 'Layer' in topology_table[idxx]:
                        layer_list = []
                        for x in topology_table[idxx][3::]:
                            layer_list.append(str(x))
                        layer_list = layer_list[0:result_idx + 1]
                        layer_list = [x.split('/') for x in layer_list]
                        # ##print(layer_list)
                    elif 'Layer Change' in topology_table[idxx]:
                        layer_change_list = []
                        for x in topology_table[idxx][3::]:
                            layer_change_list.append(str(x))
                        layer_change_list = layer_change_list[0:result_idx + 1]
                    elif 'Connect Component' in row_tmp:
                        connect_sch_list = []
                        for x in topology_table[idxx][3::]:
                            connect_sch_list.append(str(x))
                        connect_sch_list = connect_sch_list[0:result_idx + 1]
                    elif 'Cross Over' in row_tmp:
                        cross_over_list = []
                        for x in topology_table[idxx][3::]:
                            cross_over_list.append(str(x))
                        cross_over_list = cross_over_list[0:result_idx + 1]
                    elif 'Trace Width' in topology_table[idxx]:
                        trace_width_list = []
                        for x in topology_table[idxx][3::]:
                            trace_width_list.append(str(x))
                        trace_width_list = trace_width_list[0:result_idx + 1]
                        for idx in xrange(len(trace_width_list)):
                            # 得到的数都为浮点数
                            tw = trace_width_list[idx]
                            # 数据为浮点数时，如3.5，或者有 / 时（差分信号）
                            if isfloat(trace_width_list[idx]) or tw.find('/') > -1:
                                trace_width_list[idx] = ['%.3f' % float(x) for x in trace_width_list[idx].split('/')]
                    elif 'Space' in row_tmp:
                        space_list = []
                        for x in topology_table[idxx][3::]:
                            space_list.append(str(x))
                        space_list = space_list[0:result_idx + 1]
                    elif 'Min' in row_tmp:
                        spec_min_list = []
                        for x in topology_table[idxx][3::]:
                            spec_min_list.append(str(x))
                        spec_min_list = spec_min_list[0:result_idx + 1]
                    elif 'Max' in row_tmp:
                        spec_max_list = []
                        for x in topology_table[idxx][3::]:
                            spec_max_list.append(str(x))
                        spec_max_list = spec_max_list[0:result_idx + 1]

                signal_type = topology_table[1][1]
                #####################################################
                table1 = active_sheet.range(result_ind1).expand('table').value
                if type(table1[0]) == type(u''):
                    table1 = [table1]
                ########################################################
                start_sch_list = [tt[0] for tt in table1]
                end_sch_list = [tt[1] for tt in table1]
                start_net_list = [tt[2] for tt in table1]

                # Check ignore syntax to skip specified segment check
                ignore_segment_dict = dict()
                start_net_list = list(start_net_list)
                for idx in xrange(len(start_net_list)):
                    net_name = start_net_list[idx]
                    # !代表什么
                    if net_name.find('!') > -1:
                        ignore_segment_list = start_net_list[idx].split('!')[1::]
                        start_net_list[idx] = start_net_list[idx].split('!')[0]
                        ignore_segment_dict[
                            (start_sch_list[idx], start_net_list[idx], end_sch_list[idx])] = ignore_segment_list
                    else:
                        ignore_segment_dict[(start_sch_list[idx], start_net_list[idx], end_sch_list[idx])] = []

                start_net_list = tuple(start_net_list)

                layer_mismatch, segment_mismatch, bundle_mismatch, total_mismatch, DQS_TO_DQS_mismatch \
                    = False, False, False, False, False
                DQ_TO_DQ_mismatch, DQ_TO_DQS_mismatch, CMD_mismatch, CTL_mismatch, DLL_mismatch \
                    = False, False, False, False, False
                if signal_type == 'Differential':
                    act_sheet = wb.sheets['ACT_diff']
                    if 'Layer Mismatch' in segment_list:
                        layer_mismatch = True
                    if 'Segment Mismatch' in segment_list:
                        segment_mismatch = True
                    if 'Skew to Bundle' in segment_list:
                        bundle_mismatch = True
                    if 'Total Mismatch' in segment_list:
                        total_mismatch = True
                    if 'Relative Length Spec(DQS to DQS)' in segment_list:
                        DQS_TO_DQS_mismatch = True
                elif signal_type == 'Single-ended':
                    act_sheet = wb.sheets['ACT_se']
                    if 'Relative Length Spec(DQ to DQ)' in segment_list:
                        DQ_TO_DQ_mismatch = True
                    if 'Relative Length Spec(DQ to DQS)' in segment_list:
                        DQ_TO_DQS_mismatch = True
                    if 'CMD or ADD to CLK Length Matching' in segment_list:
                        CMD_mismatch = True
                    if 'CTL to CLK Length Matching' in segment_list:
                        CTL_mismatch = True
                    if 'DLL Group Length Matching' in segment_list:
                        DLL_mismatch = True

                act_content = act_sheet.range('A1').current_region.value  # net name 所有信息
                # print(act_content)
                act_data_dict = dict()
                stub_value = []
                for idx123 in xrange(len(start_sch_list)):
                    for line in act_content:
                        if start_sch_list[idx123] == line[0] and end_sch_list[idx123] == line[2] and start_net_list[
                            idx123] == line[1]:
                            if act_data_dict.get(
                                    (start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123])) == None:
                                # ##print((start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123]))
                                act_data_dict[
                                    ((start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123]), 'via')] = \
                                    line[3]
                                act_data_dict[
                                    ((start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123]),
                                     'total_length')] = line[4].split()[1]
                                act_data_dict[
                                    (start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123])] = [x for x
                                                                                                               in
                                                                                                               line[5::]
                                                                                                               if
                                                                                                               x not in [
                                                                                                                   '',
                                                                                                                   None]]

                                # 通过过孔via数目判断是否有stub
                                # 对data数据进行处理
                                net_data = act_data_dict[
                                    (start_sch_list[idx123], start_net_list[idx123], end_sch_list[idx123])]

                            if topology_table[0][1] and topology_table[0][1].upper() in ['LOW LOSS', 'MID LOSS']:
                                if int(line[3].split()[1]) > 0:
                                    # # ##print(topology_table[0][1])
                                    stub_layer_list = []
                                    for idx in range(len(net_data)):
                                        item = net_data[idx]
                                        if str(item).find(":") > -1:
                                            stub_layer_list += [x for x in item.split(':') if x in All_Layer_List]
                                    index_list = []
                                    for i in range(len(stub_layer_list)):
                                        for j in range(len(layerLength_list)):
                                            if stub_layer_list[i] == layerLength_list[j].keys()[0]:
                                                # # ##print(j)  #具体走线层匹配的索引位置
                                                # havingLength_layer_list.append({stub_layer_list[i]:layerLength_list[j].values()[0]}) #给对应层匹配相应的长度
                                                index_list.append(j)
                                    new_index_list = []
                                    new_index_list = sorted(set(index_list))
                                    top_stub = bottom_stub = 0

                                    try:
                                        for top_idx in range(1, new_index_list[-2]):
                                            # print(layerLength_list[top_idx].values()[0])
                                            top_stub += float(layerLength_list[top_idx].values()[0])
                                    except:
                                        pass

                                    try:
                                        for bot_idx in range(new_index_list[1] + 1, len(layerLength_list) - 1):
                                            bottom_stub += float(layerLength_list[bot_idx].values()[0])
                                    except:
                                        pass

                                    final_stub = round(max(top_stub, bottom_stub), 2)
                                    stub_value.append(final_stub)
                                else:
                                    pass
                # print('act_data_dict',act_data_dict)
                if topology_table[0][1] and topology_table[0][1].upper() in ['LOW LOSS', 'MID LOSS']:
                    # 获得当前的sheet名称
                    active_sheet = wb.sheets.active
                    PCIE_sheet = wb.sheets[active_sheet]
                    for cell in PCIE_sheet.api.UsedRange.Cells:
                        if cell.Value == 'Stub':
                            Stub_ind = (cell.Row, cell.Column)
                            break
                    stub_table = active_sheet.range(Stub_ind).current_region.value
                    # # ##print(stub_table)
                    for idx in range(len(topology_table[2])):
                        if topology_table[2][idx] == 'Total Length':
                            len_idx = idx
                            break
                    if topology_table[0][1].upper() == 'LOW LOSS':
                        if max(stub_value) > stub_table[1][0]:
                            active_sheet.range(cell.Row[0] + 10, cell.Column[1] + len_idx).value = stub_table[2][2]
                        else:
                            active_sheet.range(cell.Row[0] + 10, cell.Column[1] + len_idx).value = stub_table[1][2]

                    elif topology_table[0][1].upper() == 'MID LOSS':
                        if max(stub_value) > stub_table[1][0] and max(stub_value) < stub_table[2][0]:
                            active_sheet.range(cell.Row[0] + 10, cell.Column[1] + len_idx).value = stub_table[2][1]
                        else:
                            active_sheet.range(cell.Row[0] + 10, cell.Column[1] + len_idx).value = stub_table[1][1]
                else:
                    pass

                skew_to_target_mismatch = False
                segment_list = list(segment_list)
                target_list_all = list()
                # 对Skew to Target与Group Mismatch的情况进行讨论
                for idx in xrange(len(segment_list)):
                    seg = segment_list[idx]
                    if seg.find('Skew to Target') > -1:
                        skew_to_target_mismatch = True
                        target_list_all.append(segment_list[idx].split(':')[1::])
                        target_list_all[-1] = [x[1:-1].split(',') for x in target_list_all[-1]]
                        segment_list[idx] = 'Skew to Target'
                segment_list = tuple(segment_list)
                # Get the list of "Group Mismatch Target"
                group_mismatch = False
                topology_name_list_all = list()
                segment_list = list(segment_list)
                for idx in xrange(len(segment_list)):
                    if segment_list[idx].find('Group Mismatch') > -1:
                        group_mismatch = True
                        topology_name_list_all.append(segment_list[idx].split(':')[1::])
                        # # ##print(topology_name_list_all)
                        segment_list[idx] = 'Group Mismatch'
                segment_list = tuple(segment_list)
                # ##print(segment_list)

                act_value_list = []
                all_result_list = []
                all_width_list = []
                all_layer_list = []
                for ind1 in xrange(len(start_sch_list)):
                    if act_data_dict.get((start_sch_list[ind1], start_net_list[ind1], end_sch_list[ind1])):
                        act_value_list.append(
                            act_data_dict[(start_sch_list[ind1], start_net_list[ind1], end_sch_list[ind1])])
                for ind1 in xrange(len(act_value_list)):
                    half_result_list = []
                    half_width_list = []
                    half_layer_list = []
                    del_ind_list = []
                    act_value = act_value_list[ind1]

                    for ind2 in xrange(len(act_value)):
                        if str(act_value[ind2]).find(':') > -1 and isfloat(str(act_value[ind2]).split(':')[-1]) \
                                or str(act_value[ind2]).find('net$') > -1:
                            half_width_list.append(str(act_value[ind2]).split(':')[-1])

                        if str(act_value[ind2]).find(':') > -1:
                            for x in act_value[ind2].split(':'):
                                if x in All_Layer_List:
                                    half_layer_list.append(x)

                        if str(act_value[ind2]).find('net$') > -1:
                            half_layer_list.append(act_value[ind2])

                        if str(act_value[ind2]).find(':') > -1 and isfloat(str(act_value[ind2]).split(':')[-2]):
                            half_width_list.append(str(act_value[ind2]).split(':')[-2])

                        if str(act_value[ind2]).find(':') == -1:
                            half_result_list.append(act_value[ind2])
                    all_result_list.append(half_result_list)
                    all_width_list.append(half_width_list)
                    all_layer_list.append(half_layer_list)

                if signal_type == 'Differential':
                    if total_mismatch:
                        act_data_dict = TotalMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                    if layer_mismatch:
                        act_data_dict = LayerMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                    if bundle_mismatch:
                        act_data_dict = Skew2BundleMismatch(4, start_sch_list, start_net_list, end_sch_list,
                                                            act_data_dict)
                    if DQS_TO_DQS_mismatch:
                        act_data_dict = DQSDLLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                    # if amd_lother_cap:
                    #     act_data_dict = ForAMD_LOther(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                    # if amd_lother_via:
                    #     act_data_dict = ForAMD_LOther(start_sch_list, start_net_list, end_sch_list, act_data_dict, CAP_type = 'PTH', Length_Type = 'Pin2Pin')
                else:
                    if DQ_TO_DQ_mismatch:
                        act_data_dict = DQTODQMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                    if DQ_TO_DQS_mismatch:
                        act_data_dict = DQTODQSMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                    if CMD_mismatch:
                        act_data_dict = CMDCTLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict)
                    if CTL_mismatch:
                        act_data_dict = CMDCTLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict,
                                                       False)
                    if DLL_mismatch:
                        act_data_dict = DQSDLLMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict,
                                                       False)
                if skew_to_target_mismatch:
                    act_data_dict = Skew2TargetMismatch(start_sch_list, start_net_list, end_sch_list, act_data_dict,
                                                        target_list_all, act_content)
                if group_mismatch:
                    # 若无特殊情况 topology_name_list_all为空[]
                    act_data_dict = GroupMismatch(topology_name_list_all, start_sch_list, start_net_list, end_sch_list,
                                                  act_data_dict, act_content)

                act_data_layer_list = []
                result_dict = dict()
                check_length_seg_list = list()
                idx_all = -1
                segment_result_dict = {}
                real_segment_list = []
                # 分段索引
                segment_org_name_list = [str(x) for x in xrange(1, len(all_width_list[0]))]

                for id1 in xrange(len(start_sch_list)):
                    idx_all += 1
                    if act_data_dict.get((start_sch_list[id1], start_net_list[id1], end_sch_list[id1])) != None:
                        result_list = list()
                        segment_out_name_list = []
                        act_data = act_data_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]
                        act_data_layer_list.append(act_data)
                        length_dict = dict()
                        count1 = 0
                        ignore_key1 = '$CN%d' % count1

                        group_mismatch_count = 0
                        skew2target_count = 0
                        # # ##print('act_data', act_data)
                        for idx1 in xrange(len(segment_list)):
                            # # ##print('segment_list', segment_list[idx1])
                            seg = segment_list[idx1]

                            layer_change_count = 0

                            # Check Order trace width, connect sch, change layer
                            check_length = False
                            check_connect_sch = False
                            check_layer_change = False
                            check_layer_change_num = 1
                            check_cross_over = False
                            # # ##print('trace_width_list', trace_width_list[idx1])

                            # 如果有trace width才去check
                            # # ##print('trace_width_list[idx1]', trace_width_list[idx1])
                            connect_sch_spec = ''
                            #
                            if type(trace_width_list[idx1]) == type([]):
                                check_length = True
                                # # ##print('trace_width_list[idx1]', trace_width_list[idx1])
                                spec_trace_width_list = trace_width_list[idx1]
                                spec_layer_type_list = layer_list[idx1]

                                try:
                                    # # ##print('connect_sch_list[idx1]', connect_sch_list[idx1])
                                    if connect_sch_list[idx1] not in ['N', 'NA', 'None', 'All', 'Connect Component']:
                                        check_connect_sch = True
                                        connect_sch_spec = connect_sch_list[idx1]
                                except:
                                    pass

                                try:
                                    if layer_change_list[idx1] == 'Y':
                                        check_layer_change = True
                                    elif re.findall('^Y\d$', layer_change_list[idx1]) != []:
                                        check_layer_change = True
                                        check_layer_change_num = int(re.findall('^Y(\d)$', layer_change_list[idx1])[0])
                                except:
                                    pass

                            if check_length:
                                # # ##print('check_connect_sch', check_connect_sch)
                                # # ##print('ok1')
                                check_length_seg_list.append(segment_list[idx1])
                                # # ##print(check_length_seg_list)
                                length = 0
                                segment_part_name = ''
                                end = False
                                check_layer_change_wrong = False
                                check_cross_over_wrong = False
                                while end is False:
                                    # # ##print(segment_list[idx1])
                                    # # ##print(ignore_segment_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])])
                                    # 没有发现！符号，都为空
                                    if segment_list[idx1] in ignore_segment_dict[
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                        result_list.append(0)
                                        break
                                    # 不知ignore_key1有何作用
                                    if ignore_key1 in ignore_segment_dict[
                                        (
                                                start_sch_list[id1], start_net_list[id1],
                                                end_sch_list[id1])] or act_data == []:
                                        result_list.append('NA')
                                        break
                                    # # ##print(act_data)
                                    for idx2 in xrange(len(act_data)):
                                        # # ##print('ok2')
                                        # # ##print(1, act_data)
                                        # # ##print('act_data', act_data)
                                        act_seg = act_data[idx2]
                                        # # ##print('act_seg', act_seg)
                                        if act_data[idx2].find(':') > -1:
                                            # # ##print('ok3')
                                            # # ##print(act_data[idx2])

                                            # 实际的 trace width 和实际的 layer
                                            trace_width = \
                                                ['%.3f' % float(x) for x in act_data[idx2].split(':') if isfloat(x)][
                                                    0]
                                            # # ##print(4, trace_width)
                                            layer = [x for x in act_data[idx2].split(':') if x in All_Layer_List][0]
                                            layer_type = layer_type_dict[layer]

                                            if trace_width in spec_trace_width_list and layer_type in spec_layer_type_list:
                                                length += float('%.3f' % act_data[idx2 + 1])
                                                try:
                                                    segment_part_name += segment_org_name_list[idx2]
                                                    # segment_org_name_list = segment_org_name_list[idx2 + 1:]
                                                except:
                                                    pass
                                                act_data = act_data[idx2 + 2::]

                                                if check_connect_sch:
                                                    if connect_sch_spec == 'Y':  # check any connect sch
                                                        if act_data[idx2].split(':')[-1].find('[') == 0:
                                                            end = True
                                                            connect_sch_spec = ''
                                                            break
                                                    else:  # check specific connect sch with the keyword of sch name
                                                        if connect_sch_spec.find('/') > -1:
                                                            connect_sch_spec_list = connect_sch_spec.split('/')
                                                            for sch_tmp1 in connect_sch_spec_list:
                                                                if sch_tmp1.find('-') > -1:
                                                                    if '[%s]' % sch_tmp1 == act_data[idx2].split(':')[
                                                                        -1]:
                                                                        end = True
                                                                        break
                                                                else:
                                                                    if act_seg.split(':')[-1].find(sch_tmp1) > -1:
                                                                        end = True
                                                                        break
                                                        else:
                                                            if connect_sch_spec.find('-') > -1:
                                                                if '[%s]' % connect_sch_spec == \
                                                                        act_data[idx2].split(':')[
                                                                            -1]:
                                                                    end = True
                                                                    break
                                                            else:
                                                                if act_seg.split(':')[-1].find(connect_sch_spec) > -1:
                                                                    end = True
                                                                    break

                                                if check_layer_change:

                                                    try:
                                                        act_seg_next = act_data[0]

                                                        if str(act_seg_next).find('net$') > -1:
                                                            end = True
                                                            break

                                                        layer_next = \
                                                            [x for x in act_seg_next.split(':') if x in All_Layer_List][
                                                                0]
                                                        if layer != layer_next:
                                                            layer_change_count += 1
                                                            if layer_change_count == check_layer_change_num:
                                                                end = True
                                                                break
                                                    except:
                                                        if act_data != []:
                                                            # 表示抓到下一段net(net name改變), 需示警
                                                            result_list.append('Warning!')
                                                            check_layer_change_wrong = True
                                                            end = True
                                                            break

                                                if check_cross_over:
                                                    try:
                                                        act_seg_next = act_data[0]

                                                        if str(act_seg_next).find('net$') > -1:
                                                            end = True
                                                            break

                                                    except:
                                                        if act_data != []:
                                                            result_list.append('Warning!')
                                                            check_cross_over_wrong = True
                                                            end = True
                                                            break

                                            else:  # trace width not match, jump to next spec of topology segment
                                                end = True

                                            break
                                        elif act_seg.find('net$') > -1:
                                            end = True
                                            break
                                    if act_data == []:
                                        end = True
                                    if end and not check_layer_change_wrong and not check_cross_over_wrong:
                                        length_dict[seg] = float('%.3f' % length)
                                        result_list.append(float('%.3f' % length))
                                        if segment_part_name != '':
                                            segment_out_name_list.append(segment_part_name)
                            else:

                                if segment_list[idx1] == 'Component Name':
                                    count1 += 1
                                    ignore_key1 = '$CN%d' % count1

                                    if len(act_data) > 0 and ignore_key1 not in ignore_segment_dict[
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] and act_data[
                                        0].find(
                                        'net$') > -1:
                                        act_data_1 = copy.deepcopy(act_data[1])
                                        # 防止取到cross
                                        if act_data_1.find(u'CROSS$') > -1:
                                            act_data_1 = act_data_1[7:]
                                        result_list.append(act_data_1.split(':')[0][1:-1])
                                        # # ##print(1, result_list)
                                        result_list.append(act_data[0][4::])
                                        # # ##print(2, result_list)
                                        act_data = act_data[1::]
                                    else:
                                        result_list.append('NA')
                                        result_list.append('NA')

                                elif seg == 'Total Length':
                                    result_list.append(act_data_dict[(
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]), 'total_length')])

                                elif segment_list[idx1] == 'Via Count':
                                    result_list.append(act_data_dict[(
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]), 'via')].split()[
                                                           1])
                                elif re.findall(r'\+|\-|\*|\/', seg) != []:
                                    if seg not in ignore_segment_dict[
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                        sub_seg_list = re.split(r'\+|\-|\*|\/|\(|\)', segment_list[idx1])
                                        sub_seg_list = [x for x in sub_seg_list if x != '']
                                        seg_tmp = str(seg)

                                        sub_seg_len_list = []
                                        for x in sub_seg_list:
                                            sub_seg_len_list.append(length_dict.get(x, 0))

                                        for id_sub in xrange(len(sub_seg_list)):
                                            try:
                                                float(sub_seg_list[id_sub])
                                            except ValueError:
                                                seg_tmp = seg_tmp.replace(sub_seg_list[id_sub],
                                                                          str(float(sub_seg_len_list[id_sub])))

                                        result_list.append(eval(seg_tmp))

                                    else:
                                        result_list.append('NA')
                                elif seg == 'Total Mismatch':
                                    result_list.append(act_data_dict[(
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                        'total_mismatch')])
                                elif segment_list[idx1] == 'Layer Mismatch':
                                    result_list.append(act_data_dict[(
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                        'layer_mismatch')])
                                elif segment_list[idx1] == 'Skew to Bundle':
                                    result_list.append(act_data_dict[(
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                        'bundle_mismatch')])
                                elif seg == 'Skew to Target':

                                    if '$SK2T' not in ignore_segment_dict[
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] and '$SK2T%d' % (
                                            skew2target_count + 1) not in ignore_segment_dict[
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                        if target_list_all != []:
                                            result_list.append(act_data_dict[(
                                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                                'skew2target%d' % skew2target_count)])
                                        else:
                                            result_list.append('NA')
                                    else:
                                        result_list.append('NA')

                                    skew2target_count += 1

                                elif segment_list[idx1] == 'Group Mismatch':
                                    if '$GM' not in ignore_segment_dict[
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] and '$GM%d' % (
                                            group_mismatch_count + 1) not in ignore_segment_dict[
                                        (start_sch_list[id1], start_net_list[id1], end_sch_list[id1])]:
                                        if topology_name_list_all != []:
                                            result_list.append(act_data_dict[(
                                                (start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                                'group_mismatch%d' % group_mismatch_count)])
                                        else:
                                            result_list.append('NA')
                                    else:
                                        result_list.append('NA')
                                    group_mismatch_count += 1

                                elif segment_list[idx1] == 'Relative Length Spec(DQS to DQS)':
                                    result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                       end_sch_list[id1]), 'DQS_TO_DQS')])
                                elif segment_list[idx1] == 'Relative Length Spec(DQ to DQ)':
                                    result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                       end_sch_list[id1]), 'DQ_TO_DQ')])
                                elif segment_list[idx1] == 'Relative Length Spec(DQ to DQS)':
                                    result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                       end_sch_list[id1]), 'DQ_TO_DQS')])
                                elif segment_list[idx1] == 'CMD or ADD to CLK Length Matching':
                                    result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                       end_sch_list[id1]), 'CMD_TO_CLK')])
                                elif segment_list[idx1] == 'CTL to CLK Length Matching':
                                    result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                       end_sch_list[id1]), 'CTL_TO_CLK')])
                                elif segment_list[idx1] == 'DLL Group Length Matching':
                                    result_list.append(act_data_dict[((start_sch_list[id1], start_net_list[id1],
                                                                       end_sch_list[id1]), 'DLL_Group')])
                                elif segment_list[idx1] == 'Result':
                                    cal_title_list = ['Total Mismatch', 'Layer Mismatch', 'Total Length', 'Via Count',
                                                      'Result',
                                                      'Segment Mismatch', 'Skew to Bundle', 'Skew to Target',
                                                      'Group Mismatch',
                                                      'Relative Length Spec(DQS to DQS)',
                                                      'Relative Length Spec(DQ to DQ)',
                                                      'Relative Length Spec(DQ to DQS)',
                                                      'CMD or ADD to CLK Length Matching',
                                                      'CTL to CLK Length Matching', 'DLL Group Length Matching']

                                    total_length_tmp = [result_list[idx_tmp] for idx_tmp in xrange(len(segment_list)) if
                                                        segment_list[idx_tmp].find('+') == -1 and segment_list[
                                                            idx_tmp] not in cal_title_list]
                                    total_length_tmp = sum([float(x) for x in total_length_tmp if isfloat(x)])

                                    if abs(float(
                                            act_data_dict[
                                                ((start_sch_list[id1], start_net_list[id1], end_sch_list[id1]),
                                                 'total_length')]) - total_length_tmp) < 0.01:
                                        result_list.append('Setting OK')
                                    else:
                                        result_list.append('Something Wrong')
                                elif seg in ['Segment Name', 'Start Segment Name']:
                                    pass
                                else:
                                    result_list.append('NA')
                        result_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] = result_list
                    else:
                        result_dict[(start_sch_list[id1], start_net_list[id1], end_sch_list[id1])] = ['NA'] * len(
                            segment_list)

                if segment_mismatch:
                    result_dict, color_ind_list = SegmentMismatch(start_sch_list, start_net_list, end_sch_list,
                                                                  result_dict,
                                                                  all_width_list, all_layer_list, all_result_list,
                                                                  segment_out_name_list)

                idx1 = -1

                for idb in xrange(len(start_sch_list)):
                    idx1 += 1
                    SetCellBorder(active_sheet, (result_ind2[0] + idx1, result_ind2[1] - 3))
                    SetCellBorder(active_sheet, (result_ind2[0] + idx1, result_ind2[1] - 2))
                    SetCellBorder(active_sheet, (result_ind2[0] + idx1, result_ind2[1] - 1))
                    if result_dict.get((start_sch_list[idb], start_net_list[idb], end_sch_list[idb])) != None:
                        for idx2 in xrange(
                                len(result_dict[(start_sch_list[idb], start_net_list[idb], end_sch_list[idb])])):
                            active_sheet.range((result_ind2[0] + idx1, result_ind2[1] + idx2)).value = \
                                result_dict[(start_sch_list[idb], start_net_list[idb], end_sch_list[idb])][idx2]
                if segment_mismatch:
                    for ind in xrange(len(start_sch_list) / 2):
                        # 判断颜色
                        try:
                            if color_ind_list[ind] == 0:
                                active_sheet.range(
                                    (start_ind[0] + 12 + ind * 2, start_ind[1] + se1_ind)).api.Interior.ColorIndex = 4
                                active_sheet.range(
                                    (start_ind[0] + 13 + ind * 2, start_ind[1] + se1_ind)).api.Interior.ColorIndex = 4
                            elif color_ind_list[ind] == 1:
                                active_sheet.range(
                                    (start_ind[0] + 12 + ind * 2, start_ind[1] + se1_ind)).api.Interior.ColorIndex = 3
                                active_sheet.range(
                                    (start_ind[0] + 13 + ind * 2, start_ind[1] + se1_ind)).api.Interior.ColorIndex = 3
                        except:
                            pass

                for idxx in xrange(len(topology_table)):
                    row_tmp = topology_table[idxx]
                    if 'Start Segment Name' in row_tmp:
                        segment_list = topology_table[idxx][3::]
                        segment_list = [x for x in segment_list if x not in ['', None]]
                        result_idx = segment_list.index('Result')
                    elif 'Min' in topology_table[idxx]:
                        spec_min_list = [str(x) for x in topology_table[idxx][3::]]
                        spec_min_list = spec_min_list[0:result_idx + 1]
                    elif 'Max' in row_tmp:
                        spec_max_list = [str(x) for x in topology_table[idxx][3::]]
                        spec_max_list = spec_max_list[0:result_idx + 1]
                len1 = len(active_sheet.range(result_ind1).options(expand='table', ndim=2).value)

                for idx1 in xrange(len1):
                    result_cell_idx = (result_ind2[0] + idx1, result_ind2[1] + result_idx)
                    if active_sheet.range(result_cell_idx).value != 'Something Wrong':
                        active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 4
                        active_sheet.range(result_cell_idx).value = 'Pass'
                    elif active_sheet.range(result_cell_idx).value == 'Something Wrong':
                        active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
                # for idx1 in xrange(int(len1 / 2)):
                #     result_cell_idx1 = (result_ind2[0] + idx1 * 2, result_ind2[1] + result_idx)
                #     result_cell_idx2 = (result_ind2[0] + idx1 * 2 + 1, result_ind2[1] + result_idx)
                #     if active_sheet.range(result_cell_idx1).value != 'Something Wrong' and \
                #             active_sheet.range(result_cell_idx2).value != 'Something Wrong':
                #         if color_ind_list != []:
                #             if color_ind_list[idx1] == 1:
                #                 active_sheet.range(result_cell_idx1).value = 'Fail'
                #                 active_sheet.range(result_cell_idx2).value = 'Fail'
                #                 active_sheet.range(result_cell_idx1).api.Interior.ColorIndex = 3
                #                 active_sheet.range(result_cell_idx2).api.Interior.ColorIndex = 3

                for idx1 in xrange(len(segment_list)):
                    for idx2 in xrange(len1):
                        spec_min = spec_min_list[idx1]
                        spec_max = spec_max_list[idx1]
                        result_cell_idx = (result_ind2[0] + idx2, result_ind2[1] + result_idx)
                        if active_sheet.range(result_cell_idx).value != 'Something Wrong':
                            if isfloat(active_sheet.range((result_ind2[0] + idx2, result_ind2[1] + idx1)).value):
                                if spec_min == 'NA' and isfloat(spec_max):
                                    if float(active_sheet.range(
                                            (result_ind2[0] + idx2, result_ind2[1] + idx1)).value) <= float(
                                        spec_max):
                                        active_sheet.range(
                                            (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 4
                                    else:
                                        active_sheet.range(
                                            (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 3
                                        active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
                                        active_sheet.range(result_cell_idx).value = 'Fail'
                                elif isfloat(spec_min) and isfloat(spec_max):
                                    if float(spec_min) <= float(
                                            active_sheet.range(
                                                (result_ind2[0] + idx2, result_ind2[1] + idx1)).value) <= float(
                                        spec_max):
                                        active_sheet.range(
                                            (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 4
                                    else:
                                        active_sheet.range(
                                            (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 3
                                        active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
                                        active_sheet.range(result_cell_idx).value = 'Fail'
                                elif isfloat(spec_min) and spec_max == 'NA':
                                    if float(spec_min) <= float(
                                            active_sheet.range((result_ind2[0] + idx2, result_ind2[1] + idx1)).value):
                                        active_sheet.range(
                                            (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 4
                                    else:
                                        active_sheet.range(
                                            (result_ind2[0] + idx2, result_ind2[1] + idx1)).api.Interior.ColorIndex = 3
                                        active_sheet.range(result_cell_idx).api.Interior.ColorIndex = 3
                                        active_sheet.range(result_cell_idx).value = 'Fail'

                SetCellFont_current_region(active_sheet, cell_idx, 'Times New Roman', 12, 'l')
                SetCellBorder_current_region(active_sheet, cell_idx)
                active_sheet.autofit('c')


def RunAllTopologyCheck():
    # Check All Topology Remained at the same time
    wb = Book(xlsm_path).caller()
    sheet_list = [sh.name for sh in wb.sheets]
    sheet_nochange_list = ['Cover', 'History', 'Summary', 'User-Guide', 'Setting', 'NetList',
                           'SymbolList', 'ACT_diff', 'ACT_diff_count', 'ACT_se',
                           'ACT_se_count', 'power', 'PKG_Len_COMPONENT NAME', 'Color_Lib']
    sheet_template_list = list(set(sheet_list) - set(sheet_nochange_list))

    for sh in wb.sheets:
        # 剩下 Netlist, template, power, PKG_Len_COMPONENT NAME, Color_Lib
        # 代码可优化
        ###########################
        # if sh.name not in ['Cover', 'History', 'Summary', 'User-Guide', 'Setting', 'NetList',
        #                    'SymbolList', 'ACT_diff', 'ACT_diff_count', 'ACT_se', 'ACT_se_count']:
        ##########################
        # 优化后的代码只循环两遍
        if sh.name in sheet_template_list:
            # # print(sh.name)
            for cell in sh.api.UsedRange.Cells:
                # 只有 template 中有 Topology
                if cell.Value == 'Topology' and sh.xrange((cell.Row + 1, cell.Column)).value == 'Signal Type':
                    # 只会输出第一个Topology所在的单元格的坐标
                    # # print((cell.Row, cell.Column))
                    # try:
                    CheckTopology(specified_range=sh.xrange((cell.Row, cell.Column)))
                    # except:
                    #     pass


# 先忽略
def FormatConverter():
    # Convert the old topology table
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection

    if selection_range.value == 'Topology':
        # Extract the Topology data
        start_ind = (selection_range.row, selection_range.column)
        original_table_content = active_sheet.range(start_ind).current_region.value
        active_sheet.range(start_ind).current_region.clear()

        for idxx in xrange(len(original_table_content)):
            row_tmp = original_table_content[idxx]
            if 'Segment Name' in original_table_content[idxx]:
                segment_list = original_table_content[idxx]
                result_idx = segment_list.index('Result')
                segment_list = segment_list[1:result_idx + 1]
            elif 'Layer' in original_table_content[idxx]:
                layer_list = original_table_content[idxx]
                layer_list = layer_list[1:result_idx + 1]
            elif 'Trace Width' in row_tmp:
                trace_width_list = original_table_content[idxx]
                trace_width_list = trace_width_list[1:result_idx + 1]
                trace_width_list = ["'%s" % x for x in trace_width_list]
            elif 'Space' in original_table_content[idxx]:
                space_list = original_table_content[idxx]
                space_list = space_list[1:result_idx + 1]
            elif 'Min' in row_tmp:
                spec_min_list = original_table_content[idxx]
                spec_min_list = spec_min_list[1:result_idx + 1]
            elif 'Max' in row_tmp:
                spec_max_list = original_table_content[idxx]
                spec_max_list = spec_max_list[1:result_idx + 1]
            elif 'Net Name' in original_table_content[idxx]:
                bottom_title_list = original_table_content[idxx]
                bottom_title_list = bottom_title_list[1:result_idx + 1]

        active_sheet.range(start_ind).value = 'Topology'
        active_sheet.range((start_ind[0] + 1, start_ind[1])).value = 'Signal Type'

        active_sheet.range((start_ind[0] + 2, start_ind[1])).value = 'Start Component Name-Pin Number'
        active_sheet.range((start_ind[0] + 2, start_ind[1])).api.Interior.ColorIndex = 43
        active_sheet.range((start_ind[0] + 2, start_ind[1]), (start_ind[0] + 2 + 9, start_ind[1])).api.MergeCells = True
        active_sheet.range((start_ind[0] + 2, start_ind[1])).api.VerticalAlignment = Constants.xlCenter

        active_sheet.range((start_ind[0] + 2, start_ind[1] + 1)).value = 'End Component Name-Pin Number'
        active_sheet.range((start_ind[0] + 2, start_ind[1] + 1)).api.Interior.ColorIndex = 43
        active_sheet.range((start_ind[0] + 2, start_ind[1] + 1),
                           (start_ind[0] + 2 + 9, start_ind[1] + 1)).api.MergeCells = True
        active_sheet.range((start_ind[0] + 2, start_ind[1] + 1)).api.VerticalAlignment = Constants.xlCenter

        col_title_list1 = ['Start Segment Name', 'Layer', 'Layer Change', 'Connect Component', 'Cross Over',
                           'Trace Width', 'Space', 'Min', 'Max', 'Net Name']
        col_title_list2 = ['Segment Name', 'Layer', 'Layer Change', 'Connect Component', 'Cross Over', 'Trace Width',
                           'Space', 'Min', 'Max', 'Net Name']

        for idx in xrange(len(col_title_list1)):
            active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 2)).value = col_title_list1[idx]
            active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 2)).api.Interior.ColorIndex = 44

        current_col_idx = start_ind[1] + 2
        current_row_idx = start_ind[0] + 2
        for idx1 in xrange(len(segment_list)):
            seg = segment_list[idx1]
            current_col_idx += 1
            if seg == 'Segment Name':
                active_sheet.range((current_row_idx, current_col_idx)).value = 'Component Name'
                active_sheet.range((current_row_idx, current_col_idx)).api.Interior.ColorIndex = 43
                active_sheet.range((current_row_idx, current_col_idx),
                                   (current_row_idx + 9, current_col_idx)).api.MergeCells = True
                active_sheet.range((current_row_idx, current_col_idx)).api.VerticalAlignment = Constants.xlCenter
                current_col_idx += 1
            for idx2 in xrange(len(col_title_list2)):
                col_title = col_title_list2[idx2]
                if segment_list[idx1] == 'Segment Name':
                    active_sheet.range(
                        (current_row_idx + idx2, current_col_idx)).api.Interior.Color = RgbColor.rgbYellow
                else:
                    if col_title_list2[idx2] in ['Layer Change', 'Connect Component', 'Cross Over']:
                        active_sheet.range(
                            (current_row_idx + idx2, current_col_idx)).api.Interior.Color = RgbColor.rgbLightSteelBlue
                    else:
                        active_sheet.range(
                            (current_row_idx + idx2, current_col_idx)).api.Interior.Color = RgbColor.rgbSkyBlue

                if col_title_list2[idx2] == 'Segment Name':
                    active_sheet.range((current_row_idx + idx2, current_col_idx)).value = segment_list[idx1]
                elif col_title == 'Layer':
                    layer_change, connect_sch = False, False
                    layer_list_tmp = layer_list[idx1].split('/')
                    for layer_tmp in layer_list_tmp:
                        if layer_tmp in ['MSY', 'SLY', 'DSLY']:
                            layer_change = True
                        if layer_tmp in ['MSC', 'SLC', 'DSLC']:
                            connect_sch = True
                    for idx_tmp in xrange(len(layer_list_tmp)):
                        if layer_list_tmp[idx_tmp].find('C') > -1 or layer_list_tmp[idx_tmp].find('Y') > -1:
                            layer_list_tmp[idx_tmp] = layer_list_tmp[idx_tmp][0:-1]

                    active_sheet.range((current_row_idx + idx2, current_col_idx)).value = '/'.join(layer_list_tmp)
                elif col_title_list2[idx2] in ['Layer Change', 'Connect Component', 'Cross Over']:
                    if segment_list[idx1] != 'Segment Name':
                        active_sheet.range((current_row_idx + idx2, current_col_idx)).value = 'N'
                        if col_title == 'Layer Change' and layer_change:
                            active_sheet.range((current_row_idx + idx2, current_col_idx)).value = 'Y'
                        elif col_title_list2[idx2] == 'Connect Component' and connect_sch:
                            active_sheet.range((current_row_idx + idx2, current_col_idx)).value = 'Y'

                    elif segment_list[idx1] == 'Segment Name':
                        active_sheet.range((current_row_idx + idx2, current_col_idx)).value = col_title
                elif col_title == 'Trace Width':
                    active_sheet.range((current_row_idx + idx2, current_col_idx)).value = trace_width_list[idx1]
                elif col_title_list2[idx2] == 'Space':
                    active_sheet.range((current_row_idx + idx2, current_col_idx)).value = space_list[idx1]
                elif col_title_list2[idx2] == 'Min':
                    active_sheet.range((current_row_idx + idx2, current_col_idx)).value = spec_min_list[idx1]
                elif col_title_list2[idx2] == 'Max':
                    active_sheet.range((current_row_idx + idx2, current_col_idx)).value = spec_max_list[idx1]
                elif col_title == 'Net Name':
                    active_sheet.range((current_row_idx + idx2, current_col_idx)).value = bottom_title_list[idx1]

        SetCellFont_current_region(active_sheet, start_ind, 'Times New Roman', 12, 'l')
        SetCellBorder_current_region(active_sheet, start_ind)
        active_sheet.autofit('c')

        active_sheet.range((start_ind[0], start_ind[1] + 2), (start_ind[0] + 1, start_ind[1] + len(
            wb.sheets[active_sheet].range(start_ind).current_region.value[0]) - 1)).api.MergeCells = True


def LoadTopologyFormat(topology_type):
    # Create the supported topology table
    app = QtGui.QApplication(sys.argv)
    form = TopologyLayoutForm(topology_type)
    form.show()
    app.exec_()


# template生成表格
def LoadTopologyFormat_GUI():
    app = QtGui.QApplication(sys.argv)
    form = TopologyTypeSelection()
    form.show()
    app.exec_()


def DeleteTopologyTable():
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)
    if selection_range.value == 'Topology' or selection_range.value == 'Simple Topology':
        active_sheet.range(start_ind).current_region.clear()


def LoadSimpleTopology_GUI():
    app = QtGui.QApplication(sys.argv)
    form = SimpleTopology()
    form.show()
    app.exec_()


# 创建Simple Topology表格
def LoadSimpleTopologyFormat():
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)

    selection_range.value = 'Simple Topology'
    active_sheet.range(start_ind).api.Interior.Color = RgbColor.rgbSandyBrown
    active_sheet.range((start_ind[0] + 1, start_ind[1])).value = 'Start Component Name-Pin Number'
    active_sheet.range((start_ind[0] + 1, start_ind[1])).api.Interior.ColorIndex = 43
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 1)).value = 'End Component Name-Pin Number'
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 1)).api.Interior.ColorIndex = 43
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 2)).value = 'Start Net Name'
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 2)).api.Interior.ColorIndex = 43

    SetCellFont_current_region(active_sheet, start_ind, 'Times New Roman', 12, 'l')
    SetCellBorder_current_region(active_sheet, start_ind)
    active_sheet.autofit('c')


# 对Simple Topology表格进行填写
def LoadSimpleTopology():
    # start_time = time.clock()
    # Show the topology distribution
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)

    if selection_range.value == 'Simple Topology':

        active_sheet.range((start_ind[0] + 2, start_ind[1] + 3)).expand('table').clear()

        act_sheet_name_list = ['ACT_diff', 'ACT_se']
        try:
            if type(active_sheet.range((start_ind[0] + 2, start_ind[1])).expand('table').value[0]) == type(u''):
                start_sch_list = [active_sheet.range((start_ind[0] + 2, start_ind[1])).expand('table').value[0]]
                end_sch_list = [active_sheet.range((start_ind[0] + 2, start_ind[1])).expand('table').value[1]]
                start_net_list = [active_sheet.range((start_ind[0] + 2, start_ind[1])).expand('table').value[2]]
            else:
                # 注：*zip(m, n) 是 zip(m, n) 的逆过程
                start_sch_list, end_sch_list, start_net_list = zip(
                    *active_sheet.range((start_ind[0] + 2, start_ind[1])).expand('table').value)

        except TypeError:
            pass
            # start_sch_list, end_sch_list = [], []
            #
            # start_net_list = active_sheet.range((start_ind[0]+2, start_ind[1]+2)).options(expand = 'table', ndim = 1).value
            # start_net_list = [x for x in start_net_list if x != '-' and x.find('Total_Len:') == -1]

            # 这段代码是否有必要，对Simple Topology中只有Start Net Name的情况进行了讨论
            # 个人认为可以删除
            #########################################################
            # start_net_list_tmp = []
            # for net in start_net_list:
            #     for act_sheet_name in act_sheet_name_list:
            #         if wb.sheets[act_sheet_name].range('A1').value != None:
            #             act_content = wb.sheets[act_sheet_name].range('A1').current_region.value[1::]
            #             match_line = []
            #             for x in act_content:
            #                 if x[1] == net:
            #                     match_line.append(x[0:3])
            #             if match_line != []:
            #                 start_sch, start_net, end_sch = zip(*match_line)
            #                 start_net_list_tmp += start_net
            #                 start_sch_list += start_sch
            #                 end_sch_list += end_sch
            #                 break
            ########################################################

            # start_net_list = list(start_net_list_tmp)
        # start1_time = time.clock()
        # # print(1, start1_time - start_time)
        start_sch_list = [x for x in start_sch_list if x != '-']
        end_sch_list = [x for x in end_sch_list if x != '-']
        start_net_list = [x for x in start_net_list if x != '-' and x.find('Total_Len:') == -1]
        active_sheet.range((start_ind[0] + 2, start_ind[1])).expand('table').clear()
        # start2_time = time.clock()
        # # print(2, start2_time - start1_time)
        act_data_dict = dict()
        for act_sheet_name in act_sheet_name_list:
            act_content = wb.sheets[act_sheet_name].range('A1').current_region.value
            if act_content != None:
                for start_sch, end_sch, start_net in zip(start_sch_list, end_sch_list, start_net_list):
                    for line in act_content:
                        if [start_sch, start_net, end_sch] == line[0:3] and act_data_dict.get(
                                (start_sch, start_net, end_sch)) == None:
                            act_data_dict[(start_sch, start_net, end_sch)] = [x for x in line[5::] if
                                                                              x not in ['', None]]
                            break

        # start3_time = time.clock()
        # # print(3, start3_time - start2_time)
        # # print(act_data_dict)
        # # print(len(act_data_dict))
        topology_content_all = list()
        topology_content_len_all = list()
        idx = 0
        for start_sch, end_sch, start_net in zip(start_sch_list, end_sch_list, start_net_list):
            idx += 1
            if act_data_dict.get((start_sch, start_net, end_sch)) != None:
                topology_content = act_data_dict[(start_sch, start_net, end_sch)]

                topology_content_tmp = list()
                for i_ in xrange(len(topology_content)):
                    if str(topology_content[i_]).find(':') > -1:
                        # (u'[CPU-AF2]:TOP:3.5', 234.54)
                        topology_content_tmp.append((topology_content[i_], topology_content[i_ + 1]))
                        # break
                    # for i_ in xrange(len(topology_content)):
                    elif str(topology_content[i_]).find('net$') == 0:
                        topology_content_tmp.append((topology_content[i_], '-'))
                        # break
                topology_content, topology_content_len = zip(*topology_content_tmp)
                topology_content = list(topology_content)
                topology_content_len = list(topology_content_len)

                topology_content_all.append(list(topology_content))
                topology_content_len_all.append(list(topology_content_len))
        # start4_time = time.clock()
        # # print(4, start4_time - start3_time)
        # 设置表格无线条
        active_sheet.range(start_ind).current_region.api.Borders.LineStyle = LineStyle.xlLineStyleNone
        allegro_report_path, layer_type_dict, start_sch_name_list, progress_ind, All_Layer_List = GetSetting()
        # # print(layer_type_dict)
        All_Layer_List = list(set(layer_type_dict.keys()))

        max_len = max([len(x) for x in topology_content_all])
        for idx in xrange(len(topology_content_all)):
            for idx_ in xrange(max_len - len(topology_content_all[idx])):
                topology_content_all[idx].append('-')
                topology_content_len_all[idx].append('-')

        topology_content_all_all = zip(zip(start_sch_list, end_sch_list, start_net_list), topology_content_all)

        start_sch_net, topology_content_all = zip(*topology_content_all_all)
        # start5_time = time.clock()
        # # print(5, start5_time - start4_time)
        active_sheet.range((start_ind[0] + 2, start_ind[1])).expand('table').value = start_sch_net
        active_sheet.range((start_ind[0] + 2, start_ind[1] + 3)).expand('table').value = topology_content_all
        # # print(topology_content_all)
        # 可在此设置表格线框与颜色
        for idx in xrange(len(topology_content_all)):
            active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 2)).api.Borders(BordersIndex.xlEdgeTop) \
                .LineStyle = LineStyle.xlContinuous
            active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 2)).api.Borders(BordersIndex.xlEdgeTop) \
                .Weight = BorderWeight.xlThin
            active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 2)).api.Borders(BordersIndex.xlEdgeBottom) \
                .LineStyle = LineStyle.xlContinuous
            active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 2)).api.Borders(BordersIndex.xlEdgeBottom) \
                .Weight = BorderWeight.xlThin
            active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 2)).api.Borders(BordersIndex.xlEdgeLeft) \
                .LineStyle = LineStyle.xlContinuous
            active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 2)).api.Borders(BordersIndex.xlEdgeLeft) \
                .Weight = BorderWeight.xlThin
            active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 2)).api.Interior.ColorIndex = 35

        # 个人代码,加边框样式，加单元格颜色
        #######################################################
        for idx in xrange(len(topology_content_all)):
            # 设置初始颜色
            color_ind = 35
            net_ind = 0
            len_t = len(topology_content_all[idx])
            for idy in xrange(len_t):
                item_t = topology_content_all[idx][idy]
                if item_t == '-' or item_t == '':
                    break
                # # print((start_ind[0] + 2 + idx, start_ind[1]+3+idy))
                # 设置边框
                active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 3 + idy)).api.Borders(BordersIndex.xlEdgeTop) \
                    .LineStyle = LineStyle.xlContinuous
                active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 3 + idy)).api.Borders(BordersIndex.xlEdgeTop) \
                    .Weight = BorderWeight.xlThin
                active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 3 + idy)).api.Borders(
                    BordersIndex.xlEdgeBottom) \
                    .LineStyle = LineStyle.xlContinuous
                active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 3 + idy)).api.Borders(
                    BordersIndex.xlEdgeBottom) \
                    .Weight = BorderWeight.xlThin

                # 设置单元格颜色
                if str(item_t).find('net$') > -1:
                    # 变换线名就变换颜色
                    color_ind = 37 + net_ind
                    net_ind += 1
                active_sheet.range((start_ind[0] + 2 + idx, start_ind[1] + 3 + idy)).api.Interior.ColorIndex = color_ind
                if str(item_t).find(':') > -1:
                    # eg: '[FRONT_USB_HEADER-2]:BOTTOM:5.1'
                    layer = None
                    for x in item_t.split(':'):
                        if x in All_Layer_List:
                            layer = x
                    try:
                        x = topology_content_all[idx][idy + 1]
                        if str(x).find(':') > -1:
                            for x_ in x.split(':'):
                                if x_ in All_Layer_List:
                                    layer_next = x_
                                    # 换层计数
                                    if layer_next != layer:
                                        # 加粗左边框
                                        active_sheet.range(
                                            (start_ind[0] + 2 + idx, start_ind[1] + 4 + idy)).api.Borders(
                                            BordersIndex.xlEdgeLeft) \
                                            .LineStyle = LineStyle.xlContinuous
                                        active_sheet.range(
                                            (start_ind[0] + 2 + idx, start_ind[1] + 4 + idy)).api.Borders(
                                            BordersIndex.xlEdgeLeft) \
                                            .Weight = BorderWeight.xlThick
                    except IndexError:
                        pass
            ###############################################

        SetCellFont_current_region(active_sheet, start_ind, 'Times New Roman', 12, 'l')
        active_sheet.autofit('c')
        # start6_time = time.clock()
        # # print('end', start6_time - start_time)


def load_calc_grade_gui():
    '''调用tab表格gui'''
    app = QtGui.QApplication(sys.argv)
    form = LoadCalcGradeForm()
    form.show()
    app.exec_()


class LoadCalcGradeForm(QtGui.QDialog):
    def __init__(self, parent=None):
        super(LoadCalcGradeForm, self).__init__(parent)
        self.setWindowTitle('Calcute Grade')
        self.b111 = QtGui.QPushButton('show_item', self)
        self.b111.clicked.connect(self.show_item)
        self.b222 = QtGui.QPushButton('Calc Grade', self)
        self.b222.clicked.connect(self.calc_grade)
        self.b333 = QtGui.QPushButton('Cancel', self)
        self.b333.clicked.connect(self.close_dialog)

        layout = QtGui.QGridLayout()
        layout.addWidget(self.b111, 0, 0)
        layout.addWidget(self.b222, 1, 0)
        layout.addWidget(self.b333, 2, 0)
        self.setLayout(layout)

    def close_dialog(self):
        QtGui.QDialog.accept(self)

    def show_item(self):
        QtGui.QDialog.accept(self)
        CalcGrade()

    def calc_grade(self):
        QtGui.QDialog.accept(self)
        CalcGrade(True)


def CalcGrade(user_defined=False):
    root_path = os.getcwd()
    # print(root_path)
    root_path = '\\'.join(root_path[:-5].split('\\'))
    parents = os.listdir(root_path)

    for parent in parents:
        if parent.find('.xlsm') > -1 and parent.find('$') == -1 and parent != 'LayoutScore.xlsm':
            wb2_file = parent
        if parent.find('.xlsx') > -1 and parent.find('$') == -1:
            wb3_file = parent
    # # print(wb2_file, xlsm_path)
    wb1 = Book(xlsm_path).caller()
    wb2 = Book(wb2_file)
    wb3 = Book(wb3_file)

    # 对checklist进行评分
    if user_defined == False:
        # # print(123)
        sheet_name_list = []
        sheet_table = ['Cover', 'History', 'Summary', 'User-Guide', 'Setting', 'NetList', 'SymbolList', 'ACT_diff',
                       'ACT_se', 'ACT_diff_count', 'ACT_se_count', 'LayoutScore', 'Netlist', 'Color_Lib',
                       'PKG_Len_COMPONENT NAME', 'power']
        for i in xrange(wb2.sheets.count):
            # # print(wb.sheets[i].name)
            if wb2.sheets[i].name.title().find('Template') > -1:
                sheet_table.append(wb2.sheets[i].name)
            sheet_name_list.append(wb2.sheets[i].name)
        # # print(sheet_name_list)
        sheet_name_list = [x for x in sheet_name_list if x not in sheet_table]
        # # print(sheet_name_list)
        # 三种item，三种评分规则
        item1_list = ['DMI', 'CLK', 'CLOCK', 'USB2.0', 'USB3.0', 'USB3.1', 'PCIE', 'CPU-PCIE', 'PCH-PCIE', 'CPU PCIE',
                      'PCH PCIE', 'CPU_PCIE', 'PCH_PCIE', 'HDMI', 'SATA', 'M2', 'M.2', 'TYPEC', 'USB-TYPEC',
                      'USB TYPEC',
                      'USB_TYPEC', 'USB TYPE_C', 'USB TYPE-C', 'DP', 'LAN', 'CNVI']
        item2_list = ['SPI', 'LPC', 'SMBUS', 'HDA', 'ESPI']

        check_list1 = []
        check_list2 = []
        check_list3 = []

        # 分类item
        for x in sheet_name_list:
            y = x.upper()
            if y in item1_list:
                check_list1.append(x)
            elif y in item2_list:
                check_list2.append(x)
            else:
                check_list3.append(x)

        item1_name = 'High Speed Differential checklist('
        item2_name = 'High Speed single-Ended1 checklist('
        item3_name = 'Lower Speed single-Ended2 checklist('

        item1_name = output_item_name(check_list1, item1_name)
        item2_name = output_item_name(check_list2, item2_name)
        item3_name = output_item_name(check_list3, item3_name)

        if check_list1:
            wb1.sheets['LayoutScore'].range((9, 3)).value = item1_name
        if check_list2:
            wb1.sheets['LayoutScore'].range((10, 3)).value = item2_name
        if check_list3:
            wb1.sheets['LayoutScore'].range((11, 3)).value = item3_name

        wb1.sheets['LayoutScore'].range((12, 3)).value = 'DFE(Rule1-1, Rule1-2, Rule1-3, Rule5-1)'
        wb1.sheets['LayoutScore'].range((13, 3)).value = 'DFE(Rule3-1, Rule4-1, Rule4-2)'

        wb1.sheets['LayoutScore'].autofit('c')
    else:
        check_list1 = wb1.sheets['LayoutScore'].range((9, 3)).value.split('(')[-1][:-1].split('/')
        check_list2 = wb1.sheets['LayoutScore'].range((10, 3)).value.split('(')[-1][:-1].split('/')
        check_list3 = wb1.sheets['LayoutScore'].range((11, 3)).value.split('(')[-1][:-1].split('/')
        # # print(check_list3)
        # 对每个item进行评分包括：1.fail个数 2.total length超过10%或20% 3.total或者分段的长度不匹配
        # 4.过孔数目超过多于2个 5.将每个sheet fail的个数及总个数show在sheet开头
        # # print('pre', start_time2 - start_time1)
        # 对item1进行评分
        item1_count_all = 0.0
        item1_count_fail = 0
        via_grade = 0
        total_length_grade1 = 0
        total_length_grade_list1 = []
        mismatch_grade1 = 0
        if check_list1 != ['']:
            for sh_name in check_list1:
                # # print(sh_name)
                sh1_count_all = 0
                sh1_count_fail = 0
                # # print(sh_name)
                # start1_time = time.clock()
                for cell in wb2.sheets[sh_name].api.UsedRange.Cells:
                    if cell.Value == 'Topology':
                        content = wb2.sheets[sh_name].range((cell.Row, cell.Column)).current_region.value
                        # # print(content)
                        content2 = content[2]
                        via_out_list = []
                        total_lenght_out_list = []
                        try:
                            via_idx = content2.index('Via Count')
                        except:
                            via_idx = 0
                        try:
                            total_length_idx = content2.index('Total Length')
                        except:
                            total_length_idx = 0

                        # 对mismatch进行管控
                        try:
                            layer_mismatch_idx = content2.index('Layer Mismatch')
                        except:
                            layer_mismatch_idx = False
                        try:
                            segment_mismatch_idx = content2.index('Segment Mismatch')
                        except:
                            segment_mismatch_idx = False
                        try:
                            group_mismatch_idx = content2.index('Group Mismatch')
                        except:
                            group_mismatch_idx = False
                        try:
                            total_mismatch_idx = content2.index('Total Mismatch')
                        except:
                            total_mismatch_idx = False

                        for con_idx in xrange(len(content)):
                            # 得出总的net数目及fail的net数
                            if content[con_idx][-1] == 'Pass':
                                item1_count_all += 1
                                sh1_count_all += 1
                            elif content[con_idx][-1] == 'Fail' or content[con_idx][-1] == 'Minor':
                                item1_count_all += 1
                                item1_count_fail += 1
                                sh1_count_all += 1
                                sh1_count_fail += 1

                                # 判断via
                                if via_idx:
                                    via_standard = content[10][via_idx]

                                    if wb2.sheets[sh_name].range((cell.Row + con_idx, cell.Column + via_idx)). \
                                            api.Interior.ColorIndex in [3, 44]:
                                        via_out_list.append(content[con_idx][via_idx])

                                # 判断 total_length
                                if total_length_idx:
                                    total_length_standard = content[10][total_length_idx]

                                    if wb2.sheets[sh_name].range((cell.Row + con_idx, cell.Column + total_length_idx)). \
                                            api.Interior.ColorIndex in [3, 44]:
                                        total_lenght_out_list.append(content[con_idx][total_length_idx])

                                if mismatch_grade1 == 0 and layer_mismatch_idx and wb2.sheets[sh_name]. \
                                        xrange((cell.Row + con_idx, cell.Column + layer_mismatch_idx)) \
                                        .api.Interior.ColorIndex in [3, 44]:
                                    mismatch_grade1 = 0.3
                                    break
                                elif mismatch_grade1 == 0 and total_mismatch_idx and wb2.sheets[sh_name]. \
                                        xrange((cell.Row + con_idx, cell.Column + total_mismatch_idx)) \
                                        .api.Interior.ColorIndex in [3, 44]:
                                    mismatch_grade1 = 0.3
                                    break
                                elif mismatch_grade1 == 0 and segment_mismatch_idx and wb2.sheets[sh_name] \
                                        .xrange((cell.Row + con_idx, cell.Column + segment_mismatch_idx)). \
                                        api.Interior.ColorIndex in [3, 44]:
                                    mismatch_grade1 = 0.3
                                    break
                                elif mismatch_grade1 == 0 and group_mismatch_idx and wb2.sheets[sh_name]. \
                                        xrange((cell.Row + con_idx, cell.Column + group_mismatch_idx)). \
                                        api.Interior.ColorIndex in [3, 44]:
                                    mismatch_grade1 = 0.3
                                    break

                        # 如果via超过管控
                        if via_out_list:
                            for via_item in via_out_list:
                                if (via_item - via_standard) > 2:
                                    via_grade = 0.2

                        # 如果total length 超过管控（10% and 20%）
                        if total_lenght_out_list:
                            for length_item in total_lenght_out_list:
                                if length_item > total_length_standard and \
                                        (length_item - total_length_standard) / total_length_standard > 0.2:
                                    total_length_grade_list1.append(0.5)
                                    break
                                elif length_item > total_length_standard and \
                                        (length_item - total_length_standard) / total_length_standard > 0.1:
                                    total_length_grade_list1.append(0.3)
                # # print(sh1_count_fail, sh1_count_all)
                wb2.sheets[sh_name].range((3, 1)).value = 'all : ' + str(sh1_count_all) + ' fail : ' + \
                                                          str(sh1_count_fail)
                wb2.sheets[sh_name].autofit('c')
        #         start2_time = time.clock()
        #         # print(start2_time - start1_time)
        # # print('item1', item1_count_fail, item1_count_all)
        # 得出分数
        if check_list1 != ['']:
            if item1_count_fail:
                if total_length_grade_list1:
                    total_length_grade1 = max(total_length_grade_list1)
                item1_grade = 5.5 - round(item1_count_fail / item1_count_all * 5.5, 2) - round(via_grade, 2) - \
                              round(total_length_grade1, 2) - round(mismatch_grade1, 2)
            else:
                item1_grade = 5.5
        else:
            item1_grade = 0
        # start3_time = time.clock()
        # 对item2进行评分
        item2_count_all = 0.0
        item2_count_fail = 0
        total_length_grade2 = 0
        total_length_grade_list2 = []
        mismatch_grade2 = 0
        if check_list2 != ['']:
            for sh_name in check_list2:
                # # print(sh_name)
                sh2_count_all = 0
                sh2_count_fail = 0
                # # print(sh_name)
                # start4_time = time.clock()
                for cell in wb2.sheets[sh_name].api.UsedRange.Cells:
                    # # print(wb2.sheets[sh_name].range((6, 2)).value)
                    if cell.Value == 'Topology':
                        # # print(cell.Row, cell.Column)
                        content = wb2.sheets[sh_name].range((cell.Row, cell.Column)).current_region.value

                        total_lenght_out_list = []

                        try:
                            total_length_idx = content[2].index('Total Length')
                        except:
                            total_length_idx = 0

                        # 对mismatch进行管控
                        try:
                            group_mismatch_idx = content[2].index('Group Mismatch')
                        except:
                            group_mismatch_idx = False

                        for con_idx in xrange(len(content)):
                            # 得出总的net数目及fail的net数
                            if content[con_idx][-1] == 'Pass':
                                item2_count_all += 1
                                sh2_count_all += 1
                            elif content[con_idx][-1] == 'Fail' or content[con_idx][-1] == 'Minor':
                                item2_count_all += 1
                                sh2_count_all += 1
                                item2_count_fail += 1
                                sh2_count_fail += 1

                            if total_length_idx:
                                total_length_standard = content[10][total_length_idx]

                                if wb2.sheets[sh_name].range((cell.Row + con_idx, cell.Column + total_length_idx)). \
                                        api.Interior.ColorIndex in [3, 44]:
                                    total_lenght_out_list.append(content[con_idx][total_length_idx])

                            if mismatch_grade2 == 0 and group_mismatch_idx and wb2.sheets[sh_name]. \
                                    xrange((cell.Row + con_idx, cell.Column + group_mismatch_idx)) \
                                    .api.Interior.ColorIndex in [3, 44]:
                                mismatch_grade2 = 0.08

                        # 如果total length 超过管控（20%）
                        if total_lenght_out_list:
                            for length_item in total_lenght_out_list:
                                if length_item > total_length_standard and \
                                        (length_item - total_length_standard) / total_length_standard > 0.2:
                                    total_length_grade_list2.append(0.1)
                                    break
                                elif length_item > total_length_standard and \
                                        (length_item - total_length_standard) / total_length_standard > 0.1:
                                    total_length_grade_list2.append(0.08)

                wb2.sheets[sh_name].range((3, 1)).value = 'all : ' + str(sh2_count_all) + ' fail : ' + \
                                                          str(sh2_count_fail)
                wb2.sheets[sh_name].autofit('c')
        #         start5_time = time.clock()
        #         # print(start5_time - start4_time)
        # # print('item2', item2_count_fail, item2_count_all)
        # 得出分数
        if check_list2 != ['']:
            if item2_count_fail:
                if total_length_grade_list2:
                    total_length_grade2 = max(total_length_grade_list2)
                item2_grade = 1.2 - round(item2_count_fail / item2_count_all * 1.2, 2) - round(total_length_grade2, 2) - \
                              round(mismatch_grade2, 2)
            else:
                item2_grade = 1.2
        else:
            item2_grade = 0
        # start6_time = time.clock()
        # # print('item2', start6_time - start5_time)
        # 对item3进行评分
        item3_count_all = 0.0
        item3_count_fail = 0
        total_length_grade3 = 0
        total_length_grade_list3 = []
        mismatch_grade3 = 0
        # # print(check_list3)
        if check_list3 != ['']:
            # # print(123)
            for sh_name in check_list3:
                # # print(sh_name)
                sh3_count_all = 0
                sh3_count_fail = 0
                # start7_time = time.clock()
                # # print(sh_name)
                for cell in wb2.sheets[sh_name].api.UsedRange.Cells:
                    if cell.Value == 'Topology':
                        content = wb2.sheets[sh_name].range((cell.Row, cell.Column)).current_region.value
                        # # print(content)

                        total_lenght_out_list = []

                        try:
                            total_length_idx = content[2].index('Total Length')
                        except:
                            total_length_idx = 0
                        # 对mismatch进行管控
                        try:
                            group_mismatch_idx = content[2].index('Group Mismatch')
                        except:
                            group_mismatch_idx = False

                        for con_idx in xrange(len(content)):
                            # 得出总的net数目及fail的net数
                            if content[con_idx][-1] == 'Pass':
                                item3_count_all += 1
                                sh3_count_all += 1
                            elif content[con_idx][-1] == 'Fail' or content[con_idx][-1] == 'Minor':
                                item3_count_all += 1
                                item3_count_fail += 1
                                sh3_count_all += 1
                                sh3_count_fail += 1

                            if total_length_idx:
                                total_length_standard = content[10][total_length_idx]

                                if wb2.sheets[sh_name].range((cell.Row + con_idx, cell.Column + total_length_idx)). \
                                        api.Interior.ColorIndex in [3, 44]:
                                    total_lenght_out_list.append(content[con_idx][total_length_idx])

                            if mismatch_grade3 == 0 and group_mismatch_idx and wb2.sheets[sh_name]. \
                                    xrange((cell.Row + con_idx, cell.Column + group_mismatch_idx)) \
                                    .api.Interior.ColorIndex in [3, 44]:
                                mismatch_grade3 = 0.03

                        # 如果total length 超过管控（20%）
                        if total_lenght_out_list:
                            for length_item in total_lenght_out_list:
                                if length_item > total_length_standard and \
                                        (length_item - total_length_standard) / total_length_standard > 0.2:
                                    total_length_grade_list3.append(0.05)
                                    break
                                elif length_item > total_length_standard and \
                                        (length_item - total_length_standard) / total_length_standard > 0.1:
                                    total_length_grade_list3.append(0.03)
                wb2.sheets[sh_name].range((3, 1)).value = 'all : ' + str(sh3_count_all) + ' fail : ' + \
                                                          str(sh3_count_fail)
                wb2.sheets[sh_name].autofit('c')
        #         start8_time = time.clock()
        #         # print(start8_time - start7_time)
        # # print('item3', item3_count_fail, item3_count_all)
        # 得出分数
        if check_list3:
            if item3_count_fail:
                if total_length_grade_list3:
                    total_length_grade3 = max(total_length_grade_list3)
                item3_grade = 0.3 - round(item3_count_fail / item3_count_all * 0.3, 2) - round(total_length_grade3, 2) - \
                              round(mismatch_grade3, 2)
            else:
                item3_grade = 0.3
        else:
            item3_grade = 0
        #
        # # print(item1_grade, item2_grade, item3_grade)
        #
        # start9_time = time.clock()
        # # print('item3', start9_time - start8_time)

        # 对DFE进行评分
        netlist_sheet = wb2.sheets['NetList']
        start_time = time.clock()
        for cell in netlist_sheet.api.UsedRange.Cells:
            # # print(2222)
            if cell.Value == 'Differential':
                # # print(333)
                diff_list1 = []
                diff_list2 = []

                # 个人认为，多余操作，后来证明不多余
                #####################################
                for dp in netlist_sheet.range((cell.Row + 1, cell.Column)).options(expand='table', ndim=2).value:
                    diff_list1.append(dp[0])
                    diff_list2.append(dp[1])

                diff_list = diff_list1 + diff_list2

                diff_dict = dict()
                for idx_dp in xrange(len(diff_list1)):
                    diff_dict[diff_list1[idx_dp]] = diff_list2[idx_dp]
                    diff_dict[diff_list2[idx_dp]] = diff_list1[idx_dp]
                # ######################################
                # # print(len(diff_dict.keys()))
                break
        # start_time1 = time.clock()
        # # print(start_time1 - start_time)
        try:
            rule1_1_sheet = wb3.sheets['Rule1-1']
        except:
            rule1_1_sheet = wb3.sheets['Rule 1-1']
        try:
            rule1_2_sheet = wb3.sheets['Rule1-2']
        except:
            rule1_2_sheet = wb3.sheets['Rule 1-2']
        try:
            rule1_3_sheet = wb3.sheets['Rule1-3']
        except:
            rule1_3_sheet = wb3.sheets['Rule 1-3']
        try:
            rule3_1_sheet = wb3.sheets['Rule3-1']
        except:
            rule3_1_sheet = wb3.sheets['Rule3-1']
        try:
            rule4_1_sheet = wb3.sheets['Rule 4-1']
        except:
            rule4_1_sheet = wb3.sheets['Rule4-1']
        try:
            rule4_2_sheet = wb3.sheets['Rule 4-2']
        except:
            rule4_2_sheet = wb3.sheets['Rule4-2']
        try:
            rule5_1_sheet = wb3.sheets['Rule5-1']
        except:
            rule5_1_sheet = wb3.sheets['Rule 5-1']

        def get_diff_net(sheet):
            diff_all_list = []
            diff_real_list = []
            for cell in sheet.api.UsedRange.Cells:
                if cell.Value == 'Index':
                    index = (cell.Row + 1, cell.Column)
                    content = sheet.range(index).expand('table').value
            try:
                for x_ind in xrange(len(content)):
                    if sheet.range((x_ind + 8, 9)).api.Interior.ColorIndex != 14:
                        if content[x_ind][1] in diff_list:
                            diff_all_list.append(content[x_ind][1])

                for x in diff_all_list:
                    if diff_dict[x] in diff_all_list:
                        diff_real_list.append(x)
            except:
                diff_real_list = []

            return list(set(diff_real_list))

        diff_real_list1_1 = get_diff_net(rule1_1_sheet)
        rule1_1_sheet.range((1, 1)).value = 'Diff Pair : ' + str(len(diff_real_list1_1))
        # start_time2 = time.clock()
        # # print(start_time2 - start_time1)
        # # print(diff_real_list1_1)
        diff_real_list1_2 = get_diff_net(rule1_2_sheet)
        rule1_2_sheet.range((1, 1)).value = 'Diff Pair : ' + str(len(diff_real_list1_2))
        # start_time3 = time.clock()
        # # print(start_time3 - start_time2)
        # # print(diff_real_list1_2)
        diff_real_list1_3 = get_diff_net(rule1_3_sheet)
        rule1_3_sheet.range((1, 1)).value = 'Diff Pair : ' + str(len(diff_real_list1_3))
        # # print(diff_real_list1_3)
        diff_real_list3_1 = get_diff_net(rule3_1_sheet)
        rule3_1_sheet.range((1, 1)).value = 'Diff Pair : ' + str(len(diff_real_list3_1))
        # # print(diff_real_list3_1)
        diff_real_list4_1 = get_diff_net(rule4_1_sheet)
        rule4_1_sheet.range((1, 1)).value = 'Diff Pair : ' + str(len(diff_real_list4_1))
        # # print(diff_real_list4_1)
        diff_real_list4_2 = get_diff_net(rule4_2_sheet)
        rule4_2_sheet.range((1, 1)).value = 'Diff Pair : ' + str(len(diff_real_list4_2))
        # # print(diff_real_list4_2)
        diff_real_list5_1 = get_diff_net(rule5_1_sheet)
        rule5_1_sheet.range((1, 1)).value = 'Diff Pair : ' + str(len(diff_real_list5_1))
        # # print(diff_real_list5_1)
        diff_count1 = (len(diff_real_list1_1) + len(diff_real_list1_2) + len(diff_real_list1_3)
                       + len(diff_real_list5_1)) / 2
        diff_count1_score = diff_count1 * 0.05
        diff_count2 = (len(diff_real_list3_1) + len(diff_real_list4_1) + len(diff_real_list4_2)) / 2
        diff_count2_score = diff_count2 * 0.05

        # start10_time = time.clock()
        #
        # # print('DFE', start10_time - start9_time)

        # 对分数进行处理，不为负
        def trans_score(score):
            if score < 0:
                score = 0
            else:
                score = round(score, 2)
            return score

        # 生成分数表格
        if check_list1 != ['']:
            wb1.sheets['LayoutScore'].range((9, 4)).value = 5.50
            wb1.sheets['LayoutScore'].range((9, 5)).value = round((item1_count_fail / item1_count_all) * 5.5, 2)
            wb1.sheets['LayoutScore'].range((9, 6)).value = round(mismatch_grade1, 2)
            wb1.sheets['LayoutScore'].range((9, 7)).value = round(via_grade, 2)
            wb1.sheets['LayoutScore'].range((9, 8)).value = round(total_length_grade1, 2)
            wb1.sheets['LayoutScore'].range((9, 9)).value = trans_score(item1_grade)
        else:
            wb1.sheets['LayoutScore'].range((9, 4)).value = 0
            wb1.sheets['LayoutScore'].range((9, 5)).value = 0
            wb1.sheets['LayoutScore'].range((9, 6)).value = 0
            wb1.sheets['LayoutScore'].range((9, 7)).value = 0
            wb1.sheets['LayoutScore'].range((9, 8)).value = 0
            wb1.sheets['LayoutScore'].range((9, 8)).value = 0

        if check_list2 != ['']:
            wb1.sheets['LayoutScore'].range((10, 4)).value = 1.20
            wb1.sheets['LayoutScore'].range((10, 5)).value = round((item2_count_fail / item2_count_all) * 1.2, 2)
            wb1.sheets['LayoutScore'].range((10, 6)).value = round(mismatch_grade2, 2)
            wb1.sheets['LayoutScore'].range((10, 7)).value = 0
            wb1.sheets['LayoutScore'].range((10, 8)).value = round(total_length_grade2, 2)
            wb1.sheets['LayoutScore'].range((10, 9)).value = trans_score(item2_grade)
        else:
            wb1.sheets['LayoutScore'].range((10, 4)).value = 0
            wb1.sheets['LayoutScore'].range((10, 5)).value = 0
            wb1.sheets['LayoutScore'].range((10, 6)).value = 0
            wb1.sheets['LayoutScore'].range((10, 7)).value = 0
            wb1.sheets['LayoutScore'].range((10, 8)).value = 0
            wb1.sheets['LayoutScore'].range((10, 9)).value = 0

        if check_list3 != ['']:
            wb1.sheets['LayoutScore'].range((11, 4)).value = 0.30
            wb1.sheets['LayoutScore'].range((11, 5)).value = round((item3_count_fail / item3_count_all) * 0.3, 2)
            wb1.sheets['LayoutScore'].range((11, 6)).value = round(mismatch_grade3, 2)
            wb1.sheets['LayoutScore'].range((11, 7)).value = 0
            wb1.sheets['LayoutScore'].range((11, 8)).value = round(total_length_grade3, 2)
            wb1.sheets['LayoutScore'].range((11, 9)).value = trans_score(item3_grade)
        else:
            wb1.sheets['LayoutScore'].range((11, 4)).value = 0
            wb1.sheets['LayoutScore'].range((11, 5)).value = 0
            wb1.sheets['LayoutScore'].range((11, 6)).value = 0
            wb1.sheets['LayoutScore'].range((11, 7)).value = 0
            wb1.sheets['LayoutScore'].range((11, 8)).value = 0
            wb1.sheets['LayoutScore'].range((11, 9)).value = 0

        wb1.sheets['LayoutScore'].range((12, 4)).value = 2
        wb1.sheets['LayoutScore'].range((12, 5)).value = diff_count1_score
        wb1.sheets['LayoutScore'].range((12, 6)).value = '/'
        wb1.sheets['LayoutScore'].range((12, 7)).value = '/'
        wb1.sheets['LayoutScore'].range((12, 8)).value = '/'
        wb1.sheets['LayoutScore'].range((12, 9)).value = trans_score(2 - diff_count1_score)

        wb1.sheets['LayoutScore'].range((13, 4)).value = 1
        wb1.sheets['LayoutScore'].range((13, 5)).value = diff_count2_score
        wb1.sheets['LayoutScore'].range((13, 6)).value = '/'
        wb1.sheets['LayoutScore'].range((13, 7)).value = '/'
        wb1.sheets['LayoutScore'].range((13, 8)).value = '/'
        wb1.sheets['LayoutScore'].range((13, 9)).value = trans_score(1 - diff_count2_score)
        wb1.sheets['LayoutScore'].range((14, 9)).value = round(trans_score(item1_grade) + trans_score(item2_grade)
                                                               + trans_score(item3_grade) +
                                                               trans_score(2 - diff_count1_score) +
                                                               trans_score(1 - diff_count2_score), 2)


def output_item_name(check_list, item_name):
    for idx1 in xrange(len(check_list)):
        # # print(check_list[idx1])
        if idx1 == (len(check_list) - 1):
            item_name = item_name + check_list[idx1] + ')'
        else:
            item_name = item_name + check_list[idx1] + '/'
    return item_name


def GenerateSummary():
    start_time = time.clock()
    wb = Book(xlsm_path).caller()

    # Get BoardFile Name
    try:
        allegro_report_path, layer_type_dict, start_sch_name_list, progress_ind, All_Layer_List = GetSetting()
        try:
            report_content = open(allegro_report_path, 'r').read().splitlines()
            brd_name = None
            for i in report_content:
                if i.find('.brd') > -1:
                    brd_name = i.split('/')[-1]
                    break
        except:
            brd_name = allegro_report_path.split('/')[-1].split('_allegro_report.rpt')[0]
    except:
        brd_name = ''
    # 让我来看看是什么原因
    # # #print(brd_name)
    # # #print(wb.sheets.count)
    # Get Sheet Name
    sheet_name_list = []
    sheet_table = ['Cover', 'History', 'Summary', 'User-Guide', 'Setting', 'NetList', 'SymbolList', 'ACT_diff',
                   'ACT_se',
                   'ACT_diff_count', 'ACT_se_count']
    for i in xrange(wb.sheets.count):
        # # #print(wb.sheets[i].name)
        if wb.sheets[i].name.title().find('Template') > -1:
            sheet_table.append(wb.sheets[i].name)
        sheet_name_list.append(wb.sheets[i].name)
    sheet_name_list = [x for x in sheet_name_list if x not in sheet_table]

    #  # #print(sheet_name_list)
    # Get idx in Summary Sheet
    # summary_sheet = wb.sheets.active
    summary_sheet = wb.sheets['Summary']
    brd_idx, result_idx = None, None
    for cell in summary_sheet.api.UsedRange.Cells:
        if cell.Value == '(1) Layout boardfile:':
            brd_idx = (cell.Row, cell.Column + 2)
        if cell.Value == '3.Check Results':
            result_idx = (cell.Row + 3, cell.Column)
            summary_sheet.range(result_idx).expand('table').clear()

    # Write boardfile name
    if brd_name != '':
        summary_sheet.range(brd_idx).value = brd_name

    # Dict of Suggestion
    suggest_dict = {'Total Length': 'Shrink the total trace length',
                    'Layer Mismatch': 'Adjust the layer mismatch', 'Total Mismatch': 'Adjust the total mismatch',
                    'Segment Mismatch': 'Adjust the segment mismatch',
                    'Skew to Bundle': 'Adjust the trace length', 'Skew to Target': 'Adjust the trace length',
                    'Via Count': 'Reduce the via count'}

    # Get Summary Result
    fail_number = 0
    total_check_number = 0
    check_result_dict = dict()
    for sh_name in sheet_name_list:
        PASS = 'Pass'
        Topology_Sheet = False
        VaildData = False
        for cell in wb.sheets[sh_name].api.UsedRange.Cells:
            if not PASS:
                break
            if cell.Value == 'Topology':
                Topology_Sheet = True
                content = wb.sheets[sh_name].range((cell.Row, cell.Column)).current_region.value
                # # #print(content)
                ##################
                # useless code
                for idx_ in xrange(len(content)):
                    if content[idx_][0] == 'Topology':
                        # # #print(idx_)
                        content = content[idx_::]
                        break
                ##################
                # # #print(content)
                segment_list = content[2]
                for idx1 in xrange(len(content)):
                    x = content[idx1]
                    if 'Pass' in content[idx1]:
                        VaildData = True
                        total_check_number += 1
                    if 'Fail' in x:  # x or 'Minor' in x:
                        # if 'Fail' in x:
                        PASS = 'Fail'
                        fail_number += 1
                        total_check_number += 1
                        # what is minor
                        # elif 'Minor' in content[idx1]:
                        #     PASS = 'Minor'
                        #     total_check_number += 1
                        for idx2 in xrange(len(content[idx1])):
                            if wb.sheets[sh_name].range(
                                    (cell.Row + idx1, cell.Column + idx2)).api.Interior.ColorIndex == 3 \
                                    and wb.sheets[sh_name].range((cell.Row + idx1, cell.Column + idx2)).value != 'Fail':
                                if check_result_dict.get((sh_name, 'fail_item')) is None:

                                    item = segment_list[idx2].split(':')[0]
                                    check_result_dict[(sh_name, 'fail_item')] = [item]

                                    if suggest_dict.get(item) is None:
                                        check_result_dict[(sh_name, 'suggestion')] = ['Adjust the segment trace length']
                                    else:
                                        check_result_dict[(sh_name, 'suggestion')] = [suggest_dict[item]]
                                else:
                                    item = segment_list[idx2].split(':')[0]
                                    check_result_dict[(sh_name, 'fail_item')] += [item]

                                    if suggest_dict.get(item) is None:
                                        check_result_dict[(sh_name, 'suggestion')] += [
                                            'Adjust the segment trace length']
                                    else:
                                        check_result_dict[(sh_name, 'suggestion')] += [suggest_dict[item]]

        if Topology_Sheet:
            check_result_dict[(sh_name, 'pass_or_fail')] = PASS
            if PASS == 'Pass' and not VaildData:
                check_result_dict[(sh_name, 'pass_or_fail')] = 'NA'

    # # #print(check_result_dict)

    # Write Result
    idx = -1
    for item in sheet_name_list:
        if check_result_dict.get((item, 'pass_or_fail')) != None:
            idx += 1
            summary_sheet.range((result_idx[0] + idx, result_idx[1] + 1)).value = check_result_dict[
                (item, 'pass_or_fail')]
            if check_result_dict[(item, 'pass_or_fail')] == 'Pass':
                summary_sheet.range((result_idx[0] + idx, result_idx[1] + 1)).api.Interior.ColorIndex = 4
                summary_sheet.range((result_idx[0] + idx, result_idx[1] + 2)).value = '-'
                summary_sheet.range((result_idx[0] + idx, result_idx[1] + 3)).value = '-'
                summary_sheet.range((result_idx[0] + idx, result_idx[1] + 4)).value = '-'
            elif check_result_dict[(item, 'pass_or_fail')] == 'NA':
                summary_sheet.range((result_idx[0] + idx, result_idx[1] + 2)).value = '-'
                summary_sheet.range((result_idx[0] + idx, result_idx[1] + 3)).value = '-'
                summary_sheet.range((result_idx[0] + idx, result_idx[1] + 4)).value = '-'
            elif check_result_dict[(item, 'pass_or_fail')] == 'Fail':
                summary_sheet.range((result_idx[0] + idx, result_idx[1] + 1)).api.Interior.ColorIndex = 3

                fail_item_list = check_result_dict[(item, 'fail_item')]
                fail_item_list = set(fail_item_list)
                fail_item_list = ['%s is out of SPEC' % x for x in fail_item_list]

                fail_item = ''
                for i in xrange(len(fail_item_list)):
                    fail_item += '(%d). %s\n' % (i + 1, fail_item_list[i])

                suggest_list = check_result_dict[(item, 'suggestion')]
                suggest_list = set(suggest_list)
                suggest_item = ''

                count = 0
                for i in suggest_list:
                    suggest_item += '(%d). %s\n' % (count + 1, i)
                    count += 1

                summary_sheet.range((result_idx[0] + idx, result_idx[1] + 2)).value = fail_item
                summary_sheet.range((result_idx[0] + idx, result_idx[1] + 3)).value = suggest_item
                summary_sheet.range((result_idx[0] + idx, result_idx[1] + 4)).value = '-'

            summary_sheet.range((result_idx[0] + idx, result_idx[1])).value = item

    summary_sheet.range((result_idx[0] + idx + 1, result_idx[1])).value = 'Fail Number:Check Number'
    summary_sheet.range((result_idx[0] + idx + 1, result_idx[1] + 1)).value = '\'%d:%d' % (
        fail_number, total_check_number)
    summary_sheet.range((result_idx[0] + idx + 1, result_idx[1] + 2)).value = '-'
    summary_sheet.range((result_idx[0] + idx + 1, result_idx[1] + 3)).value = '-'
    summary_sheet.range((result_idx[0] + idx + 1, result_idx[1] + 4)).value = '-'

    summary_sheet.range((result_idx[0] + idx + 2, result_idx[1])).value = 'Fail Rate'
    summary_sheet.range((result_idx[0] + idx + 2, result_idx[1] + 1)).value = '%.3f' % (
            float(fail_number) / float(total_check_number) * 100) + '%'
    summary_sheet.range((result_idx[0] + idx + 2, result_idx[1] + 2)).value = '-'
    summary_sheet.range((result_idx[0] + idx + 2, result_idx[1] + 3)).value = '-'
    summary_sheet.range((result_idx[0] + idx + 2, result_idx[1] + 4)).value = '-'

    #############

    summary_sheet.range(result_idx).expand('table').api.Borders.LineStyle = LineStyle.xlContinuous
    SetCellFont_current_region(summary_sheet, result_idx, 'Times New Roman', 12, 'l')
    summary_sheet.range(result_idx).current_region.api.VerticalAlignment = Constants.xlCenter

    summary_sheet.autofit('c')
    summary_sheet.autofit('r')

    # Write datetime for "Cover" Sheet
    t = time.time()
    date = datetime.datetime.fromtimestamp(t).strftime('%Y/%m/%d')
    cover_sheet = wb.sheets['Cover']
    for cell in cover_sheet.api.UsedRange.Cells:

        if str(cell.Value).find('Issue Date:') > -1:
            cover_sheet.range((cell.Row, cell.Column)).value = 'Issue Date: %s' % date

    for cell in cover_sheet.api.UsedRange.Cells:
        if cell.Value == 'DATE':
            cover_sheet.range((cell.Row, cell.Column + 1)).value = date
            cover_sheet.range((cell.Row, cell.Column + 3)).value = date
            cover_sheet.range((cell.Row, cell.Column + 5)).value = date

    end_time = time.clock()

    # #print('GenerateSummary', end_time - start_time)


def ClearSummary():
    wb = Book(xlsm_path).caller()
    summary_sheet = wb.sheets['Summary']
    for cell in summary_sheet.api.UsedRange.Cells:
        if cell.Value == '(1) Layout boardfile:':
            summary_sheet.range((cell.Row, cell.Column + 2)).clear()
    for cell in summary_sheet.api.UsedRange.Cells:
        if cell.Value == '3.Check Results':
            summary_sheet.range((cell.Row + 3, cell.Column)).expand('table').clear()


# set2 中清除按钮
def ClearLayerList(layer_idx=None):
    wb = Book(xlsm_path).caller()
    setting_sheet = wb.sheets['Setting']

    if not layer_idx:
        for cell in setting_sheet.api.UsedRange.Cells:
            if cell.Value == 'Layer Type Definition:':
                layer_idx = (cell.Row + 1, cell.Column)
                break

    setting_sheet.range(layer_idx).expand('table').clear()


# Set2
def LoadStackup():
    wb = Book(xlsm_path).caller()
    setting_sheet = wb.sheets['Setting']
    # 获取报告绝对路径，叠层字典，元件名称列表，progress坐标，叠层名称列表（去重）
    allegro_report_path, layer_type_dict, start_sch_name_list, progress_ind, All_Layer_List = GetSetting()
    # 获取值
    SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data, npr_brd_data = read_allegro_data(
        allegro_report_path)
    layer_list, type_list = getalllayerlist(stackup_brd_data)

    for cell in setting_sheet.api.UsedRange.Cells:
        if cell.Value == 'Layer Type Definition:':
            layer_idx = (cell.Row + 1, cell.Column)
            break

    # 清除 Layer Type Definition 表格的数据
    setting_sheet.range(layer_idx).expand('table').clear()

    # 读取并设置 Layer Type Definition 表格的数据
    # layer_data_list = [[layer_list[idx], 'NA'] if type_list[idx] == 'PLANE' else [layer_list[idx], '-'] for idx in xrange(len(layer_list))]

    # myCode
    # 读取并设置 Layer Type Definition 表格的数据
    # 可多导入几张图，DSL,SL的区别，可优化 ----Gorgeous
    layer_data_list = []
    for idx in xrange(len(layer_list)):
        if type_list[idx] == 'PLANE':
            layer_data_list.append([layer_list[idx], 'NA'])
        # elif re.search('TOP|BOTTOM', layer_list[idx]):
        elif layer_list[idx] == 'TOP' or layer_list[idx] == 'BOTTOM':
            layer_data_list.append([layer_list[idx], 'MS'])
        else:
            layer_data_list.append([layer_list[idx], '-'])
    # myCode
    setting_sheet.range(layer_idx).options(expand='table', ndim=2).value = layer_data_list

    # 设置 Layer Type Definition 表格的背景颜色与边框样式
    setting_sheet.range(layer_idx).expand('table').api.Interior.ColorIndex = 43
    setting_sheet.range(layer_idx).offset(0, 1).expand('table').api.Interior.ColorIndex = 6

    SetCellFont_current_region(setting_sheet, layer_idx, 'Times New Roman', 12, 'l')
    SetCellBorder_current_region(setting_sheet, layer_idx)


# Set3
def LoadTXList():
    wb = Book(xlsm_path).caller()
    setting_sheet = wb.sheets['Setting']

    # 获取 Min. Pin Number 与 Start Component Name List 两个表格的位置坐标
    for cell in setting_sheet.api.UsedRange.Cells:
        if cell.Value == 'Min. Pin Number':
            pin_limit_idx = (cell.Row + 1, cell.Column)
            break
    for cell in setting_sheet.api.UsedRange.Cells:
        if cell.Value == 'Start Component Name List:':
            comp_list_idx = (cell.Row + 1, cell.Column)
            break

    # 最小 pin number 为用户输入
    min_pin_number = setting_sheet.range(pin_limit_idx).value

    # 清除 Start Component Name List 表格的数据
    setting_sheet.range(comp_list_idx).expand('table').clear()

    # 写入数据
    try:
        min_pin_number = int(min_pin_number)
        # 获得键
        allegro_report_path, layer_type_dict, start_sch_name_list, progress_ind, All_Layer_List = GetSetting()
        # 获得值
        SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data, npr_brd_data = read_allegro_data(
            allegro_report_path)
        SCH_content = SCH_brd_data.GetData()

        SCH_dict = dict()
        for line in SCH_content:
            if line[0] != '':
                # 键出现一次值为1，键出现两次值为2
                if SCH_dict.get(line[0]):
                    SCH_dict[line[0]] += 1
                else:
                    SCH_dict[line[0]] = 1

        sch_list = []
        for x in SCH_dict.keys():
            if SCH_dict[x] >= min_pin_number:
                sch_list.append(x)
        # 排序
        sch_list.sort()

        setting_sheet.range(comp_list_idx).expand('table').value = [[x] for x in sch_list]

        SetCellFont_current_region(setting_sheet, comp_list_idx, 'Times New Roman', 12, 'l')
        SetCellBorder_current_region(setting_sheet, comp_list_idx)

        SetCellFont_current_region(setting_sheet, pin_limit_idx, 'Times New Roman', 12, 'l')
        SetCellBorder_current_region(setting_sheet, pin_limit_idx)
    except:
        pass


# set3 的清除按钮
def ClearStartComponentList():
    wb = Book(xlsm_path).caller()
    setting_sheet = wb.sheets['Setting']
    for cell in setting_sheet.api.UsedRange.Cells:
        if cell.Value == 'Start Component Name List:':
            setting_sheet.range((cell.Row + 1, cell.Column)).expand('table').clear()
            break


# DDR通过

###################### BY Gorgeous #######################

################################################
# for DDR tab check
################################################

class TabData:

    def __init__(self, ind, net_name, tab_list, tab_dict, tab_1, tab_2, tab_3, num_1, num_2, num_3):
        self.ind = ind
        self.net_name = net_name
        self.tab_list = tab_list
        self.tab_dict = tab_dict
        self.tab_1 = tab_1
        self.tab_2 = tab_2
        self.tab_3 = tab_3
        self.num_1 = num_1
        self.num_2 = num_2
        self.num_3 = num_3

    def get_ind(self):
        return self.ind

    def get_net_name(self):
        return self.net_name

    def get_tab_list(self):
        return self.tab_list

    def get_tab_dict(self):
        return self.tab_dict

    def get_tab_name_1(self):
        return self.tab_1

    def get_tab_name_2(self):
        return self.tab_2

    def get_tab_name_3(self):
        return self.tab_3

    def get_tab_num_1(self):
        return self.num_1

    def get_tab_num_2(self):
        return self.num_2

    def get_tab_num_3(self):
        return self.num_3


class TabNum:

    def __init__(self, max_ind, min_ind, moving_range, tab_num_list, none_list, tab_name_list_1,
                 tab_name_list_2, tab_name_list_3, tab_num_list_1, tab_num_list_2, tab_num_list_3):
        self.max_ind = max_ind
        self.min_ind = min_ind
        self.moving_range = moving_range
        self.tab_num_list = tab_num_list
        self.none_list = none_list
        self.tab_name_list_1 = tab_name_list_1
        self.tab_name_list_2 = tab_name_list_2
        self.tab_name_list_3 = tab_name_list_3
        self.tab_num_list_1 = tab_num_list_1
        self.tab_num_list_2 = tab_num_list_2
        self.tab_num_list_3 = tab_num_list_3

    def get_max(self):
        return self.max_ind

    def get_min(self):
        return self.min_ind

    def get_moving_range(self):
        return self.moving_range

    def get_tab_num_list(self):
        return self.tab_num_list

    def get_none_list(self):
        return self.none_list

    def get_name_list_1(self):
        return self.tab_name_list_1

    def get_name_list_2(self):
        return self.tab_name_list_2

    def get_name_list_3(self):
        return self.tab_name_list_3

    def get_num_list_1(self):
        return self.tab_num_list_1

    def get_num_list_2(self):
        return self.tab_num_list_2

    def get_num_list_3(self):
        return self.tab_num_list_3


def get_tab_data():
    '''从Excel表中获取table数据'''

    # current_path,a,b,c,d = GetSetting()
    # current_path = current_path.split('/')[0:-1]
    # current_path = [str(x) for x in current_path]
    # current_path.append('test_data.xlsx')
    # current_path = '/'.join(current_path)
    current_path = os.getcwd().split('\\')
    current_path = current_path[0:-1]
    current_path.append('test_data.xlsx')
    current_path = '\\'.join(current_path)
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(current_path)
    sht = wb.sheets['sheet1']
    sheet_value_list = [x.split() for x in sht.range('A1').expand('table').value]
    wb.close()
    app.quit()

    # # print(sheet_value_list)
    return sheet_value_list


def classify_tab_data():
    '''处理table数据'''

    net_name_list = []
    tab_data_list = []
    tab_name_list_1 = []
    tab_name_list_2 = []
    tab_name_list_3 = []
    tab_num_list_1 = []
    tab_num_list_2 = []
    tab_num_list_3 = []
    tab_data_dict = {}
    net_name_ind = 0

    tab_value_list = get_tab_data()

    for tab_value in tab_value_list:

        # # print(1, len(tab_value))
        # # print(2, tab_value)
        net_name_list.append(tab_value[0])
        tab_num_list_1.append(tab_value[1])
        tab_name_list_1.append(tab_value[2])

        if len(tab_value) == 3:
            tab_data_list.append([int(tab_value[1]), 0, 0])
            tab_num_list_2.append(0)
            tab_num_list_3.append(0)
            tab_name_list_2.append('UNKNOWN')
            tab_name_list_3.append('UNKNOWN')
        elif len(tab_value) == 5:
            tab_data_list.append([int(tab_value[1]), int(tab_value[3]), 0])
            tab_num_list_2.append(tab_value[3])
            tab_num_list_3.append(0)
            tab_name_list_2.append(tab_value[4])
            tab_name_list_3.append('UNKNOWN')
        else:
            tab_data_list.append([int(tab_value[1]), int(tab_value[3]), int(tab_value[5])])
            tab_num_list_2.append(tab_value[3])
            tab_num_list_3.append(tab_value[5])
            tab_name_list_2.append(tab_value[4])
            tab_name_list_3.append(tab_value[6])

        # # print(1, tab_data_list)

        tab_data_dict[tab_value[0]] = tab_data_list
        net_name_ind += 1

    return TabData(net_name_ind, net_name_list, tab_data_list, tab_data_dict,
                   tab_name_list_1, tab_name_list_2, tab_name_list_3, tab_num_list_1, tab_num_list_2, tab_num_list_3)


def find_ind(find_data, find_list):
    out_ind_list = []

    for ind in xrange(0, len(find_list)):
        if find_data == find_list[ind]:
            out_ind_list.append(ind)
    return out_ind_list


def compare_tab_num(net_needed_check):
    '''对每组net的tab数量进行比对'''

    tab_num_list = []
    tab_num_no_none_list = []
    tab_name_checked_list_1 = []
    tab_name_checked_list_2 = []
    tab_name_checked_list_3 = []
    tab_num_checked_list_1 = []
    tab_num_checked_list_2 = []
    tab_num_checked_list_3 = []
    tab_name_checked_dict_1 = {}
    tab_name_checked_dict_2 = {}
    tab_name_checked_dict_3 = {}
    tab_num_checked_dict_1 = {}
    tab_num_checked_dict_2 = {}
    tab_num_checked_dict_3 = {}
    tab_value_sum_dict = {}

    tab_data_object = classify_tab_data()
    tab_value_list = tab_data_object.get_tab_list()
    net_name_list = tab_data_object.get_net_name()
    net_name_ind = tab_data_object.get_ind()
    tab_name_list_1 = tab_data_object.get_tab_name_1()
    tab_name_list_2 = tab_data_object.get_tab_name_2()
    tab_name_list_3 = tab_data_object.get_tab_name_3()
    tab_num_list_1 = tab_data_object.get_tab_num_1()
    tab_num_list_2 = tab_data_object.get_tab_num_2()
    tab_num_list_3 = tab_data_object.get_tab_num_3()

    # 计算每条线上的tab总数
    for ind in xrange(net_name_ind):
        # # print(tab_num_list_2[213])
        # # print(net_name_list[ind])
        # # print(tab_num_list_2[ind])

        tab_value_sum_dict[net_name_list[ind]] = tab_value_list[ind][0] + tab_value_list[ind][1] \
                                                 + tab_value_list[ind][2]
        tab_name_checked_dict_1[net_name_list[ind]] = tab_name_list_1[ind]
        tab_name_checked_dict_2[net_name_list[ind]] = tab_name_list_2[ind]
        tab_name_checked_dict_3[net_name_list[ind]] = tab_name_list_3[ind]
        tab_num_checked_dict_1[net_name_list[ind]] = tab_num_list_1[ind]
        tab_num_checked_dict_2[net_name_list[ind]] = tab_num_list_2[ind]
        tab_num_checked_dict_3[net_name_list[ind]] = tab_num_list_3[ind]

    none_ind = -1
    none_list = []

    for net in net_needed_check:
        none_ind += 1
        try:
            tab_num_list.append(tab_value_sum_dict[net.strip()])
            tab_name_checked_list_1.append(tab_name_checked_dict_1[net.strip()])
            tab_name_checked_list_2.append(tab_name_checked_dict_2[net.strip()])
            tab_name_checked_list_3.append(tab_name_checked_dict_3[net.strip()])
            tab_num_checked_list_1.append(tab_num_checked_dict_1[net.strip()])
            tab_num_checked_list_2.append(tab_num_checked_dict_2[net.strip()])
            tab_num_checked_list_3.append(tab_num_checked_dict_3[net.strip()])
        except KeyError:
            none_list.append(none_ind)
            tab_num_list.append('None')
            tab_name_checked_list_1.append('None')
            tab_name_checked_list_2.append('None')
            tab_name_checked_list_3.append('None')
            tab_num_checked_list_1.append('None')
            tab_num_checked_list_2.append('None')
            tab_num_checked_list_3.append('None')

    # # print(tab_num_list)
    for x in tab_num_list:
        if x != 'None':
            tab_num_no_none_list.append(x)

    if len(tab_num_no_none_list) > 0:
        max_num = max(tab_num_no_none_list)
        max_num_ind = find_ind(max_num, tab_num_list)
        min_num = min(tab_num_no_none_list)
        min_num_ind = find_ind(min_num, tab_num_list)

        tab_num_list = [[x] for x in tab_num_list]
        moving_range = max_num - min_num
        # # print(1, max_num)
        # # print(2, min_num)
        return TabNum(max_num_ind, min_num_ind, moving_range, tab_num_list, none_list,
                      tab_name_checked_list_1, tab_name_checked_list_2, tab_name_checked_list_3,
                      tab_num_checked_list_1, tab_num_checked_list_2, tab_num_checked_list_3)
    else:
        tab_num_list = [[x] for x in tab_num_list]
        return TabNum(1, 2, 3, tab_num_list, none_list,
                      tab_name_checked_list_1, tab_name_checked_list_2, tab_name_checked_list_3,
                      tab_num_checked_list_1, tab_num_checked_list_2, tab_num_checked_list_3)


def load_check_tab_gui():
    '''调用tab表格gui'''
    app = QtGui.QApplication(sys.argv)
    form = LoadCheckTabForm()
    form.show()
    app.exec_()


class LoadCheckTabForm(QtGui.QDialog):
    def __init__(self, parent=None):
        super(LoadCheckTabForm, self).__init__(parent)
        self.setWindowTitle('Check Tab Number')
        self.b111 = QtGui.QPushButton('Insert Table', self)
        self.b111.clicked.connect(self.load_table)
        self.b222 = QtGui.QPushButton('Check Tab', self)
        self.b222.clicked.connect(self.check_tab)
        self.b333 = QtGui.QPushButton('Clear Table', self)
        self.b333.clicked.connect(self.clear_table)
        self.b444 = QtGui.QPushButton('Cancel', self)
        self.b444.clicked.connect(self.close_dialog)

        layout = QtGui.QGridLayout()
        layout.addWidget(self.b111, 0, 0)
        layout.addWidget(self.b222, 1, 0)
        layout.addWidget(self.b333, 2, 0)
        layout.addWidget(self.b444, 4, 0)
        self.setLayout(layout)

    def close_dialog(self):
        QtGui.QDialog.accept(self)

    def load_table(self):
        QtGui.QDialog.accept(self)
        load_check_tab_table()

    def check_tab(self):
        QtGui.QDialog.accept(self)
        complete_table_length_item()

    def clear_table(self):
        QtGui.QDialog.accept(self)
        clear_tab_table()


def load_check_tab_table():
    """生成check tab表格"""
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)

    selection_range.value = 'Check Tab'
    active_sheet.range(start_ind).api.Interior.ColorIndex = 44
    active_sheet.range((start_ind[0], start_ind[1] + 1)).value = '1.2/1.3/1.4'
    active_sheet.range((start_ind[0] + 1, start_ind[1])).value = 'Net_Name'
    active_sheet.range((start_ind[0] + 1, start_ind[1])).api.Interior.ColorIndex = 43
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 1)).value = 'Tab_1'
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 1)).api.Interior.ColorIndex = 43
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 2)).value = 'Tab_1_number'
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 2)).api.Interior.ColorIndex = 43
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 3)).value = 'Tab_2'
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 3)).api.Interior.ColorIndex = 43
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 4)).value = 'Tab_2_number'
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 4)).api.Interior.ColorIndex = 43
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 5)).value = 'Tab_3'
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 5)).api.Interior.ColorIndex = 43
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 6)).value = 'Tab_3_number'
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 6)).api.Interior.ColorIndex = 43
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 7)).value = 'Total_Tab_Number'
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 7)).api.Interior.ColorIndex = 43
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 8)).value = 'Total_Tab_Length'
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 8)).api.Interior.ColorIndex = 43
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 9)).value = 'Results'
    active_sheet.range((start_ind[0] + 1, start_ind[1] + 9)).api.Interior.ColorIndex = 43

    SetCellFont_current_region(active_sheet, start_ind, 'Times New Roman', 12, 'l')

    SetCellBorder_current_region(active_sheet, start_ind)
    active_sheet.autofit('c')


def change_list_dimension(target_list):
    target_list = [[x] for x in target_list]
    return target_list


def count_tab_length(tab_1, tab_2, tab_3, factor_list):
    tab_ind = 0
    tab_length = []
    for len_ind in xrange(len(tab_1)):
        if tab_1[tab_ind] == 'None':
            tab_length.append(['None'])
        else:
            tab_length.append([factor_list[0] * int(tab_1[tab_ind]) + factor_list[1] * int(tab_2[tab_ind])
                               + factor_list[2] * int(tab_3[tab_ind])])
        tab_ind += 1
    return tab_length


def check_tab_number():
    '''加载tab number并进行管控'''
    global tab_length, active_sheet, selection_range, start_ind

    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)

    net_name_check_list = []

    factor_list = active_sheet.range(start_ind[0], start_ind[1] + 1).value.split('/')
    factor_list = [float(x) for x in factor_list]

    # 获取需要check的net_name
    if selection_range.value == 'Check Tab':

        net_name_ind = (selection_range.row + 2, selection_range.column)
        active_sheet.range((net_name_ind[0], net_name_ind[1] + 1)).expand('table').clear()
        net_name_check_list = active_sheet.range(net_name_ind).expand('table').value
        try:
            fa = net_name_check_list[0][0][1]
            net_name_check_list = [x[0] for x in net_name_check_list]
        except:
            pass

        # 获取要填入表格的数据
        compare_tab_num_object = compare_tab_num(net_name_check_list)
        max_ind_list = compare_tab_num_object.get_max()
        min_ind_list = compare_tab_num_object.get_min()
        moving_range = compare_tab_num_object.get_moving_range()
        tab_num_list = compare_tab_num_object.get_tab_num_list()
        none_list = compare_tab_num_object.get_none_list()
        tab_name_list_1 = compare_tab_num_object.get_name_list_1()
        tab_name_list_2 = compare_tab_num_object.get_name_list_2()
        tab_name_list_3 = compare_tab_num_object.get_name_list_3()
        tab_num_list_1 = compare_tab_num_object.get_num_list_1()
        tab_num_list_2 = compare_tab_num_object.get_num_list_2()
        tab_num_list_3 = compare_tab_num_object.get_num_list_3()

        net_name_check_list = [[x.strip()] for x in net_name_check_list]

        # 计算tab_length
        tab_length = count_tab_length(tab_num_list_1, tab_num_list_2, tab_num_list_3, factor_list)

        # 改变维度
        tab_name_list_1 = change_list_dimension(tab_name_list_1)
        tab_name_list_2 = change_list_dimension(tab_name_list_2)
        tab_name_list_3 = change_list_dimension(tab_name_list_3)
        tab_num_list_1 = change_list_dimension(tab_num_list_1)
        tab_num_list_2 = change_list_dimension(tab_num_list_2)
        tab_num_list_3 = change_list_dimension(tab_num_list_3)

        active_sheet.range(net_name_ind).expand('table').clear()

        # 将数值填入表中
        active_sheet.range(net_name_ind).value = net_name_check_list
        active_sheet.range(net_name_ind[0], net_name_ind[1] + 1).value = tab_name_list_1
        active_sheet.range(net_name_ind[0], net_name_ind[1] + 2).value = tab_num_list_1
        active_sheet.range(net_name_ind[0], net_name_ind[1] + 3).value = tab_name_list_2
        active_sheet.range(net_name_ind[0], net_name_ind[1] + 4).value = tab_num_list_2
        active_sheet.range(net_name_ind[0], net_name_ind[1] + 5).value = tab_name_list_3
        active_sheet.range(net_name_ind[0], net_name_ind[1] + 6).value = tab_num_list_3
        active_sheet.range(net_name_ind[0], net_name_ind[1] + 7).value = tab_num_list
        active_sheet.range(net_name_ind[0], net_name_ind[1] + 8).value = tab_length

        SetCellFont_current_region(active_sheet, (start_ind[0] + 2, start_ind[1]), 'Times New Roman', 12, 'c')
        SetCellBorder_current_region(active_sheet, start_ind)
        active_sheet.range((start_ind[0] + 1, start_ind[1])).api.HorizontalAlignment = Constants.xlLeft

        # 上色
        for x in xrange(len(tab_num_list)):
            # 设置net_name左对齐
            active_sheet.range(start_ind).api.HorizontalAlignment = Constants.xlLeft
            active_sheet.range((start_ind[0] + 2 + x, start_ind[1])).api.HorizontalAlignment = Constants.xlLeft
            # net_name颜色
            active_sheet.range((start_ind[0] + 2 + x, start_ind[1])).api.Interior.ColorIndex = 42
            active_sheet.range((start_ind[0] + 2 + x, start_ind[1] + 1)).api.Interior.ColorIndex = 45
            active_sheet.range((start_ind[0] + 2 + x, start_ind[1] + 2)).api.Interior.ColorIndex = 45
            active_sheet.range((start_ind[0] + 2 + x, start_ind[1] + 3)).api.Interior.ColorIndex = 45
            active_sheet.range((start_ind[0] + 2 + x, start_ind[1] + 4)).api.Interior.ColorIndex = 45
            active_sheet.range((start_ind[0] + 2 + x, start_ind[1] + 5)).api.Interior.ColorIndex = 45
            active_sheet.range((start_ind[0] + 2 + x, start_ind[1] + 6)).api.Interior.ColorIndex = 45
            active_sheet.range((start_ind[0] + 2 + x, start_ind[1] + 7)).api.Interior.ColorIndex = 27
            active_sheet.range((start_ind[0] + 2 + x, start_ind[1] + 8)).api.Interior.ColorIndex = 37

        # 合并单元格
        active_sheet.range((start_ind[0] + 2, start_ind[1] + 9),
                           (start_ind[0] + 2 + len(tab_num_list) - 1, start_ind[1] + 9)).api.MergeCells = True
        active_sheet.range((start_ind[0] + 2, start_ind[1] + 9)).api.VerticalAlignment = \
            Constants.xlCenter

        if moving_range > 2:
            # for x in xrange(len(tab_num_list)):
            active_sheet.range((start_ind[0] + 2, start_ind[1] + 9)).value = 'Fail'
            active_sheet.range((start_ind[0] + 2, start_ind[1] + 9)).api.Interior.ColorIndex = 3
        else:
            active_sheet.range((start_ind[0] + 2, start_ind[1] + 9)).value = 'Pass'
            active_sheet.range((start_ind[0] + 2, start_ind[1] + 9)).api.Interior.ColorIndex = 4

        # # print(none_list)
        for ind in none_list:
            active_sheet.range((start_ind[0] + 2 + ind, start_ind[1] + 1)).api.Interior.ColorIndex = 15
            active_sheet.range((start_ind[0] + 2 + ind, start_ind[1] + 2)).api.Interior.ColorIndex = 15
            active_sheet.range((start_ind[0] + 2 + ind, start_ind[1] + 3)).api.Interior.ColorIndex = 15
            active_sheet.range((start_ind[0] + 2 + ind, start_ind[1] + 4)).api.Interior.ColorIndex = 15
            active_sheet.range((start_ind[0] + 2 + ind, start_ind[1] + 5)).api.Interior.ColorIndex = 15
            active_sheet.range((start_ind[0] + 2 + ind, start_ind[1] + 6)).api.Interior.ColorIndex = 15
            active_sheet.range((start_ind[0] + 2 + ind, start_ind[1] + 7)).api.Interior.ColorIndex = 15
            active_sheet.range((start_ind[0] + 2 + ind, start_ind[1] + 8)).api.Interior.ColorIndex = 15
            active_sheet.range((start_ind[0] + 2 + ind, start_ind[1] + 9)).api.Interior.ColorIndex = 15
            active_sheet.range((start_ind[0] + 2 + ind, start_ind[1] + 9)).value = 'None'
        active_sheet.autofit('c')


def complete_table_length_item():
    global tab_length, active_sheet, start_ind, key_cell_list
    DQS_real_length_list = []
    DQ_real_length_list = []

    check_tab_number()

    DQS_table_ind = getPreTable(2)
    DQ_table_ind = getPreTable()

    # # print(DQS_table_ind, DQ_table_ind)

    DQS_length_value_list = get_pre_table_length_list(active_sheet, DQS_table_ind)
    DQ_length_value_list = get_pre_table_length_list(active_sheet, DQ_table_ind)
    # # print(tab_length)
    # # print(DQ_length_value_list)
    # # print(key_cell_list)

    # 计算length + tab length 的值
    if len(DQ_length_value_list) < len(tab_length):
        DQS_len = len(DQS_length_value_list)
        for i in xrange(DQS_len):
            DQS_real_length_list.append([round(DQS_length_value_list[i] + tab_length[i][0], 2)])
        for n in xrange(len(DQ_length_value_list)):
            DQ_real_length_list.append([round(DQ_length_value_list[n] + tab_length[n + DQS_len][0], 2)])

        DQS_Tab_length_list = []
        DQ_Tab_length_list = []
        for i in xrange(len(DQ_real_length_list)):
            dqt = round(max(DQ_real_length_list)[0] - DQ_real_length_list[i][0], 2)
            dqt1 = round(DQ_real_length_list[i][0] - min(DQ_real_length_list)[0], 2)
            DQ_Tab_length_list.append(max(dqt, dqt1))

            dqst = round(max(DQS_real_length_list)[0] - DQ_real_length_list[i][0], 2)
            dqst1 = round(DQ_real_length_list[i][0] - min(DQS_real_length_list)[0], 2)
            DQS_Tab_length_list.append(max(dqst, dqst1))
        # 分别为DQS与DQ表格创建Total length +  Tab length项
        load_pre_length_table(DQS_table_ind, key_cell_list, DQS_real_length_list, False)
        load_pre_length_table(DQ_table_ind, key_cell_list, DQ_real_length_list)
        # 为DQ与DQS Tab Spec创建列表
        check_Tab_DQS_DQ_Spec(DQ_table_ind, DQ_Tab_length_list, DQS_Tab_length_list)
    else:
        #  只有一个Topology表时加载 Tab+Toal length 列
        for n in xrange(len(DQ_length_value_list)):
            DQ_real_length_list.append([round(DQ_length_value_list[n] + tab_length[n][0], 2)])

        load_pre_length_table(DQ_table_ind, key_cell_list, DQ_real_length_list)
        # 为CTL_CMD 创建列表项
        check_CTL_CMD_TabLength(DQ_table_ind, DQ_real_length_list)


def check_CTL_CMD_TabLength(start_table_ind, CTL_CMD_real_length_list):
    global tab_length, active_sheet

    # 获取CLK信号的 Total Length
    CLK_table_ind = getFirstTable()
    CLK_value_list = []
    if CLK_table_ind:
        topology_table = active_sheet.range(CLK_table_ind).current_region.value
        for ind in xrange(len(topology_table[2])):
            if topology_table[2][ind] == 'Total Length':
                len_ind = ind
        # 获取DQS长度数据
        for value in topology_table[12:]:
            CLK_value_list.append(value[len_ind])

    if active_sheet.range(start_table_ind).value == 'Topology':
        topology_table = active_sheet.range(start_table_ind).current_region.value
        DQ_length_ind = (start_table_ind[0] + 2, start_table_ind[1] + len(topology_table[0]))
        cmd_ctl_cell_list = []
        # 获取管控制
        for ind in xrange(len(topology_table[2])):
            if topology_table[2][ind] == 'CMD or ADD to CLK Length Matching':
                len_ind = ind
            elif topology_table[2][ind] == 'CTL to CLK Length Matching':
                len_ind = ind
        for value in topology_table[3:12]:
            cmd_ctl_cell_list.append(value[len_ind])

        if topology_table[2][-1] == None:
            DQ_length_ind = (DQ_length_ind[0], DQ_length_ind[1] - 1)

        # 计算CMD CTL Tab Length Matching
        cmd_ctl_real_spec = []
        for idx in xrange(len(CTL_CMD_real_length_list)):
            dqt = round(max(CLK_value_list) - CTL_CMD_real_length_list[idx][0], 2)
            dqt1 = round(CTL_CMD_real_length_list[idx][0] - min(CLK_value_list), 2)
            cmd_ctl_real_spec.append(max(dqt, dqt1))
        # 刷新表单

        # 填入相应的值
        if topology_table[2][len_ind] == 'CMD or ADD to CLK Length Matching':
            active_sheet.range(DQ_length_ind).value = 'CMD or ADD to CLK (Tab Length) Matching'
        if topology_table[2][len_ind] == 'CTL to CLK Length Matching':
            active_sheet.range(DQ_length_ind).value = 'CTL to CLK (Tab Length) Matching'
        active_sheet.range((DQ_length_ind[0] + 1, DQ_length_ind[1])).value = change_list_dimension(cmd_ctl_cell_list)
        for x in xrange(len(cmd_ctl_cell_list) + 1):
            active_sheet.range((DQ_length_ind[0] + x, DQ_length_ind[1])).api.Interior.Color = RgbColor.rgbSkyBlue
        # 填入数值
        active_sheet.range((DQ_length_ind[0] + 10, DQ_length_ind[1])).value = change_list_dimension(cmd_ctl_real_spec)

        # 通过管控制上色
        # # print(cmd_ctl_cell_list[7])
        for x in xrange(len(cmd_ctl_real_spec)):
            if cmd_ctl_real_spec[x] > cmd_ctl_cell_list[7]:
                active_sheet.range(DQ_length_ind[0] + 10 + x, DQ_length_ind[1]).api.Interior.ColorIndex = 3
            else:
                active_sheet.range(DQ_length_ind[0] + 10 + x, DQ_length_ind[1]).api.Interior.ColorIndex = 4

        # 合并单元格

        active_sheet.range((start_table_ind[0], start_table_ind[1] + 2),
                           (start_table_ind[0] + 1, DQ_length_ind[1])).api.MergeCells = True
        SetCellFont_current_region(active_sheet, start_table_ind, 'Times New Roman', 12, 'l')
        SetCellBorder_current_region(active_sheet, start_table_ind)


def check_Tab_DQS_DQ_Spec(DQ_table_ind, DQ_Tab_length_list, DQS_Tab_length_list):
    global tab_length, active_sheet
    dq_cell_list = []
    dqs_cell_list = []
    if active_sheet.range(DQ_table_ind).value == 'Topology':
        topology_table = active_sheet.range(DQ_table_ind).current_region.value
        DQ_length_ind = (DQ_table_ind[0] + 2, DQ_table_ind[1] + len(topology_table[0]))

        # 获取管控数值
        for ind in xrange(len(topology_table[2])):
            if topology_table[2][ind] == 'Relative Length Spec(DQ to DQ)':
                len_ind = ind
            if topology_table[2][ind] == 'Relative Length Spec(DQ to DQS)':
                len_ind1 = ind
        for value in topology_table[3:12]:
            dq_cell_list.append(value[len_ind])
            dqs_cell_list.append(value[len_ind1])

        if topology_table[2][-1] == None:
            DQ_length_ind = (DQ_length_ind[0], DQ_length_ind[1] - 2)
        # # print(DQ_length_ind)

        active_sheet.range(DQ_length_ind).value = 'DQ Tab length Spec'
        active_sheet.range(DQ_length_ind[0], DQ_length_ind[1] + 1).value = 'DQS Tab length Spec'
        active_sheet.range((DQ_length_ind[0] + 1, DQ_length_ind[1])).value = change_list_dimension(dq_cell_list)
        active_sheet.range((DQ_length_ind[0] + 1, DQ_length_ind[1] + 1)).value = change_list_dimension(dqs_cell_list)

        for x in xrange(len(key_cell_list) + 1):
            active_sheet.range((DQ_length_ind[0] + x, DQ_length_ind[1])).api.Interior.Color = RgbColor.rgbSkyBlue
            active_sheet.range((DQ_length_ind[0] + x, DQ_length_ind[1] + 1)).api.Interior.Color = RgbColor.rgbSkyBlue
        # 填入length数值
        active_sheet.range((DQ_length_ind[0] + 10, DQ_length_ind[1])).value = change_list_dimension(DQ_Tab_length_list)
        active_sheet.range((DQ_length_ind[0] + 10, DQ_length_ind[1] + 1)).value = change_list_dimension(
            DQS_Tab_length_list)
        # # print(dq_cell_list[7])
        # # print(DQ_Tab_length_list)

        # 通过管控值上色
        for x in xrange(len(DQ_Tab_length_list)):
            if DQ_Tab_length_list[x] > dq_cell_list[7]:
                active_sheet.range(DQ_length_ind[0] + 10 + x, DQ_length_ind[1]).api.Interior.ColorIndex = 3
            else:
                active_sheet.range(DQ_length_ind[0] + 10 + x, DQ_length_ind[1]).api.Interior.ColorIndex = 4
            if DQS_Tab_length_list[x] > dqs_cell_list[7]:
                active_sheet.range(DQ_length_ind[0] + 10 + x, DQ_length_ind[1] + 1).api.Interior.ColorIndex = 3
            else:
                active_sheet.range(DQ_length_ind[0] + 10 + x, DQ_length_ind[1] + 1).api.Interior.ColorIndex = 4
        # # print(DQ_length_ind[1])

        SetCellFont_current_region(active_sheet, DQ_table_ind, 'Times New Roman', 12, 'l')
        SetCellBorder_current_region(active_sheet, DQ_table_ind)
        # 合并单元格
        active_sheet.range((DQ_table_ind[0], DQ_table_ind[1] + 3),
                           (DQ_table_ind[0] + 1, DQ_length_ind[1] + 1)).api.MergeCells = True

        # # 合并 DQ Tab length Spec,DQS Tab length Spec 单元格
        # active_sheet.range((DQ_table_ind[0] + 12, DQ_length_ind[1]),
        #                    (DQ_table_ind[0] + len(topology_table)-1, DQ_length_ind[1])).api.MergeCells = True
        # active_sheet.range((DQ_table_ind[0] + 12, DQ_length_ind[1]+1),
        #                    (DQ_table_ind[0] + len(topology_table)-1, DQ_length_ind[1]+1)).api.MergeCells = True
        # 内容居中
        active_sheet.range(DQ_table_ind[0] + 12, DQ_length_ind[1]).api.VerticalAlignment = Constants.xlCenter
        active_sheet.range(DQ_table_ind[0] + 12, DQ_length_ind[1] + 1).api.VerticalAlignment = Constants.xlCenter


def load_pre_length_table(pre_table_ind, key_cell_list, real_length_list, DQ_flag=True):
    global tab_length, active_sheet
    key_cell_list = change_list_dimension(key_cell_list)
    my_flag = 0

    if active_sheet.range(pre_table_ind).value == 'Topology':
        topology_table = active_sheet.range(pre_table_ind).current_region.value
        pre_length_ind = (pre_table_ind[0] + 2, pre_table_ind[1] + len(topology_table[0]))

        # 如果之前有表格就删除
        if DQ_flag:
            # 如果添加DQ,DQS项，清除相应位置上的数据
            for ind in xrange(len(topology_table[2])):
                try:
                    if topology_table[2][ind].find('DQS Tab length Spec') > -1:
                        my_flag = 1
                        active_sheet.range((pre_length_ind[0], pre_length_ind[1] - 3),
                                           (pre_length_ind[0] + len(topology_table) - 3, pre_length_ind[1] - 1)).clear()
                    elif topology_table[2][ind].find('Total length + Total Tab length') > -1:
                        my_flag = 2
                        active_sheet.range((pre_length_ind[0], pre_length_ind[1]),
                                           (pre_length_ind[0] + len(topology_table) - 3, pre_length_ind[1] + 2)).clear()
                    elif topology_table[2][ind].find('CTL to CLK (Tab Length) Matching') > -1 or topology_table[2][
                        ind].find('CMD or ADD to CLK (Tab Length) Matching') > -1:
                        my_flag = 3
                        active_sheet.range((pre_length_ind[0], pre_length_ind[1] - 2),
                                           (pre_length_ind[0] + len(topology_table) - 3, pre_length_ind[1] - 1)).clear()

                except AttributeError:
                    pass
            if my_flag == 1:
                pre_length_ind = (pre_length_ind[0], pre_length_ind[1] - 3)
            elif my_flag == 2:
                pre_length_ind = (pre_length_ind[0], pre_length_ind[1] - 1)
            elif my_flag == 3:
                pre_length_ind = (pre_length_ind[0], pre_length_ind[1] - 2)
        else:
            for ind in xrange(len(topology_table[2])):
                try:
                    if topology_table[2][ind].find('Total length + Total Tab length') > -1:
                        my_flag = 1
                        active_sheet.range((pre_length_ind[0], pre_length_ind[1] - 1),
                                           (pre_length_ind[0] + len(topology_table) - 3, pre_length_ind[1] - 1)).clear()
                except AttributeError:
                    pass

            if my_flag:
                pre_length_ind = (pre_length_ind[0], pre_length_ind[1] - 1)

        active_sheet.range(pre_length_ind).value = 'Total length + Total Tab length'
        active_sheet.range((pre_length_ind[0] + 1, pre_length_ind[1])).value = key_cell_list

        for x in xrange(len(key_cell_list) + 1):
            active_sheet.range((pre_length_ind[0] + x, pre_length_ind[1])).api.Interior.Color = RgbColor.rgbSkyBlue

        # 填入length数值
        active_sheet.range((pre_length_ind[0] + 10, pre_length_ind[1])).value = real_length_list
        # 比较是否通过管控进行上色
        for x in xrange(len(real_length_list)):
            active_sheet.range((pre_length_ind[0] + 10 + x, pre_length_ind[1])).api.Interior.ColorIndex = 4

        length_section = [key_cell_list[6][0], key_cell_list[7][0]]
        if length_section[0] != 'NA' and length_section[1] == 'NA':
            for x in xrange(len(real_length_list)):
                if real_length_list[x][0] < length_section[0]:
                    active_sheet.range((pre_length_ind[0] + 10 + x, pre_length_ind[1])).api.Interior.ColorIndex = 3
        elif length_section[0] == 'NA' and length_section[1] != 'NA':
            for x in xrange(len(real_length_list)):
                if real_length_list[x][0] > length_section[1]:
                    active_sheet.range((pre_length_ind[0] + 10 + x, pre_length_ind[1])).api.Interior.ColorIndex = 3
        elif length_section[0] != 'NA' and length_section[1] != 'NA':
            for x in xrange(len(real_length_list)):
                if real_length_list[x][0] > length_section[1] or real_length_list[x][0] < length_section[0]:
                    active_sheet.range((pre_length_ind[0] + 10 + x, pre_length_ind[1])).api.Interior.ColorIndex = 3

        # 设置表格格式
        SetCellFont_current_region(active_sheet, pre_table_ind, 'Times New Roman', 12, 'l')
        SetCellBorder_current_region(active_sheet, pre_table_ind)
        # 合并单元格
        active_sheet.range((pre_table_ind[0], pre_table_ind[1] + 2),
                           (pre_table_ind[0] + 1, pre_length_ind[1])).api.MergeCells = True


def clear_tab_table():
    '''清除tab表格'''
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row + 2, selection_range.column)

    active_sheet.range(start_ind).expand('table').clear()


def get_pre_table_length_list(active_sheet, pre_table_ind):
    '''获得DQ表格上两行的DQS表格中的length 数据'''
    global key_cell_list
    pre_length_value_list = []

    if active_sheet.range(pre_table_ind).value == 'Topology':

        key_cell_list = []
        topology_table = active_sheet.range(pre_table_ind).current_region.value
        for ind in xrange(len(topology_table[2])):
            if topology_table[2][ind] == 'Total Length':
                len_ind = ind
        # 获取DQ长度数据及DQ表格result数据
        for value in topology_table[3:12]:
            key_cell_list.append(value[len_ind])
        for value in topology_table[12:]:
            pre_length_value_list.append(value[len_ind])
        return pre_length_value_list


# 解决DIMM到DIMM之间线长管控

def load_count_DIMM_gui():
    '''调用tab表格gui'''
    app = QtGui.QApplication(sys.argv)
    form = LoadCountDimmForm()
    form.show()
    app.exec_()


class LoadCountDimmForm(QtGui.QDialog):
    def __init__(self, parent=None):
        super(LoadCountDimmForm, self).__init__(parent)
        self.setWindowTitle('Count DIMM TO DIMM')
        self.b111 = QtGui.QPushButton('Insert DIMM Table', self)
        self.b111.clicked.connect(self.load_dimm_table)
        self.b222 = QtGui.QPushButton('Calc DIMM', self)
        self.b222.clicked.connect(self.calc_dimm)
        self.b333 = QtGui.QPushButton('Cancel', self)
        self.b333.clicked.connect(self.close_dialog)

        layout = QtGui.QGridLayout()
        layout.addWidget(self.b111, 0, 0)
        layout.addWidget(self.b222, 1, 0)
        layout.addWidget(self.b333, 2, 0)
        self.setLayout(layout)

    def close_dialog(self):
        QtGui.QDialog.accept(self)

    def load_dimm_table(self):
        QtGui.QDialog.accept(self)
        load_count_dimm_table()

    def calc_dimm(self):
        QtGui.QDialog.accept(self)
        count_dimm_length()


def load_count_dimm_table():
    '''生成DIMM_TO_DIMM表格'''
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)
    my_flag = 0

    if selection_range.value == 'Topology':
        topology_table = selection_range.current_region.value
        dimm_ind = (start_ind[0] + 2, start_ind[1] + len(topology_table[0]))

        # 如果之前有表格就删除
        for ind in xrange(len(topology_table[2])):
            if topology_table[2][ind].find('DIMM TO DIMM Length') > -1 or \
                    topology_table[2][ind].find('DIMM TO DIMM Length') > -1:
                my_flag = 1
                active_sheet.range((dimm_ind[0], dimm_ind[1] - 1),
                                   (dimm_ind[0] + len(topology_table) - 3, dimm_ind[1] - 1)).clear()

        # 创建DIMM表格
        DIMM_cell_value_list = [['DIMM TO DIMM Length'], ['NA'], ['NA'], ['NA'], ['NA'], ['NA'],
                                ['NA'], ['NA'], ['NA'], ['500']]

        if my_flag:
            dimm_ind = (dimm_ind[0], dimm_ind[1] - 1)

        active_sheet.range((dimm_ind[0], dimm_ind[1])).value = DIMM_cell_value_list

        for x in xrange(len(DIMM_cell_value_list)):
            active_sheet.range((dimm_ind[0] + x, dimm_ind[1])).api.Interior.Color = RgbColor.rgbSkyBlue

        # 设置表格格式
        SetCellFont_current_region(active_sheet, start_ind, 'Times New Roman', 12, 'l')
        SetCellBorder_current_region(active_sheet, start_ind)

        # 合并单元格
        active_sheet.range((start_ind[0], start_ind[1] + 2), (start_ind[0] + 1, dimm_ind[1])).api.MergeCells = True


def count_dimm_length():
    '''对DIMM_TO_DIMM之间length长度进行管控'''
    wb = Book(xlsm_path).caller()
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)
    my_flag = 0
    net_name_list = []
    line_needed_list = []
    dimm_length_list = []

    if selection_range.value == 'Topology':
        topology_table = selection_range.current_region.value
        signal_type_ind = (start_ind[0] + 1, start_ind[1] + 1)
        signal_type = active_sheet.range(signal_type_ind).value
        dimm_ind = (start_ind[0] + 2, start_ind[1] + len(topology_table[2]))

        # 从表格中获取net_name
        if topology_table[11][2] == 'Net Name':
            for x in xrange(12, len(topology_table)):
                net_name_list.append(topology_table[x][2])

        # 判断之前有无剩余表格
        for ind in xrange(len(topology_table[2])):

            # 如果存在之前运行的数据，则清除表格
            if topology_table[2][ind].find('DIMM TO DIMM Length') > -1 or \
                    topology_table[2][ind].find('DIMM TO DIMM Length') > -1:
                my_flag = 1
                active_sheet.range((dimm_ind[0] + 10, dimm_ind[1] - 1),
                                   (dimm_ind[0] + len(topology_table) - 3, dimm_ind[1] - 1)).clear()
        if my_flag:
            dimm_ind = (dimm_ind[0], dimm_ind[1] - 1)

        # 获取DIMM_TO_DIMM线长
        if signal_type == 'Differential':
            act_sheet = wb.sheets['ACT_diff']

        elif signal_type == 'Single-ended':
            act_sheet = wb.sheets['ACT_se']

        act_all = act_sheet.range('A1').current_region.value

        line_list = []
        for net in net_name_list:
            for x in act_all:
                if x[1] == net:
                    line_list.append(x)

        for line in line_list:
            if line[0].find('XMM') > -1 and line[2].find('XMM') > -1:
                line_needed_list.append(line)
        # 排除重复的情况
        line_needed_list = line_needed_list[::2]

        for line in line_needed_list:
            value_ind = 0
            length_value = 0
            for value in line:
                value_ind += 1
                if str(value).find(':') > -1:
                    length_value += line[value_ind]

            dimm_length_list.append(length_value)

        dimm_length_list = change_list_dimension(dimm_length_list)
        # 填入DIMM表格数据
        active_sheet.range((dimm_ind[0] + 10, dimm_ind[1])).value = dimm_length_list
        # 比较是否通过管控进行上色
        DIMM_TO_DIMM_guide = active_sheet.range((dimm_ind[0] + 9, dimm_ind[1])).value

        fail_ind = -1
        for x in dimm_length_list:
            fail_ind += 1
            if x[0] > DIMM_TO_DIMM_guide:
                active_sheet.range((dimm_ind[0] + 10 + fail_ind, dimm_ind[1])) \
                    .api.Interior.ColorIndex = 3
            else:
                active_sheet.range((dimm_ind[0] + 10 + fail_ind, dimm_ind[1])) \
                    .api.Interior.ColorIndex = 4
        # 合并单元格
        SetCellFont_current_region(active_sheet, start_ind, 'Times New Roman', 12, 'l')
        SetCellBorder_current_region(active_sheet, start_ind)


# Command line argument Definition
if __name__ == '__main__':

    for i in sys.argv:
        if i.find('.xlsm') > -1:
            xlsm_path = r'%s' % i
            break
    # xlsm_path = r'C:\Users\Tommy\Desktop\Software_project\bug\20190828\L_IB460CX_S_PCH_V_2DPC_DT_SIM_Checklist_A1.8_20190827.xlsm'
    # xlsm_path = r'C:\Users\Tommy\Desktop\Software_project\bug\20190912\PyACT_for_Checklist_template2.0.xlsm'
    # xlsm_path = r'F:\brd_checklist_debug_20190806\Z2_Penghu-ED_CML_WS_4Layer_SIM_Checklist_A1.1_20190715.xlsm'
    Book(xlsm_path).set_mock_caller()

    getpinnumio()
    reload(sys)
    sys.setdefaultencoding('UTF-8')
    for i in sys.argv:
        if i == 'read_allegro_report':
            LoadAllegroFile()
            break
        elif i == 'net_type_detect':
            NetTypeDetect()
            break
        elif i == 'clear_net_list':
            ClearNetList()
            break
        # elif i == 'symbol_type_detect':
        #     SymbolListDetect()
        #     break
        elif i == 'run_signal_topology':
            RunSignalTopology()
            break
        elif i == 'load_start_end_sch':
            LoadStartEndComponent()
            break
        elif i == 'check_topology':
            CheckTopology()
            break
        elif i == 'check_all_topology':
            RunAllTopologyCheck()
            break
        elif i == 'check_pass_fail_only':
            ShowResults()
            break
        elif i == 'clear_check_results':
            ClearCheckResults()
            break
        elif i == 'format_convert':
            FormatConverter()
            break
        elif i == 'load_topology_format_gui':
            LoadTopologyFormat_GUI()
            break
        elif i == 'load_diff_topology':
            LoadTopologyFormat('Differential')
            break
        elif i == 'load_se_topology':
            LoadTopologyFormat('Single-ended')
            break
        elif i == 'delete_topology_table':
            DeleteTopologyTable()
            break
        elif i == 'load_simple_topology_gui':
            LoadSimpleTopology_GUI()
            break
        elif i == 'load_simple_topology_format':
            LoadSimpleTopologyFormat()
            break
        elif i == 'load_simple_topology':
            LoadSimpleTopology()
            break
        elif i == 'generate_summary':
            GenerateSummary()
            break
        elif i == 'clear_summary':
            ClearSummary()
            break
        elif i == 'load_layer_list':
            LoadStackup()
            break
        elif i == 'clear_layer_list':
            ClearLayerList()
            break
        elif i == 'load_start_sch_list':
            LoadTXList()
            break
        elif i == 'clear_start_sch_list':
            ClearStartComponentList()
            break
        elif i == 'power_trace':
            power_trace()
            break
        elif i == 'power_clear':
            power_clear()
            break
        elif i == 'check_tab':
            load_check_tab_gui()
            break
        elif i == 'count_dimm':
            load_count_DIMM_gui()
            break
        elif i == 'calc_grade':
            load_calc_grade_gui()
            break
        elif i == 'batch_update':
            BatchUpdate_Topology()
            break

# load_calc_grade_gui()
# CheckTopology()
# BatchUpdate_Topology()
# LoadTopologyFormat_GUI()
# LoadSimpleTopologyFormat()
# LoadSimpleTopology()
# RunSignalTopology()
# load_calc_grade_gui()
# complete_table_length_item()
# load_calc_grade_gui()
# LoadAllegroFile()
# LoadStackup()
# LoadTXList()
# RunSignalTopology()
# LoadAllegroFile()
# read_allegro_data(r'C:\Users\Tommy\Desktop\Software_project\bug\D10_Bison+_X02-0415-1000.brd_allegro_report.rpt')
# NetTypeDetect()
