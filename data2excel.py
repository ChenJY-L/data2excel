import re

import xlwings.utils
from dateutil import parser
from pathlib import Path
from fnmatch import translate
import sys
# import killpywintypes
import xlwings as xw  # excel操作模块
import datetime  # 时间模块
# import time
import os
import numpy as np
from PyQt6 import QtWidgets
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtGui import QIcon
import gui as QTUI  # 此处需要根据需要改名，第二个变量不改也可以
import ico01
import ctypes

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myappid")


class GUI_Dialog(QDialog, QTUI.Ui_Data_Processing):
    def __init__(self, parent=None):
        super(GUI_Dialog, self).__init__(parent)
        self.setupUi(self)
        current_file = os.path.basename(__file__)
        current_file = os.path.splitext(current_file)[0][-8:]
        self.setWindowTitle('Data_Processing ' + current_file)
        self.setWindowIcon(QIcon(':/favicon01.ico'))
        self.DyBCCheckBox.setChecked(True)

        # 功能连接区
        self.Process.clicked.connect(self.DataProcess)
        self.FileSelect.clicked.connect(self.FileSelectF)

        # excel计算初始化
        # xlwings打开
        self.xwapp = xw.App(visible=False, add_book=False)
        self.xwapp.display_alerts = False
        self.xwapp.screen_updating = False

        # 绘图颜色配置
        self.colors = ['#60966d','#5b3660','#018abe','#e90f44','#63adee','#924c63','#7c8fa3']

    # 功能区
    # 界面刷新
    def GuiRefresh(self, textbox, text):
        textbox.setPlainText(text)
        QApplication.processEvents()

    # hex -> rgb -> int
    def hexColor2Int(self, color):
        return xw.utils.rgb_to_int(xw.utils.hex_to_rgb(color))

    # 文件选择
    def FileSelectF(self):
        self.GuiRefresh(self.Status, 'Selecting Path')
        SchPath = os.getcwd()
        if self.Path.toPlainText() != '':
            SchPath = os.path.dirname(self.Path.toPlainText())
        filenames = (QFileDialog.getOpenFileNames(self, 'Select', SchPath, "Data Files(*.dat);;All Files(*)"))
        if filenames[0] != []:
            self.Path.setPlainText(filenames[0][0])

    # 文件路径获取
    def FilePath(self, path, filename):
        filenames = os.listdir(path)
        filenamesnew = []
        for i in range(0, len(filenames)):
            if filenames[i].find(filename) >= 0:
                filenamesnew.append(filenames[i])
        if len(filenamesnew) == 0:
            basefilename = ''
        else:
            basefilename = filenamesnew[len(filenamesnew) - 1]
        return os.path.join(path, basefilename)

    # 检查sheet是否存在
    def CheckSheet(self, workbook, sheetname):
        num = len(workbook.sheets)  # 获取sheet个数
        x = []
        for sc in range(0, num):
            if sc >= 0:
                sht = workbook.sheets[sc]
                x.append(sht.name)
            else:
                pass
        try:
            ind = x.index(sheetname)
        except (ValueError, ArithmeticError):
            ind = -1
        if ind == -1:
            return False
        else:
            return True

            # 定义查找获取单元格函数

    def FindRowColRange(self, SheetName, Rttype, KeyWord, RangeStr):
        try:
            Cell_Address = SheetName.range(RangeStr).api.Find(What=KeyWord, LookAt=xw.constants.LookAt.xlWhole)
            Cell_Row = Cell_Address.Row
            Cell_Col = Cell_Address.Column
            Cell_Ran = str(Cell_Row) + ',' + str(Cell_Col)
            Cell_Adr = SheetName.range((Cell_Row, Cell_Col)).get_address(False, False)
        except:
            Cell_Adr = 'A0'
            Cell_Row = 0
            Cell_Col = 0
            Cell_Ran = '0,0'
        if Rttype == 'Adr':
            Cell_Result = Cell_Adr
        elif Rttype == 'Row':
            Cell_Result = str(Cell_Row)
        elif Rttype == 'Col':
            Cell_Result = str(Cell_Col)
        elif Rttype == 'Ran':
            Cell_Result = Cell_Ran
        return Cell_Result

    # 删除特定元素
    def popele(self, arrayin, ele):
        arrayout = []
        for index in range(len(arrayin) - 1, -1, -1):
            if arrayin[index] != ele & np.isnan(arrayin[index]):
                arrayout.append(arrayin[index])
        arrayout = np.asarray(arrayout)
        return arrayout


    """ 标注文本解析函数 """
    def parseText(self, text):
        """ 从excel中读取时间和行为
            判断依据是分号，如果一行的开头是中文，则读取第一个分号为分界点，分号之前是行为，分号之后的内容分解为时间
            如果一行的开头是数字，则读取最后一个分号，分号之后是行为，分号之前是时间
            时间段和时间点都会被处理成excel中的时间段（即小数形式）
            当读取到一行的开头为备注时，后续的全部内容都会被读取为备注信息
            返回类型为dict
        """
        schedule = []
        remark = []
        current_section = "schedule"  # 初始状态

        for line in text.strip().splitlines():
            line = line.strip().replace('：', ':')  # 统一冒号
            if not line or line.startswith('#'):
                continue

            if line.startswith("备注"):
                current_section = "remark"
                remark.append(line[3:].strip())  # 直接添加，无需额外判断
                continue

            if current_section == "remark":
                remark.append(line)
                continue

            # 定位冒号
            split_pos = line.rfind(':') if re.match(r'^\d', line) else line.find(':')

            if split_pos != -1:
                if re.match(r'^\d', line):
                    time_str, activity = line[:split_pos].strip(), line[split_pos + 1:].strip()
                else:
                    activity, time_str = line[:split_pos].strip(), line[split_pos + 1:].strip()
                # 处理时间段（例如 "17:20-17:23"）
                if '-' in time_str:
                    try:
                        start_time, end_time = time_str.split('-')
                        start_datetime = parser.parse(start_time.strip())
                        end_datetime = parser.parse(end_time.strip())

                        # 将 datetime 转为 Excel 时间
                        start_excel_time = (start_datetime.hour * 3600 + start_datetime.minute * 60 + start_datetime.second) / 86400
                        end_excel_time = (end_datetime.hour * 3600 + end_datetime.minute * 60 + end_datetime.second) / 86400

                        schedule.append({"time": (start_excel_time, end_excel_time), "activity": activity})
                    except ValueError:
                        schedule.append({"time": None, "activity": line})
                else:  # 单个时间
                    try:
                        t = parser.parse(time_str)
                        excel_time = (t.hour * 3600 + t.minute * 60 + t.second) / 86400
                        schedule.append({"time": excel_time, "activity": activity})
                    except ValueError:
                        schedule.append({"time": None, "activity": line})
            else:  # schedule 部分没找到冒号，整行都是活动
                schedule.append({"time": None, "activity": line})

        return {
            "schedule": schedule,
            "remark": "\n".join(remark).strip() if remark else None
        }

    # 数据处理专区
    def DataProcess(self):
        try:
            self.Process.setEnabled(False)
            self.starttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, '')
            # 先确定路径
            if self.Path.toPlainText() == '':
                self.GuiRefresh(self.Status, 'Obtaining File Path')
                path = os.getcwd()  # 这是当前文件位置
                filekwords = self.Original.currentText() + ')-' + 'Ch1.dat'
                if self.Original.currentText() == 'Calibrated':
                    filekwordsC = '校准后' + ')-' + '1环.dat'
                elif self.Original.currentText() == 'Original':
                    filekwordsC = '原始值' + ')-' + '1环.dat'
                self.Path.setPlainText(max(self.FilePath(path, filekwords), self.FilePath(path, filekwordsC)))
            # 确定环数
            if self.Rings.currentText() == '5 Rings':
                Ch = 5
            elif self.Rings.currentText() == '7 Rings':
                Ch = 7
            # 获取基本数据
            # 获取各环数据
            self.GuiRefresh(self.Status, 'Loading Data')
            Chpath = self.Path.toPlainText()
            Chpath = Chpath.replace('/', '\\')
            if Chpath.find('环') > 0:
                C = True
            elif Chpath.find('-Ch') > 0:
                C = False
            Chvalues = []
            for i in range(1, Ch + 1):
                match = re.match(r'^(.*?)-\d+环\.dat$', Chpath)
                if match:
                    prefix = match.group(1)
                    datakword = f"{prefix}-{i}环.dat" if C else f"{prefix}-Ch{i}.dat"

                wbd = self.xwapp.books.open(datakword)
                rawsheet = wbd.sheets[0]
                Chvalues.append(rawsheet.range(rawsheet.used_range).value)
                self.GuiRefresh(self.Status, 'Loading Single Ring Data' + str(i))
                wbd.close()
            Chvalues = np.asarray(Chvalues)
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

            # 获取温度数据
            Tempvalue = []
            if self.TempCheckBox.isChecked() == True:
                self.GuiRefresh(self.Status, 'Loading Temperature Data')
                temppath = os.path.split(self.Path.toPlainText())[0]
                tempfilekwords = '温度' if C else 'Temperature'
                wbt = self.xwapp.books.open(self.FilePath(temppath, tempfilekwords))
                tempsheet = wbt.sheets[0]
                Tempvalue = tempsheet.range(tempsheet.used_range).value
                Temptitle = Tempvalue[0][:]
                Tempvalue = Tempvalue[1:][:]
                Temptitle = np.asarray(Temptitle)
                Tempvalue = np.asarray(Tempvalue)

            # 获取标志信息
            filePath = os.path.dirname(Chpath)
            txtFiles = list(Path(filePath).glob('*.txt'))  # 获取所有 .txt 文件列表
            if txtFiles:
                # 读取第一个txt文件
                text = Path(txtFiles[0]).read_text(encoding='utf-8')
                expInfo = self.parseText(text)
            else:
                # 没有txt文件，跳过
                pass

            # 先把其他基本数据计算好（时间、序列）
            n = len(Chvalues[0])  # 单环数据行数
            m = len(Chvalues[0][0])  # 每次测量数加一（最后一列是时间）
            datarange = [i for i in range(n)]
            if self.LDCheckBox.isChecked() == True:
                wave = ['1064', '1310', '1390', '1550', '1625']
                for i in range(len(datarange) - 1, -1, -1):
                    if (datarange[i] % 6 == 5):
                        del datarange[i]
            else:
                wave = ['1050', '1219', '1314', '1409', '1550', '1609']
            n = len(datarange)
            wn = len(wave)
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])
            # 单环和差分环的数组
            ringwords = []
            diffwords = []
            cs = -1
            for r in range(0, Ch):
                ringwords.append((str(r + 1) + '环') if C else ('Ring' + str(r + 1)))
                for rl in range(r + 1, Ch):
                    cs = cs + 1
                    diffwords.append('Diff' + str(r + 1) + str(rl + 1))

            # 获取时间和周期序列
            timearr = np.empty((int(n // wn), 1), dtype=np.float64)
            cycleNoarr = np.empty((int(n // wn), 1))
            self.GuiRefresh(self.Status, 'Loading Title')
            for l in range(0, n // wn):
                timepop = self.popele(Chvalues[0][6 * l:6 * l + wn][:, m - 1], 0)
                timeele = sum(timepop) / len(timepop)
                timearr[l] = timeele
                cycleNoarr[l] = l + 1

            # 提前算好基准周期单环和差分数据
            self.GuiRefresh(self.Status, 'Calculating Base Data')
            basesingle = np.zeros((Ch, wn))  # 6列（波长）环数行（环）
            diffNo = int(Ch * (Ch - 1) // 2)
            basediff = np.zeros((diffNo, wn))
            for w in range(0, wn):
                cs = -1
                for r in range(0, Ch):
                    basesingle[r][w] = sum(Chvalues[r][(self.BaseCycle.value() - 1) * 6 + w][1:m - 1]) / (m - 2)
                    for rl in range(r + 1, Ch):
                        cs = cs + 1
                        basediff[cs][w] = sum(np.log(Chvalues[r][(self.BaseCycle.value() - 1) * 6 + w][1:m - 1] / Chvalues[rl][(self.BaseCycle.value() - 1) * 6 + w][1:m - 1])) / (m - 2)

            # 新建文件，用以放置处理好的数据，先判断文件是否存在，新建文件，打开，再判断sheet是否存在，如否，新建sheet
            # self.filestarttime=datetime.datetime.now()
            self.GuiRefresh(self.Status, 'Creating Output File')
            ProcessFilePath = os.path.join(Chpath.replace(Chpath.split('\\')[-1], ''),
                                           Chpath.split('\\')[-2] + '.xlsx') if C else (
                                           Chpath.split('Ch')[0] + 'Processed' + '.xlsx')
            if os.path.isfile(ProcessFilePath) == False:
                wb = self.xwapp.books.add()  # 在app下创建一个Book
                wb.save(ProcessFilePath)
            wb = self.xwapp.books.open(ProcessFilePath)
            # 新建sheet
            sheetnames = ['单环', '单环吸光度', '单环信噪比', '差分', '差分吸光度', '差分等效信噪比',
                          '光强和信噪比汇总', '温度数据'] if C else ['Single Ring', 'Single Ring Absorbance',
                                                                     'Single SNR', 'Differential',
                                                                     'Differential Absorbance', 'Differential SNR',
                                                                     'Summary of Intensity and SNR', 'Temperature']
            self.GuiRefresh(self.Status, 'Creating Sheets')
            if self.TempCheckBox.isChecked() == True:
                sheetNo = len(sheetnames)
            else:
                sheetNo = len(sheetnames) - 1
            for i in range(0, sheetNo):
                if self.CheckSheet(wb, sheetnames[i]) == False:
                    wb.sheets.add(sheetnames[i], after=i + 1)
            wb.save()
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

            # 下面正式处理数据
            # 先统一在所有sheet中写入时间和序号
            # self.titlestarttime=datetime.datetime.now()
            self.GuiRefresh(self.Status, 'Writing Title')
            for s in range(0, 6):
                wb.sheets[sheetnames[s]].range(3, 1).value = timearr
                wb.sheets[sheetnames[s]].range(3, 2).value = cycleNoarr
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

            # 接下来进行温度插值和写入
            # self.tempstarttime=datetime.datetime.now()
            try:
                yinterp = np.empty((timearr.shape[0], Tempvalue.shape[1] - 1))
            except:
                yinterp = np.empty((timearr.shape[0], 1))

            if self.TempCheckBox.isChecked():
                wb.sheets[sheetnames[len(sheetnames) - 1]].range(2, 1).value = timearr
                wb.sheets[sheetnames[len(sheetnames) - 1]].range(2, 2).value = cycleNoarr
                wb.sheets[sheetnames[len(sheetnames) - 1]].range(1, 2).value = Temptitle
                Tempvalue.shape
                for i in range(1, Tempvalue.shape[1]):
                    yinterp[:, i - 1] = np.interp(timearr, Tempvalue[:, 0], Tempvalue[:, i]).reshape(timearr.shape[0], )
                wb.sheets[sheetnames[len(sheetnames) - 1]].range(2, 3).value = yinterp
                self.GuiRefresh(self.Status, 'Writing Temp Data')
            wb.save()
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

            # 接下来写入光强和信噪比数据汇总的题头
            self.GuiRefresh(self.Status, 'Writing Summery Data')
            wb.sheets[sheetnames[6]].range(3, 3).value = '光强' if C else 'Intensity'
            wb.sheets[sheetnames[6]].range(3, 13).value = '信噪比' if C else 'SNR'
            wb.sheets[sheetnames[6]].range(4, 5).value = wave
            wb.sheets[sheetnames[6]].range(4, 15).value = wave
            wb.sheets[sheetnames[6]].range(5, 4).options(transpose=True).value = ringwords
            wb.sheets[sheetnames[6]].range(5, 14).options(transpose=True).value = ringwords
            wb.sheets[sheetnames[6]].range(5 + Ch, 4).options(transpose=True).value = diffwords
            wb.sheets[sheetnames[6]].range(5 + Ch, 14).options(transpose=True).value = diffwords
            wb.save()

            # 单环数据和单环吸光度数据
            # self.singstarttime=datetime.datetime.now()
            # =LN(@INDIRECT("单环!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&($C$4+2))/@INDIRECT("单环!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&ROW()))
            for r in range(0, Ch):
                # 数据头
                for s in range(0, 3):
                    wb.sheets[sheetnames[s]].range(1, 3 + 7 * r).value = ringwords[r]
                    wb.sheets[sheetnames[s]].range(2, 4 + 7 * r).value = wave
                # 各列数据
                self.GuiRefresh(self.Status, 'Writing Ring ' + str(r + 1))
                singlearr = []
                singleabsarr = []
                singlesnrarr = []
                for j in datarange:  # len(Chvalues[0])或者n
                    singles = Chvalues[r][j][1:m - 1]
                    single = sum(singles) / (m - 2)
                    singleabs = np.log(basesingle[r][j % 6] / single)
                    singlesnr = single / np.std(singles, ddof=1)
                    singlearr.append(single)
                    singleabsarr.append(singleabs)
                    singlesnrarr.append(singlesnr)
                singlearr = np.array([singlearr]).reshape(n // wn, wn)
                singleabsarr = np.array([singleabsarr]).reshape(n // wn, wn)
                singlesnrarr = np.array([singlesnrarr]).reshape(n // wn, wn)
                singleave = singlearr.mean(axis=0)
                singlesnrave = singlesnrarr.mean(axis=0)
                wb.sheets[sheetnames[0]].range(3, 4 + 7 * r).value = singlearr
                wb.sheets[sheetnames[1]].range(3, 4 + 7 * r).value = singleabsarr
                wb.sheets[sheetnames[2]].range(3, 4 + 7 * r).value = singlesnrarr
                wb.sheets[sheetnames[6]].range(5 + r, 5).value = singleave
                wb.sheets[sheetnames[6]].range(5 + r, 15).value = singlesnrave
                self.currenttime = datetime.datetime.now()
                self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])
            wb.save()
            self.GuiRefresh(self.Status, 'Writing Single Finished')

            # 差分数据和差分吸光度数据，也加入了数据汇总的数据
            # =@INDIRECT("差分!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&ROW())-@INDIRECT("差分!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&($C$4+2))
            cs = -1
            for r in range(0, Ch):
                for rl in range(r + 1, Ch):
                    # 数据头
                    cs = cs + 1
                    for s in range(3, 6):
                        wb.sheets[sheetnames[s]].range(1, 3 + 7 * cs).value = diffwords[cs]
                        wb.sheets[sheetnames[s]].range(2, 4 + 7 * cs).value = wave
                    # 各列数据
                    self.GuiRefresh(self.Status, 'Writing Diff ' + str(r + 1) + str(rl + 1))
                    diffarr = np.empty((n, 1))
                    diffabsarr = np.empty((n, 1))
                    diffsnrarr = np.empty((n, 1))
                    for j in datarange:  # len(Chvalues[0])或者n
                        diffs = np.log(Chvalues[r][j][1:m - 1] / Chvalues[rl][j][1:m - 1])

                        diff = sum(diffs) / (m - 2)
                        diffabs = diff - basediff[cs][j % 6]
                        diffsnr = 1 / np.std(diffs, ddof=1)
                        # diff = np.array([sum(diffs)/(m-2)])
                        # diffsnr = np.array([1/np.std(diffs, ddof=1)])
                        # diffabs = np.array(diff-basediff[cs][j%6])
                        diffarr[j] = diff
                        diffabsarr[j] = diffabs
                        diffsnrarr[j] = diffsnr
                        # diffarr = np.concatenate((diffarr, diff),axis=0)
                        # diffabsarr = np.concatenate((diffabsarr, diffabs),axis=0)
                        # diffsnrarr = np.concatenate((diffsnrarr, diffsnr),axis=0)

                    diffarr = diffarr.reshape(n // wn, wn)
                    diffabsarr = np.array([diffabsarr]).reshape(n // wn, wn)
                    diffsnrarr = diffsnrarr.reshape(n // wn, wn)
                    diffarrave = diffarr.mean(axis=0)
                    diffsnrave = diffsnrarr.mean(axis=0)
                    wb.sheets[sheetnames[3]].range(3, 4 + 7 * cs).value = diffarr
                    wb.sheets[sheetnames[4]].range(3, 4 + 7 * cs).value = diffabsarr
                    wb.sheets[sheetnames[5]].range(3, 4 + 7 * cs).value = diffsnrarr
                    wb.sheets[sheetnames[6]].range(5 + Ch + cs, 5).value = diffarrave
                    wb.sheets[sheetnames[6]].range(5 + Ch + cs, 15).value = diffsnrave
                    self.currenttime = datetime.datetime.now()
                    self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])
            wb.save()
            self.GuiRefresh(self.Status, 'Writing Diff Finished')

            if self.DyBCCheckBox.isChecked() == True:
                wb.sheets[sheetnames[1]].range(3, 3).value = '单环基准周期' if C else 'Single Base Cycle'
                wb.sheets[sheetnames[1]].range(4, 3).value = self.BaseCycle.value()
                wb.sheets[sheetnames[4]].range(3, 3).value = '差分基准周期' if C else 'Single Base Cycle'
                wb.sheets[sheetnames[4]].range(4, 3).value = self.BaseCycle.value()
                ss = '=LN(@INDIRECT("' + wb.sheets[
                    sheetnames[0]].name + '!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&($C$4+2))/@INDIRECT("' + wb.sheets[
                         sheetnames[0]].name + '!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&ROW()))'
                singleabsarr = np.full((n // wn, wn), ss)
                ds = '=@INDIRECT("' + wb.sheets[
                    sheetnames[3]].name + '!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&ROW())-@INDIRECT("' + wb.sheets[
                         sheetnames[3]].name + '!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&($C$4+2))'
                diffabsarr = np.full((n // wn, wn), ds)
                for r in range(0, Ch):
                    self.GuiRefresh(self.Status, 'Writing Dyna Ring ' + str(r + 1))
                    wb.sheets[sheetnames[1]].range(3, 4 + 7 * r).value = singleabsarr
                for rl in range(0, int(Ch * (Ch - 1) / 2)):
                    # 数据头
                    self.GuiRefresh(self.Status, 'Writing Dyna Diff ' + str(rl + 1))
                    wb.sheets[sheetnames[4]].range(3, 4 + 7 * rl).value = diffabsarr
                    self.currenttime = datetime.datetime.now()
                    self.GuiRefresh(self.ErrorText,
                                    'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])
                self.GuiRefresh(self.Status, 'Saving...')
                wb.save()

            try:
                rng_lcol = len(yinterp[1]) + 2
            except:
                print("No temperature data")
                rng_lcol = 2
            if self.OGTTCheckBox.isChecked():
                wb.sheets[sheetnames[7]].range(1, rng_lcol + 2).value = '血糖值' if C else 'Glucose Value'
                wb.sheets[sheetnames[7]].range(2, rng_lcol + 1).value = timearr[0]
                wb.sheets[sheetnames[7]].range(3, rng_lcol + 1).value = timearr[int(len(timearr) / 2)]
                wb.sheets[sheetnames[7]].range(4, rng_lcol + 1).value = timearr[int(len(timearr) - 1)]
                wb.sheets[sheetnames[7]].range(2, rng_lcol + 2).value = 5.5
                wb.sheets[sheetnames[7]].range(3, rng_lcol + 2).value = 10.5
                wb.sheets[sheetnames[7]].range(4, rng_lcol + 2).value = 5.5
                gcols = xw.utils.col_name(rng_lcol + 1) + ':' + xw.utils.col_name(rng_lcol + 1)
                wb.sheets[sheetnames[7]].api.Columns(gcols).NumberFormatLocal = "[$-x-systime]h:mm:ss AM/PM"

            if self.PLTCheckBox.isChecked() == True:
                for i in range(0, len(sheetnames)):
                    wb.sheets[sheetnames[i]].api.Columns("A:A").NumberFormatLocal = "[$-x-systime]h:mm:ss AM/PM"
                # 开始绘图
                # 常量赋值
                lft = 400  # 图表左距
                tp = 400  # 图表上距
                caw = 700  # 图表宽度
                cah = 400  # 图表高度
                ttftsz = 28  # 标题字体大小
                axftsz = 18  # 轴字体大小
                axttftsz = 18  # 轴标题字体大小
                gsftsz = 18  # 图例字体大小
                self.charts = []
                # 绘图信息区，后面要修改就改这里
                charttitles = ['12环差分信号vs.室温', '23环差分信号vs.测头旁皮肤温度',
                               '1050nm单环吸光度', '1219nm单环吸光度',
                               '34环差分信号vs.加热功率', '45环差分信号vs.测头下实际温度',
                               '1314nm单环吸光度', '1409nm单环吸光度',
                               '1050nm差分吸光度vs.测头下实际温度', '1550nm差分吸光度 - 1050nm差分吸光度',
                               '1550nm单环吸光度', '1609nm单环吸光度',
                               ]
                ringsindex = ['Diff12', 'Diff23', '1050', '1219',
                              'Diff34', 'Diff45', '1314', '1409',
                              'Diff1050', 'Diff1550-Diff1050', '1550', '1609']
                tempindex = ['4', '5', '0', '0',
                             '15', '12', '0', '0',
                             '12', '0', '0', '0']    # 对应sheet中的列，设置为0则不设置副坐标轴

                infoindex = [False, False, False, False,
                             True, True, True, True,
                             False, False, False, False]

                if self.OGTTCheckBox.isChecked() == True:  # OGTT时的血糖值绘制准备
                    tempindex[5] = str(rng_lcol + 2)
                    tempindex[10] = str(rng_lcol + 2)
                    charttitles[5] = '45环差分信号vs.血糖真值'
                    charttitles[10] = '1550nm单环吸光度vs.血糖真值'
                # 提前算好X轴每格大小
                # xmaxunit = int((timearr[int(len(timearr)-1)]-timearr[0])*24)/100

                pltN = len(charttitles)
                SRRange = 'A1:ZZ2'
                diffSheet = wb.sheets[sheetnames[4]]
                sglSheet = wb.sheets[sheetnames[1]]
                tempSheet = wb.sheets[sheetnames[7]]

                for p, each in enumerate(ringsindex):
                    if each == '0':
                        self.charts.append(None)
                        continue

                    # 数据源范围
                    PltRangeS = 'A:A'

                    if len(each) == 6:
                        datasheet = diffSheet
                        addr = self.FindRowColRange(datasheet, 'Col', each, SRRange)
                        addrstr = xw.utils.col_name(int(addr) + 1) + ':' + xw.utils.col_name(int(addr) + wn)
                        PltRangeS = PltRangeS + ', ' + addrstr
                        ytitle = 'ΔAd'
                    elif len(each) == 4:
                        datasheet = sglSheet
                        addr = self.FindRowColRange(datasheet, 'Col', each, SRRange)
                        for i in range(0, Ch):
                            addrstr = xw.utils.col_name(int(addr) + 7 * i) + ':' + xw.utils.col_name(int(addr) + 7 * i)
                            PltRangeS = PltRangeS + ', ' + addrstr
                            ytitle = 'ΔA'
                    elif len(each) == 3:
                        datasheet = sglSheet
                        realeach = each.replace("-", "")
                        addr = self.FindRowColRange(datasheet, 'Col', realeach, SRRange)
                        addrstr = xw.utils.col_name(int(addr) + 1) + ':' + xw.utils.col_name(int(addr) + wn)
                        PltRangeS = PltRangeS + ', ' + addrstr
                        ytitle = '单环吸光度变化量'
                    elif len(each) <= 2:
                        datasheet = tempSheet
                        addr = xw.utils.col_name(int(each))
                        addrstr = addr + ':' + addr
                        PltRangeS = PltRangeS + ', ' + addrstr
                        ytitle = wb.sheets[sheetnames[7]].range(1, int(ringsindex[p])).value
                    elif len(each) == 8:
                        # example: diff1050
                        datasheet = diffSheet
                        waveIndex = wave.index(each[-4:]) + 1
                        ytitle = 'ΔAd'
                        for i in range(1, Ch):
                            target = 'Diff' + str(i) + str(i + 1)
                            addr = self.FindRowColRange(datasheet, 'Col', target, SRRange)
                            addrstr = xw.utils.col_name(int(addr) + waveIndex) + ':' + xw.utils.col_name(int(addr) + waveIndex)
                            PltRangeS = PltRangeS + ', ' + addrstr
                    elif len(each) == 17:
                        # example: diff1550-diff1550
                        datasheet = diffSheet
                        ytitle = 'ΔAd'
                        wave1, wave2 = each[4:8], each[13:17]
                        waveIndex1, waveIndex2 = wave.index(wave1) + 1, wave.index(wave2) + 1

                        targets = ['Diff12', 'Diff23', 'Diff34', 'Diff45']
                        sheet_target = wb.sheets[sheetnames[4]]

                        # 动态获取最后一列索引，并计算新的列名
                        last_col = sheet_target.used_range.last_cell.column
                        indices = [xw.utils.col_name(last_col + 2 + i) for i in range(len(targets))]
                        sheet_target.range(xw.utils.col_name(last_col + 1) + '1').value = wave1 + ' - ' + wave2

                        # Get length of columns
                        # for target, indice in zip(targets, indices):
                        #     base_addr = int(self.FindRowColRange(datasheet, 'Col', target, SRRange))
                        #     addr1, addr2 = base_addr + waveIndex1, base_addr + waveIndex2
                        #
                        #     addrstr1, addrstr2 = f"{xw.utils.col_name(addr1)}:{xw.utils.col_name(addr1)}", \
                        #         f"{xw.utils.col_name(addr2)}:{xw.utils.col_name(addr2)}"
                        #
                        #     values1 = datasheet.range(addrstr1).options(np.array, ndim=2, empty=np.nan).value
                        #     values2 = datasheet.range(addrstr2).options(np.array, ndim=2, empty=np.nan).value
                        #     diff_values = values1 - values2
                        #
                        #     # 批量写入 Excel
                        #     sheet_target.range(f"{indice}:{indice}").value = diff_values
                        #     sheet_target.range(f"{indice}1").value = None
                        #     sheet_target.range(f"{indice}2").value = target
                        # 在 Excel 中插入计算公式
                        for target, indice in zip(targets, indices):
                            base_addr = int(self.FindRowColRange(datasheet, 'Col', target, SRRange))
                            addr1, addr2 = xw.utils.col_name(base_addr + waveIndex1), xw.utils.col_name(
                                base_addr + waveIndex2)

                            # 绝对引用公式
                            formula = f"=${addr1}2 - ${addr2}2"

                            # 填充整个列（从 2 到最后一行）
                            sheet_target.range(f"{indice}2:{indice}{datasheet.used_range.last_cell.row}").formula = formula

                            # 设置表头信息
                            sheet_target.range(f"{indice}1").value = None
                            sheet_target.range(f"{indice}2").value = target

                        PltRangeS += f", {indices[0]}1:{indices[-1]}{datasheet.used_range.last_cell.row}"
                        self.charts.append(None)
                        continue  # 仅处理数据但是不画图了

                    pltrange = datasheet.range(PltRangeS)  # 此处要改*************************************此处已改

                    # 副坐标轴数据范围
                    if int(tempindex[p]) != 0:
                        secaddr = xw.utils.col_name(int(tempindex[p]))
                        secaddrstr = secaddr + ':' + secaddr
                        SecRangeS = secaddrstr
                        secrange = tempSheet.range(SecRangeS)

                    self.GuiRefresh(self.Status, 'Plotting ' + str(p + 1) + '/' + str(pltN))

                    """ 为了解决diff1050 和diff1550-diff1050绘图位置的问题 """
                    # figure_cah = cah if len(each) < 8 else cah/2
                    figure_lft = (lft + caw * int(p % 4)) if len(each) < 8 else lft + caw
                    figure_top = (tp + cah * int(p / 4)) if len(each) <= 8 else (tp + cah * int(p / 4)) + cah/2
                    self.charts.append(
                        wb.sheets[sheetnames[4]].charts.add(left=figure_lft,
                                                            top=figure_top,
                                                            width=caw,
                                                            height=cah))  # 此处要改*************************************此处已改

                    # 设置图标类型，数据来源，图例位置
                    self.charts[p].chart_type = 'xy_scatter_lines'  # 设置图标类型是xy散点连线图
                    self.charts[p].set_source_data(pltrange)
                    chartApi = self.charts[p].api[1]
                    chartApi.Legend.Position = -4107

                    # 添加副坐标轴
                    series_count = chartApi.SeriesCollection().Count
                    if int(tempindex[p]) > (rng_lcol):
                        chartApi.SeriesCollection().Add(Source=secrange.api, SeriesLabels=True)
                        series_count = chartApi.SeriesCollection().Count
                        chartApi.FullSeriesCollection(series_count).Name = "=" + sheetnames[7] + "!" + xw.utils.col_name(rng_lcol + 2) + "1"
                        chartApi.FullSeriesCollection(series_count).XValues = "=" + sheetnames[7] + "!" + xw.utils.col_name(rng_lcol + 1) + ":" + xw.utils.col_name(rng_lcol + 1)
                        chartApi.FullSeriesCollection(series_count).Values = "=" + sheetnames[7] + "!" + xw.utils.col_name(rng_lcol + 2) + ":" + xw.utils.col_name(rng_lcol + 2)
                        chartApi.SeriesCollection(series_count).AxisGroup = 2
                        chartApi.Axes(2, 2).ReversePlotOrder = True
                        chartApi.ChartColor = 10
                        chartApi.FullSeriesCollection(series_count).Format.Line.ForeColor.RGB = 255
                        chartApi.FullSeriesCollection(series_count).MarkerBackgroundColor = 255
                        chartApi.FullSeriesCollection(series_count).MarkerForegroundColor = 255
                    else:
                        if int(tempindex[p]) != 0:
                            chartApi.SeriesCollection().Add(Source=secrange.api, SeriesLabels=True)  # 此处要改*************************************此处已改
                            chartApi.ChartColor = 10
                            series_count = chartApi.SeriesCollection().Count
                            chartApi.SeriesCollection(series_count).AxisGroup = 2

                    # series_count = self.charts[p].api[1].SeriesCollection().Count
                    # 修改chart系列标记，循环迭代每个系列
                    for i in range(1, chartApi.SeriesCollection().Count + 1):
                        series = chartApi.SeriesCollection(i)
                        # 修改每个系列的标记类型和大小
                        series.MarkerStyle = 8  # 标记类型为圆形
                        series.MarkerSize = 5  # 标记大小为5

                    # 图表整体样式
                    chartApi.ChartArea.Format.Line.ForeColor.RGB = 14277081  # 在VBA立即窗口中输入 ?RGB(217, 217, 217) 回车可查看其数值
                    # 三个坐标轴格式、颜色、线宽
                    # 之前的颜色为14277081
                    chartApi.Axes(1).MajorTickMark = xlwings.constants.Constants.xlCross
                    chartApi.Axes(1).Format.Line.Weight = 1.5
                    chartApi.Axes(1).Format.Line.ForeColor.RGB = 0
                    chartApi.Axes(2, 1).MajorTickMark = xlwings.constants.Constants.xlInside
                    chartApi.Axes(2, 1).Format.Line.Weight = 1.5
                    chartApi.Axes(2, 1).Format.Line.ForeColor.RGB = 0
                    chartApi.Axes(2, 1).MinorUnit = 0.001  if ytitle == 'ΔAd' else 0.01

                    if int(tempindex[p]) != 0:
                        chartApi.Axes(2, 2).MajorTickMark = xlwings.constants.Constants.xlInside
                        chartApi.Axes(2, 2).Format.Line.Weight = 1.5
                        chartApi.Axes(2, 2).Format.Line.ForeColor.RGB = 0
                    # X轴刻度间隔和格式
                    # self.charts[p].api[1].Axes(1).MajorUnit = xmaxunit
                    chartApi.Axes(1).TickLabels.NumberFormatLocal = "h:mm;@"
                    # 两个grid格式、颜色、线宽
                    chartApi.SetElement(334)
                    chartApi.SetElement(330)
                    chartApi.Axes(1).MajorGridlines.Format.Line.ForeColor.RGB = 14277081
                    chartApi.Axes(1).MajorGridlines.Format.Line.Weight = 0.75
                    chartApi.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = 14277081
                    chartApi.Axes(2).MajorGridlines.Format.Line.Weight = 0.75

                    # 标题，字体及大小
                    chartApi.SetElement(2)
                    chartApi.ChartTitle.Format.TextFrame2.TextRange.Characters.Text = charttitles[p]  # 此处要改*************************************此处已改
                    chartApi.ChartTitle.Format.TextFrame2.TextRange.Font.Name = "Calibri"
                    chartApi.ChartTitle.Format.TextFrame2.TextRange.Characters.Font.Size = ttftsz  # 大小
                    chartApi.ChartTitle.Format.TextFrame2.TextRange.Characters.Font.Bold = 1

                    # y坐标轴标题, 字体及大小
                    # TODO: 设置坐标轴加粗（1.5），设置刻度，黑色
                    chartApi.Axes(2, 1).HasTitle = True
                    chartApi.Axes(2, 1).AxisTitle.Characters.Text = ytitle  # 此处要改*************************************此处已改
                    chartApi.Axes(2, 1).AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Calibri"
                    chartApi.Axes(2, 1).AxisTitle.Format.TextFrame2.TextRange.Font.Size = axttftsz
                    chartApi.Axes(2, 1).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = 1

                    if int(tempindex[p]) != 0:
                        # y2坐标轴标题, 字体及大小
                        chartApi.Axes(2, 2).HasTitle = True
                        chartApi.Axes(2, 2).AxisTitle.Characters.Text = wb.sheets[sheetnames[7]].range(1, int(tempindex[p])).value  # 此处要改*************************************此处已改
                        chartApi.Axes(2, 2).AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Calibri"
                        chartApi.Axes(2, 2).AxisTitle.Format.TextFrame2.TextRange.Font.Size = axttftsz
                        chartApi.Axes(2, 2).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = 1

                    # 三个坐标轴字体大小 设置坐标轴字体加粗
                    chartApi.Axes(1).TickLabels.Font.Size = axftsz
                    chartApi.Axes(1).TickLabels.Font.Bold = 1
                    chartApi.Axes(2, 1).TickLabels.Font.Size = axftsz
                    chartApi.Axes(2, 1).TickLabels.Font.Bold = 1
                    if int(tempindex[p]) != 0:
                        chartApi.Axes(2, 2).TickLabels.Font.Size = axftsz
                        chartApi.Axes(2, 2).TickLabels.Font.Bold = 1

                    # 图例
                    chartApi.Legend.Format.TextFrame2.TextRange.Font.Size = gsftsz - 1
                    chartApi.Legend.Format.TextFrame2.TextRange.Font.Bold = 1

                    # 添加标志中的内容
                    series_count = chartApi.SeriesCollection().Count
                    series_count += 0 if int(tempindex[p]) != 0 else 1

                    if 'expInfo' in locals().keys() and (self.expInfoCheckBox.isChecked() or infoindex[p]):
                        for nitem, item in enumerate(expInfo['schedule']):
                            t = item['time'] + np.floor(timearr[0])
                            # 添加新的数据序列
                            y_min = chartApi.Axes(2, 1).MinimumScale
                            y_max = chartApi.Axes(2, 1).MaximumScale
                            series = chartApi.SeriesCollection().NewSeries()
                            itemColor = self.hexColor2Int(self.colors[nitem % len(self.colors)])
                            if len(t) == 2:
                                # 画一个框
                                pointIdx = 3    # 读取框框右上角的点
                                # itemColor = 255

                                series.XValues = [t[0], t[0], t[1], t[1], t[0]]  # 明确转换为数组
                                series.Values = [y_min, y_max, y_max, y_min, y_min]
                            elif len(t) == 1:
                                # 画一条线
                                pointIdx = 2    # 读取上面的点
                                # itemColor = 16711680

                                series.XValues = [t[0], t[0]]  # 明确转换为数组
                                series.Values = [y_min, y_max]
                            else:
                                # 未知类型，不处理
                                raise ValueError(f'Error time length {len(t)} for {item["activity"]}')
                            chartApi.Axes(2, 1).MinimumScale = y_min
                            chartApi.Axes(2, 1).MaximumScale = y_max

                            # 设置序列的格式 (直线，颜色，粗细等)
                            series.ChartType = 75
                            series.Format.Line.Weight = 2.5  # 线宽度
                            series.Format.Line.DashStyle = 4    # 绘制虚线
                            series.Format.Line.ForeColor.RGB = itemColor

                            series.Points(pointIdx).ApplyDataLabels()
                            series.Points(pointIdx).DataLabel.Text = item["activity"]
                            # series.Points(pointIdx).DataLabel.Font.Size = axftsz + 2
                            series.Points(pointIdx).DataLabel.Font.Size = 28
                            series.Points(pointIdx).DataLabel.Font.Bold = 1
                            series.Points(pointIdx).DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = itemColor
                            leg = chartApi.Legend.LegendEntries(series_count)
                            leg.Delete()

                    # 运行时间和过程记录
                    self.currenttime = datetime.datetime.now()
                    self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

                if 'expInfo' in locals().keys():
                    self.GuiRefresh(self.Status, 'Adding remarks')
                    textbox = wb.sheets[sheetnames[4]].shapes.api.AddTextbox(
                        Orientation=1,  # 文本框方向（1=水平）
                        Left=lft,  # 左边距
                        Top=tp + cah*2,    # 顶边距
                        Width=caw/2,       # 宽度
                        Height=cah,       # 高度
                    )
                    # 设置文本和格式
                    textbox.TextFrame2.TextRange.Characters.Text = "备注:\n" + expInfo['remark'] if expInfo['remark'] is not None else "备注:\n"
                    textbox.TextFrame2.TextRange.ParagraphFormat.Alignment = 2  # 居中
                    textbox.TextFrame2.TextRange.Characters.Font.Name = "Times New Roman"
                    textbox.TextFrame2.TextRange.Characters.Font.Size = 40
                    textbox.TextFrame2.TextRange.Characters.Font.Bold = 1

                # 添加标题
                titlebox = wb.sheets[sheetnames[4]].shapes.api.AddTextbox(
                    Orientation=1,  # 文本框方向（1=水平）
                    Left=lft,  # 左边距
                    Top=tp - 100,  # 顶边距
                    Width=caw*3,  # 宽度
                    Height=100,  # 高度
                )
                titlebox.Fill.ForeColor.RGB = xw.utils.rgb_to_int((255, 255, 0))
                titlebox.TextFrame2.TextRange.Characters.Text = Chpath.split('\\')[-2]
                titlebox.TextFrame2.VerticalAnchor = 3  # 居中
                titlebox.TextFrame2.TextRange.Characters.Font.Name = "Times New Roman"
                titlebox.TextFrame2.TextRange.Characters.Font.Size = 54
                titlebox.TextFrame2.TextRange.Characters.Font.Bold = 1

                self.GuiRefresh(self.Status, 'Saving...')
                wb.save()

            # 最后，保存关闭并输出情况
            wb.close()
            self.GuiRefresh(self.Status, 'Process Finished')
            self.Process.setEnabled(True)
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])
        # 错误情况放在error text里，方便排查
        except Exception as ex:
            self.GuiRefresh(self.ErrorText, str(ex))
            self.Process.setEnabled(True)
            try:
                wb.close()
            except Exception as e:
                pass

    # """对QDialog类重写，实现一些功能"""
    # """重写closeEvent方法，实现dialog窗体关闭时执行一些代码
    def closeEvent(self, event):
        # :param event: close()触发的事件
        # :return: None
        # """
        reply = QtWidgets.QMessageBox.question(self,
                                               'Exit',
                                               "Confirm Exit?",
                                               QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                                               QtWidgets.QMessageBox.StandardButton.No)
        if reply == QtWidgets.QMessageBox.StandardButton.Yes:
            #         self.stoptime=datetime.datetime.now()
            event.accept()
        else:
            event.ignore()


if __name__ == "__main__":
    # QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    form = GUI_Dialog()
    form.show()

    app.exec()

    form.xwapp.quit()
    try:
        form.xwapp.kill()
    except Exception as e:
        pass
