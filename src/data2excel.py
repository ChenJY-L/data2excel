"""
数据处理工具 - 将DAT文件转换为Excel格式并生成图表

主要功能：
1. 读取多环光谱数据文件(.dat格式)
2. 处理温度数据和实验信息
3. 计算单环和差分吸光度数据
4. 生成Excel报告和图表
5. 支持动态基准周期和血糖数据处理

作者：[作者信息]
版本：[版本信息]
最后更新：[更新日期]
"""

import re
import xlwings.utils
from dateutil import parser
from pathlib import Path
from fnmatch import translate
import sys
# import killpywintypes
import xlwings as xw  # Excel操作模块
import datetime  # 时间模块
# import time
import os
import numpy as np
from PySide6 import QtWidgets
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *
from PySide6.QtGui import QIcon
import gui as QTUI  # GUI界面模块
import ico01
import ctypes

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myappid")


class GUI_Dialog(QWidget, QTUI.Ui_Data_Processing):
    """
    数据处理主界面类

    继承自QDialog和GUI界面类，提供数据处理的图形用户界面
    主要功能包括文件选择、数据处理、Excel生成和图表绘制
    """

    # ==================== 类配置变量 ====================

    # 图表绘制配置
    CHART_LEFT = 400        # 图表左距
    CHART_TOP = 400         # 图表上距
    CHART_WIDTH = 700       # 图表宽度
    CHART_HEIGHT = 400      # 图表高度

    # 字体大小配置
    TITLE_FONT_SIZE = 28    # 标题字体大小
    AXIS_FONT_SIZE = 18     # 轴字体大小
    AXIS_TITLE_FONT_SIZE = 18  # 轴标题字体大小
    LEGEND_FONT_SIZE = 18   # 图例字体大小
    ANNOTATION_FONT_SIZE = 40  # 标注字体大小
    MAIN_TITLE_FONT_SIZE = 54  # 主标题字体大小
    LABEL_FONT_SIZE = 28    # 标签字体大小

    # 图表样式配置
    LINE_WEIGHT = 1.5              # 线宽
    CHART_BORDER_COLOR = 14277081  # 图表边框颜色
    GRID_COLOR = 14277081          # 网格线颜色
    GRID_WEIGHT = 0.75             # 网格线粗细
    AXIS_WEIGHT = 1.5              # 坐标轴线粗细
    ANNOTATION_LINE_WEIGHT = 2.5   # 标注线粗细

    # 颜色配置
    CHART_COLORS = ['#60966d','#5b3660','#018abe','#e90f44','#63adee','#924c63','#7c8fa3']
    TITLE_BOX_COLOR = (255, 255, 0)  # 标题框背景色（黄色）

    # 阈值
    ERROR_THRESH = 5e-6

    TempCorr1050 = 0.0005   # 1050差分温度矫正系数
    """
    用于设置复制放大的chart的id和需要忽略的chart
    """
    Duplicate_Target = {
        "index": [2, 5, 8, 11],
        "ignore_series": [[], [2,3,4], [], []]
    }

    def __init__(self, parent=None):
        """
        初始化GUI界面

        Args:
            parent: 父窗口对象，默认为None
        """
        super(GUI_Dialog, self).__init__(parent)
        self.setupUi(self)

        # 设置窗口标题和图标
        current_file = os.path.basename(__file__)
        current_file = os.path.splitext(current_file)[0][-8:]
        self.setWindowTitle('Data_Processing ' + current_file)
        self.setWindowIcon(QIcon(':/favicon01.ico'))

        # 设置默认选项
        # self.DyBCCheckBox.setChecked(True)
        self.OGTTCheckBox.setChecked(True)
        self.waveDiffCheckBox.setChecked(True)
        self.expInfoCheckBox.setChecked(True)
        self.tempCorrelationCheckBox.setChecked(True)
        self.duplicateCheckBox.setChecked(True)

        # 连接信号和槽函数
        self.Process.clicked.connect(self.DataProcess)
        self.FileSelect.clicked.connect(self.FileSelectF)

        # 初始化Excel应用程序
        self.xwapp = xw.App(visible=False, add_book=False)
        self.xwapp.display_alerts = False
        self.xwapp.screen_updating = False

        # 初始化图表列表
        self.charts = []

        self.isInfo = False

    # ==================== 工具函数区 ====================

    def GuiRefresh(self, textbox, text):
        """
        刷新GUI界面文本框内容

        Args:
            textbox: 要更新的文本框控件
            text: 要显示的文本内容
        """
        textbox.setPlainText(text)
        QApplication.processEvents()

    def hexColor2Int(self, color):
        """
        将十六进制颜色值转换为整数

        Args:
            color: 十六进制颜色字符串，如'#FF0000'

        Returns:
            int: 对应的整数颜色值
        """
        return xw.utils.rgb_to_int(xw.utils.hex_to_rgb(color))

    def FileSelectF(self):
        """
        文件选择对话框

        打开文件选择对话框，允许用户选择DAT数据文件
        """
        self.GuiRefresh(self.Status, 'Selecting Path')
        SchPath = os.getcwd()
        if self.Path.toPlainText() != '':
            SchPath = os.path.dirname(self.Path.toPlainText())
        filenames = (QFileDialog.getOpenFileNames(self, 'Select', SchPath, "Data Files(*.dat);;All Files(*)"))
        if filenames[0] != []:
            self.Path.setPlainText(filenames[0][0])

    def FilePath(self, path, filename):
        """
        根据关键词在指定路径中查找文件

        Args:
            path: 搜索路径
            filename: 文件名关键词

        Returns:
            str: 找到的文件完整路径，如果未找到则返回空路径
        """
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

    def CheckSheet(self, workbook, sheetname):
        """
        检查Excel工作簿中是否存在指定名称的工作表

        Args:
            workbook: Excel工作簿对象
            sheetname: 要检查的工作表名称

        Returns:
            bool: 如果工作表存在返回True，否则返回False
        """
        num = len(workbook.sheets)  # 获取工作表个数
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

    def FindRowColRange(self, SheetName, Rttype, KeyWord, RangeStr):
        """
        在Excel工作表中查找关键词并返回单元格信息

        Args:
            SheetName: Excel工作表对象
            Rttype: 返回类型 ('Adr'=地址, 'Row'=行号, 'Col'=列号, 'Ran'=行列号)
            KeyWord: 要查找的关键词
            RangeStr: 搜索范围字符串

        Returns:
            str: 根据Rttype返回相应的单元格信息
        """
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

    def parseText(self, text):
        TIME_RE = re.compile(r'^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*$')
        RANGE_RE = re.compile(r'^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*-\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*$')

        def excel_time(h, m, s=0):
            h, m, s = int(h), int(m), int(s or 0)
            if not (0 <= h < 24 and 0 <= m < 60 and 0 <= s < 60):
                raise ValueError
            return (h * 3600 + m * 60 + s) / 86400

        def parse_time(s):
            m = RANGE_RE.match(s)
            if m:
                return (excel_time(m.group(1), m.group(2), m.group(3)),
                        excel_time(m.group(4), m.group(5), m.group(6)))
            m = TIME_RE.match(s)
            if m:
                return excel_time(m.group(1), m.group(2), m.group(3))
            raise ValueError

        def to_num(s):
            s = s.strip()
            try:
                return int(s) if re.fullmatch(r'[+-]?\d+', s) else float(s)
            except Exception:
                return s

        def header_kind(line):
            # 兼容：血糖: / 基准周期: / [血糖] / [基准周期]
            t = line.strip()
            if t.startswith('[') and t.endswith(']'):
                t = t[1:-1].strip()
            if t.endswith(':'):
                t = t[:-1].strip()
            if t == '血糖':
                return 'blood_glucose'
            if t == '基准周期':
                return 'baseline_cycle'
            return None

        def looks_structured(line):
            # 用于结束备注块：冒号左右任一侧像时间/时间段，就认为是结构化
            if ':' not in line:
                return False
            a, b = line.split(':', 1)
            a, b = a.strip(), b.strip()
            return any([TIME_RE.match(a), RANGE_RE.match(a), TIME_RE.match(b), RANGE_RE.match(b)])

        schedule, blood_glucose, baseline_cycle, remark = [], [], [], []
        mode = None  # None / 'blood_glucose' / 'baseline_cycle'

        lines = text.splitlines()
        i = 0
        while i < len(lines):
            line = lines[i].strip().replace('：', ':')
            i += 1
            if not line or line.startswith('#'):
                continue

            # --- 备注：行内 or 备注块 ---
            if line.startswith('备注'):
                tail = line[2:].lstrip(':').strip()
                if tail:
                    remark.append(tail)
                    continue

                # 备注块：收集后续非结构化行；遇到段头/结构化行就停止（不吞掉那行）
                while i < len(lines):
                    s = lines[i].strip().replace('：', ':')
                    if not s or s.startswith('#'):
                        i += 1
                        continue
                    if header_kind(s) or looks_structured(s):
                        break
                    remark.append(s)
                    i += 1
                continue

            # --- 段头 ---
            k = header_kind(line)
            if k:
                mode = k
                continue

            # --- 段内：时间(段): 值 ---
            if mode in ('blood_glucose', 'baseline_cycle'):
                if ':' not in line:
                    (blood_glucose if mode == 'blood_glucose' else baseline_cycle).append(
                        {"time": None, "value": line}
                    )
                    continue

                left, right = line.rsplit(':', 1)
                try:
                    t = parse_time(left.strip())
                except Exception:
                    t = None
                (blood_glucose if mode == 'blood_glucose' else baseline_cycle).append(
                    {"time": t, "value": to_num(right)}
                )
                continue

            # --- 其他：走 schedule 规则 ---
            if ':' not in line:
                schedule.append({"time": None, "activity": line})
                continue

            if re.match(r'^\d', line):  # 数字开头：时间: 活动
                time_str, activity = line.rsplit(':', 1)
            else:  # 非数字开头：活动: 时间
                activity, time_str = line.split(':', 1)

            try:
                t = parse_time(time_str.strip())
                schedule.append({"time": t, "activity": activity.strip()})
            except Exception:
                schedule.append({"time": None, "activity": line})

        return {
            "schedule": schedule,
            "baseline_cycle": baseline_cycle or None,
            "blood_glucose": blood_glucose or None,
            "remark": "\n".join(remark).strip() if remark else None
        }

    # ==================== 数据处理子函数区 ====================

    def load_and_validate_data(self):
        """
        加载和验证数据文件

        Returns:
            tuple: (Chvalues, Ch, C, Chpath)
                - Chvalues: 各环数据数组
                - Ch: 环数
                - C: 是否为中文文件名
                - Chpath: 数据文件路径
        """
        # 确定文件路径
        if self.Path.toPlainText() == '':
            self.GuiRefresh(self.Status, 'Obtaining File Path')
            path = os.getcwd()  # 当前文件位置
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

        # 获取各环数据
        self.GuiRefresh(self.Status, 'Loading Data')
        Chpath = self.Path.toPlainText()
        Chpath = Chpath.replace('/', '\\')

        # 判断文件名类型（中文或英文）
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
        return Chvalues, Ch, C, Chpath

    def preprocess_chvalues(self, raw_chvalues_list):
        """
        对从xlwings读取的原始Chvalues列表进行预处理。
        1. 将Python的None值统一替换为NumPy的np.nan。
        2. 将数据结构转换为一个单一的、类型为float的3D NumPy数组。

        Args:
            raw_chvalues_list (list): 包含多个二维列表的原始数据。

        Returns:
            np.ndarray: 一个(Ch, Rows, Cols)形状的、可用于计算的3D NumPy数组。
        """
        self.GuiRefresh(self.Status, 'Preprocessing raw data...')
        try:
            array_3d = np.asarray(raw_chvalues_list, dtype=object)
            array_3d[array_3d == None] = np.nan
            numeric_array = array_3d.astype(float)
            if not self.LDCheckBox.isChecked():
                numeric_array[numeric_array < self.ERROR_THRESH] = np.nan
            else:
                rows_to_remove = numeric_array[1, :, -1] == 0
                # 使用反向掩码（~）来保留所有不包含NaN的行
                numeric_array = numeric_array[:, ~rows_to_remove, :]
            # 将整个数组的类型转换为float，以便进行数学运算
            return numeric_array

        except Exception as e:
            # 如果数据形状不规则，上面的方法会失败，这里提供一个备用方案
            self.GuiRefresh(self.ErrorText, f'Data shape irregular, using fallback. Error: {e}')
            processed_list = []
            for ch_data in raw_chvalues_list:
                arr = np.array(ch_data, dtype=object)
                arr[arr == None] = np.nan
                processed_list.append(arr.astype(float))
            # 注意：如果各环行数不一致，这里可能需要更复杂的处理
            return np.asarray(processed_list)

    def process_temperature_data(self, Chpath, C):
        """
        处理温度数据

        Args:
            Chpath: 数据文件路径
            C: 是否为中文文件名

        Returns:
            tuple: (Tempvalue, Temptitle) 温度数据和标题
        """
        Tempvalue = []
        Temptitle = []

        if self.TempCheckBox.isChecked():
            self.GuiRefresh(self.Status, 'Loading Temperature Data')
            temppath = os.path.split(Chpath)[0]
            tempfilekwords = '温度' if C else 'Temperature'
            wbt = self.xwapp.books.open(self.FilePath(temppath, tempfilekwords))
            tempsheet = wbt.sheets[0]
            Tempvalue = tempsheet.range(tempsheet.used_range).value
            Temptitle = Tempvalue[0][:]
            Tempvalue = Tempvalue[1:][:]
            Temptitle = np.asarray(Temptitle)
            Tempvalue = np.asarray(Tempvalue)
            wbt.close()

        return Tempvalue, Temptitle

    def parse_experiment_info(self, Chpath):
        """
        解析实验信息文件

        Args:
            Chpath: 数据文件路径

        Returns:
            dict: 实验信息字典，如果没有找到文件则返回None
        """
        filePath = os.path.dirname(Chpath)
        txtFiles = list(Path(filePath).glob('*.txt'))  # 获取所有 .txt 文件列表

        if txtFiles:
            # 读取第一个txt文件
            text = Path(txtFiles[0]).read_text(encoding='utf-8')
            expInfo = self.parseText(text)

            self.isInfo = True
            return expInfo
        else:
            # 没有txt文件，跳过
            self.isInfo = False
            return None

    def get_baseline_for_cycle(self, cycle_index, timearr, expInfo):
        """
        根据周期索引获取对应的基准周期索引

        Args:
            cycle_index: 当前周期索引（0-based）
            timearr: 时间数组
            expInfo: 实验信息（包含 baseline_cycle 配置）

        Returns:
            int: 基准周期索引（1-based，与GUI中BaseCycle一致）
        """
        # 如果没有 expInfo 或没有 baseline_cycle 配置，使用默认的 BaseCycle
        if expInfo is None or expInfo.get('baseline_cycle') is None:
            return self.BaseCycle.value()

        baseline_cycles = expInfo['baseline_cycle']
        if not baseline_cycles:
            return self.BaseCycle.value()

        # 获取当前周期对应的时间（确保获取标量值）
        if cycle_index < len(timearr):
            current_time = float(np.ravel(timearr[cycle_index])[0])
        else:
            current_time = float(np.ravel(timearr[-1])[0])
        # 将时间转换为当天的小时数部分（仅保留小数部分）
        time_of_day = current_time % 1.0

        # 遍历 baseline_cycle 配置，查找匹配的时间范围
        for item in baseline_cycles:
            time_range = item.get('time')
            value = item.get('value')

            if time_range is None or value is None:
                continue

            # 处理时间范围
            if isinstance(time_range, tuple) and len(time_range) == 2:
                start_time, end_time = time_range
                if start_time <= time_of_day < end_time:
                    return int(value)
            elif isinstance(time_range, (int, float)):
                # 单个时间点：该时间之后使用此基准周期
                if time_of_day >= time_range:
                    return int(value)

        # 如果没有匹配的配置，返回默认值
        return self.BaseCycle.value()

    def calculate_base_data(self, Chvalues, Ch, wn, m, expInfo=None):
        """
        计算基准周期的单环和差分数据
        支持多基准周期：预计算所有可能的基准周期数据

        Args:
            Chvalues: 各环数据数组
            Ch: 环数
            wn: 波长数量
            m: 每次测量数加一
            expInfo: 实验信息（包含 baseline_cycle 配置）

        Returns:
            tuple: (basesingle_dict, basediff_dict) 基准数据字典，键为基准周期索引
        """
        self.GuiRefresh(self.Status, 'Calculating Base Data')

        # 收集所有需要计算的基准周期索引
        base_cycles = {self.BaseCycle.value()}  # 默认基准周期

        if expInfo is not None and expInfo.get('baseline_cycle'):
            for item in expInfo['baseline_cycle']:
                value = item.get('value')
                if value is not None:
                    base_cycles.add(int(value))

        # 为每个基准周期计算基准数据
        basesingle_dict = {}
        basediff_dict = {}
        diffNo = int(Ch * (Ch - 1) // 2)

        for base_cycle in base_cycles:
            basesingle = np.zeros((Ch, wn))
            basediff = np.zeros((diffNo, wn))

            for w in range(0, wn):
                cs = -1
                for r in range(0, Ch):
                    basesingle[r][w] = sum(Chvalues[r][(base_cycle - 1) * 6 + w][:m - 1]) / (m - 1)
                    for rl in range(r + 1, Ch):
                        cs = cs + 1
                        basediff[cs][w] = sum(np.log(Chvalues[r][(base_cycle - 1) * 6 + w][:m - 1] / Chvalues[rl][(base_cycle - 1) * 6 + w][:m - 1])) / (m - 1)

            basesingle_dict[base_cycle] = basesingle
            basediff_dict[base_cycle] = basediff

        return basesingle_dict, basediff_dict

    def create_excel_workbook(self, Chpath, C):
        """
        创建Excel工作簿和工作表

        Args:
            Chpath: 数据文件路径
            C: 是否为中文文件名

        Returns:
            xlwings.Book: Excel工作簿对象
        """
        self.GuiRefresh(self.Status, 'Creating Output File')
        ProcessFilePath = os.path.join(Chpath.replace(Chpath.split('\\')[-1], ''),
                                       Chpath.split('\\')[-2] + '.xlsx') if C else (
                                       Chpath.split('Ch')[0] + 'Processed' + '.xlsx')

        if os.path.isfile(ProcessFilePath):
            self.GuiRefresh(self.Status, 'Removing Existing Output File')
            try:
                os.remove(ProcessFilePath)
            except OSError as exc:
                self.GuiRefresh(self.Status, 'Failed to Remove Existing File')
                QMessageBox.critical(
                    self,
                    'File In Use',
                    f'Unable to remove existing Excel file:\n{ProcessFilePath}\nPlease close the file and try again.\n\nDetails: {exc}'
                )
                raise

        wb = self.xwapp.books.add()  # 在app下创建一个Book
        wb.save(ProcessFilePath)
        wb = self.xwapp.books.open(ProcessFilePath)

        # 创建工作表
        sheetnames = ['单环', '单环吸光度', '单环信噪比', '差分', '差分吸光度', '差分等效信噪比',
                      '光强和信噪比汇总', '温度数据'] if C else ['Single Ring', 'Single Ring Absorbance',
                                                                     'Single SNR', 'Differential',
                                                                     'Differential Absorbance', 'Differential SNR',
                                                                     'Summary of Intensity and SNR', 'Temperature']
        self.GuiRefresh(self.Status, 'Creating Sheets')
        if self.TempCheckBox.isChecked() or self.isInfo: # 存在备注或选中温度选项时，创建`温度数据` Sheet
            sheetNo = len(sheetnames)
        else:
            sheetNo = len(sheetnames) - 1

        for i in range(0, sheetNo):
            if self.CheckSheet(wb, sheetnames[i]) == False:
                wb.sheets.add(sheetnames[i], after=i + 1)
        wb.save()

        return wb, sheetnames

    def write_time_and_headers(self, wb, sheetnames, timearr, cycleNoarr):
        """
        在所有工作表中写入时间和序号标题

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            timearr: 时间数组
            cycleNoarr: 周期序号数组
        """
        self.GuiRefresh(self.Status, 'Writing Title')
        for s in range(0, 6):
            wb.sheets[sheetnames[s]].range(3, 1).value = timearr
            wb.sheets[sheetnames[s]].range(3, 2).value = cycleNoarr

    def process_temperature_interpolation(self, wb, sheetnames, timearr, cycleNoarr, Tempvalue, Temptitle):
        """
        处理温度数据插值并写入Excel

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            timearr: 时间数组
            cycleNoarr: 周期序号数组
            Tempvalue: 温度数据
            Temptitle: 温度数据标题

        Returns:
            numpy.array: 插值后的温度数据
        """
        try:
            yinterp = np.empty((timearr.shape[0], Tempvalue.shape[1] - 1))
        except:
            yinterp = np.empty((timearr.shape[0], 1))

        if self.TempCheckBox.isChecked():
            Tempvalue = np.asarray(Tempvalue, dtype=float) # fix data type
            wb.sheets[sheetnames[len(sheetnames) - 1]].range(2, 1).value = timearr
            wb.sheets[sheetnames[len(sheetnames) - 1]].range(2, 2).value = cycleNoarr
            wb.sheets[sheetnames[len(sheetnames) - 1]].range(1, 2).value = Temptitle
            for i in range(1, Tempvalue.shape[1]):
                yinterp[:, i - 1] = np.interp(timearr, Tempvalue[:, 0], Tempvalue[:, i]).reshape(timearr.shape[0], )
            wb.sheets[sheetnames[len(sheetnames) - 1]].range(2, 3).value = yinterp
            self.GuiRefresh(self.Status, 'Writing Temp Data')
        else:
            wb.sheets[sheetnames[len(sheetnames) - 1]].range(2, 1).value = timearr
            wb.sheets[sheetnames[len(sheetnames) - 1]].range(2, 2).value = cycleNoarr
        wb.save()

        return yinterp

    def write_summary_data(self, wb, sheetnames, wave, ringwords, diffwords, Ch, C):
        """
        写入光强和信噪比数据汇总的标题

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            wave: 波长列表
            ringwords: 环标签列表
            diffwords: 差分标签列表
            Ch: 环数
            C: 是否为中文文件名
        """
        self.GuiRefresh(self.Status, 'Writing Summary Data')
        wb.sheets[sheetnames[6]].range(3, 3).value = '光强' if C else 'Intensity'
        wb.sheets[sheetnames[6]].range(3, 13).value = '信噪比' if C else 'SNR'
        wb.sheets[sheetnames[6]].range(4, 5).value = wave
        wb.sheets[sheetnames[6]].range(4, 15).value = wave
        wb.sheets[sheetnames[6]].range(5, 4).options(transpose=True).value = ringwords
        wb.sheets[sheetnames[6]].range(5, 14).options(transpose=True).value = ringwords
        wb.sheets[sheetnames[6]].range(5 + Ch, 4).options(transpose=True).value = diffwords
        wb.sheets[sheetnames[6]].range(5 + Ch, 14).options(transpose=True).value = diffwords
        wb.save()

    def process_single_ring_data(self, wb, sheetnames, Chvalues, Ch, ringwords, wave, datarange, basesingle_dict, n, wn, m, timearr=None, expInfo=None):
        """
        处理单环数据并写入Excel
        支持多基准周期数据处理

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            Chvalues: 各环数据数组
            Ch: 环数
            ringwords: 环标签列表
            wave: 波长列表
            datarange: 数据范围
            basesingle_dict: 基准单环数据字典，键为基准周期索引
            n: 数据行数
            wn: 波长数量
            m: 每次测量数加一
            timearr: 时间数组（用于多基准周期判断）
            expInfo: 实验信息（用于多基准周期判断）
        """
        for r in range(0, Ch):
            # 写入数据头
            for s in range(0, 3):
                wb.sheets[sheetnames[s]].range(1, 3 + 7 * r).value = ringwords[r]
                wb.sheets[sheetnames[s]].range(2, 4 + 7 * r).value = wave

            # 计算各列数据
            self.GuiRefresh(self.Status, 'Writing Ring ' + str(r + 1))
            singlearr = []
            singleabsarr = []
            singlesnrarr = []

            for j in datarange:  # len(Chvalues[0])或者n

                # 利用时间戳数值远远大于数据的特点，判断时间的索引，提取时间索引前的全部数据
                raw_data = Chvalues[r][j]
                if np.isnan(raw_data).all():
                    raw_data = np.ones(np.size(raw_data))
                time_index = np.nanargmax(raw_data)
                singles = raw_data[:time_index]
                single = np.mean(singles)

                # 获取当前周期对应的基准周期
                cycle_index = j // wn
                base_cycle = self.get_baseline_for_cycle(cycle_index, timearr, expInfo)
                basesingle = basesingle_dict[base_cycle]

                singleabs = np.log(basesingle[r][j % wn] / single)
                singlesnr = single / np.std(singles, ddof=1)
                singlearr.append(single)
                singleabsarr.append(singleabs)
                singlesnrarr.append(singlesnr)

            singlearr = np.array([singlearr]).reshape(n // wn, wn)
            singleabsarr = np.array([singleabsarr]).reshape(n // wn, wn)
            singlesnrarr = np.array([singlesnrarr]).reshape(n // wn, wn)
            singleave = singlearr.mean(axis=0)
            singlesnrave = singlesnrarr.mean(axis=0)

            # 写入数据
            wb.sheets[sheetnames[0]].range(3, 4 + 7 * r).value = singlearr
            wb.sheets[sheetnames[1]].range(3, 4 + 7 * r).value = singleabsarr
            wb.sheets[sheetnames[2]].range(3, 4 + 7 * r).value = singlesnrarr
            wb.sheets[sheetnames[6]].range(5 + r, 5).value = singleave
            wb.sheets[sheetnames[6]].range(5 + r, 15).value = singlesnrave

            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

        wb.save()
        self.GuiRefresh(self.Status, 'Writing Single Finished')

    def process_differential_data(self, wb, sheetnames, Chvalues, Ch, diffwords, wave, datarange, basediff_dict, n, wn, m, timearr=None, expInfo=None):
        """
        处理差分数据并写入Excel
        支持多基准周期数据处理

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            Chvalues: 各环数据数组
            Ch: 环数
            diffwords: 差分标签列表
            wave: 波长列表
            datarange: 数据范围
            basediff_dict: 基准差分数据字典，键为基准周期索引
            n: 数据行数
            wn: 波长数量
            m: 每次测量数加一
            timearr: 时间数组（用于多基准周期判断）
            expInfo: 实验信息（用于多基准周期判断）
        """
        cs = -1
        for r in range(0, Ch):
            for rl in range(r + 1, Ch):
                # 写入数据头
                cs = cs + 1
                for s in range(3, 6):
                    wb.sheets[sheetnames[s]].range(1, 3 + 7 * cs).value = diffwords[cs]
                    wb.sheets[sheetnames[s]].range(2, 4 + 7 * cs).value = wave

                # 计算各列数据
                self.GuiRefresh(self.Status, 'Writing Diff ' + str(r + 1) + str(rl + 1))
                diffarr = np.empty((n, 1))
                diffabsarr = np.empty((n, 1))
                diffsnrarr = np.empty((n, 1))

                for j in datarange:  # len(Chvalues[0])或者n

                    # ASKME: Same
                    # diffs = np.log(Chvalues[r][j][1:m - 1] / Chvalues[rl][j][1:m - 1])
                    # diff = sum(diffs) / (m - 2)
                    raw_data = Chvalues[r][j]
                    if np.isnan(raw_data).all():
                        raw_data = np.ones(np.size(raw_data))
                    time_index = np.nanargmax(raw_data)

                    diffs = np.log(Chvalues[r][j][:time_index] / Chvalues[rl][j][:time_index])
                    diff = np.mean(diffs)

                    # 获取当前周期对应的基准周期
                    cycle_index = j // wn
                    base_cycle = self.get_baseline_for_cycle(cycle_index, timearr, expInfo)
                    basediff = basediff_dict[base_cycle]

                    diffabs = diff - basediff[cs][j % wn]
                    diffsnr = 1 / np.std(diffs, ddof=1)
                    diffarr[j] = diff
                    diffabsarr[j] = diffabs
                    diffsnrarr[j] = diffsnr

                diffarr = diffarr.reshape(n // wn, wn)
                diffabsarr = np.array([diffabsarr]).reshape(n // wn, wn)
                diffsnrarr = diffsnrarr.reshape(n // wn, wn)
                diffarrave = diffarr.mean(axis=0)
                diffsnrave = diffsnrarr.mean(axis=0)

                # 写入数据
                wb.sheets[sheetnames[3]].range(3, 4 + 7 * cs).value = diffarr
                wb.sheets[sheetnames[4]].range(3, 4 + 7 * cs).value = diffabsarr
                wb.sheets[sheetnames[5]].range(3, 4 + 7 * cs).value = diffsnrarr
                wb.sheets[sheetnames[6]].range(5 + Ch + cs, 5).value = diffarrave
                wb.sheets[sheetnames[6]].range(5 + Ch + cs, 15).value = diffsnrave

                self.currenttime = datetime.datetime.now()
                self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

        wb.save()
        self.GuiRefresh(self.Status, 'Writing Diff Finished')

    def handle_dynamic_base_cycle(self, wb, sheetnames, Ch, n, wn, C):
        """
        处理动态基准周期功能

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            Ch: 环数
            n: 数据行数
            wn: 波长数量
            C: 是否为中文文件名
        """
        if self.DyBCCheckBox.isChecked():
            wb.sheets[sheetnames[1]].range(3, 3).value = '单环基准周期' if C else 'Single Base Cycle'
            wb.sheets[sheetnames[1]].range(4, 3).value = self.BaseCycle.value()
            wb.sheets[sheetnames[4]].range(3, 3).value = '差分基准周期' if C else 'Single Base Cycle'
            wb.sheets[sheetnames[4]].range(4, 3).value = self.BaseCycle.value()

            # 生成动态公式
            ss = '=LN(@INDIRECT("' + wb.sheets[
                sheetnames[0]].name + '!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&($C$4+2))/@INDIRECT("' + wb.sheets[
                     sheetnames[0]].name + '!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&ROW()))'
            singleabsarr = np.full((n // wn, wn), ss)
            ds = '=@INDIRECT("' + wb.sheets[
                sheetnames[3]].name + '!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&ROW())-@INDIRECT("' + wb.sheets[
                     sheetnames[3]].name + '!"&SUBSTITUTE(ADDRESS(1,COLUMN(),4),1,"")&($C$4+2))'
            diffabsarr = np.full((n // wn, wn), ds)

            # 写入动态公式
            for r in range(0, Ch):
                self.GuiRefresh(self.Status, 'Writing Dyna Ring ' + str(r + 1))
                wb.sheets[sheetnames[1]].range(3, 4 + 7 * r).value = singleabsarr
            for rl in range(0, int(Ch * (Ch - 1) / 2)):
                self.GuiRefresh(self.Status, 'Writing Dyna Diff ' + str(rl + 1))
                wb.sheets[sheetnames[4]].range(3, 4 + 7 * rl).value = diffabsarr
                self.currenttime = datetime.datetime.now()
                self.GuiRefresh(self.ErrorText,
                                'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])
            self.GuiRefresh(self.Status, 'Saving...')
            wb.save()

    def add_glucose_data(self, wb, sheetnames, timearr, yinterp, C, expInfo=None):
        """
        添加血糖数据到Excel

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            timearr: 时间数组
            yinterp: 插值后的温度数据
            C: 是否为中文文件名
            expInfo: 实验信息（包含 parseText 解析出的 blood_glucose）

        Returns:
            int: 血糖数据列的索引
        """

        try:
            rng_lcol = len(yinterp[1]) + 3
        except:
            print("No temperature data")
            rng_lcol = 2

        # 如果明确是LD数据则不添加血糖值，但仍返回列索引用于后续图表逻辑
        if self.LDCheckBox.isChecked():
            return rng_lcol

        try:
            temp_sheet = wb.sheets[sheetnames[7]]
        except Exception:
            return rng_lcol


        time_flat = np.asarray(timearr, dtype=float).reshape(-1)
        base_day = float(np.floor(time_flat[0])) if len(time_flat) else 0.0
        bg_list = (expInfo or {}).get("blood_glucose") if isinstance(expInfo, dict) else None

        records = []
        if isinstance(bg_list, list):
            for item in bg_list:
                if not isinstance(item, dict):
                    continue
                value = float(item.get("value"))
                if value is None:
                    continue

                t = item.get("time")
                if t is None:
                    records.append((None, value))
                else:
                    target = float(t)
                    if target is not None:
                        records.append((target + base_day, value))

        # 兼容旧行为：没有提供血糖数据时，默认在首/中/末三个点放置5.5
        if not records and len(time_flat) >= 1:
            records = [
                (float(time_flat[0]), 5.5),
                (float(time_flat[len(time_flat) // 2]), 5.5),
                (float(time_flat[-1]), 5.5),
            ]

        # 设置时间格式
        gcols = xw.utils.col_name(rng_lcol + 1) + ':' + xw.utils.col_name(rng_lcol + 1)
        temp_sheet.api.Columns(gcols).NumberFormatLocal = "[$-x-systime]h:mm:ss AM/PM"

        # 写入血糖数据（仅输出这几行）
        temp_sheet.range(1, rng_lcol + 2).value = '血糖值' if C else 'Glucose Value'
        temp_sheet.range(2, rng_lcol + 1).options(transpose=True).value = [t for t, _ in records] or None
        temp_sheet.range(2, rng_lcol + 2).options(transpose=True).value = [v for _, v in records] or None

        return rng_lcol

    def add_info_data(self, wb, sheetnames, expInfo, timearr):
        """
        添加实验信息到Excel

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            expInfo: 实验信息
            timearr: 时间数组
        """
        if expInfo is None:
            return

        self.GuiRefresh(self.Status, 'Writing relatice height...')

        # 定义正则表达式，检测activity是否为纯数字
        pattern = r'^-?\d+(\.\d+)?$'

        max_col = wb.sheets[sheetnames[7]].used_range.last_cell.column
        wb.sheets[sheetnames[7]].range(1, max_col + 1).value = '相对高度(cm)'

        relative_height = np.full((timearr.shape[0], 1), np.nan)
        for nitem, item in enumerate(expInfo['schedule']):
            time_data = item.get('time')
            activity = item.get('activity')
            height = float(activity) if re.fullmatch(pattern, activity) else np.nan

            if time_data is None or np.isnan(height):
                continue

            # 确保时间数据是可迭代的
            if isinstance(time_data, (int, float)):
                # 跳过时间点
                continue
            elif isinstance(time_data, (list, tuple)) and len(time_data) == 2:
                t = [time_data[0] + np.floor(timearr[0]), time_data[1] + np.floor(timearr[0])]
            else:
                continue

            relative_height[np.logical_and(timearr >= t[0], timearr < t[1])] = height

        # relative_height = np.where(np.isnan(relative_height), None, relative_height)
        wb.sheets[sheetnames[7]].range(2, max_col + 1).value = relative_height

    def add_gradient_data(self, wb, sheetnames, expInfo, timearr, C, diffwords, wave):
        """
        处理和添加详细实验信息。

        该函数主要执行三大任务：
        1. 根据实验信息中标记为数值的特定时间段，计算差分吸收光谱的变化率（梯度/斜率）。
        2. 将计算出的梯度数据，连同时间和序号，一同写入一个名为“梯度分析”的新工作表中，并采用特定格式展示。

        Args:
            wb (xw.Book): xlwings 的工作簿对象。
            sheetnames (list): 所有工作表的标准名称列表。
            expInfo (dict): 解析后的实验信息，包含时间安排和备注。
            timearr (np.array): 包含所有时间点的一维Numpy数组。
            C (bool): 中文模式标志，True表示使用中文命名。
            diffwords (list): 差分对的名称列表 (例如, ['Diff12', 'Diff23'])。
            wave (list): 波长的名称列表 (例如, ['1050', '1219'])。
        """
        # 如果没有实验信息，则直接返回
        if expInfo is None or not self.gradientCheckBox.isChecked():
            return

        # --- 1. 初始化和准备 ---

        # 定义工作表名称
        grad_sheet_name = '时间-斜率' if C else 'Time-Slope'
        slope_sheet_name = '高度-斜率' if C else 'Height-Slope'

        self.GuiRefresh(self.Status, 'Creating gradient sheets')
        # 确保“梯度分析”工作表存在，如果不存在则创建
        if self.gradientCheckBox.isChecked():
            wb.sheets.add(grad_sheet_name)
            wb.sheets.add(slope_sheet_name)

        # 获取所需的工作表对象
        gradient_sheet = wb.sheets[grad_sheet_name]
        slope_sheet = wb.sheets[slope_sheet_name]
        diff_abs_sheet = wb.sheets[sheetnames[4]]

        self.GuiRefresh(self.Status, 'Writing headers')
        # 清理并格式化“梯度分析”表
        gradient_sheet.clear()
        gradient_sheet.range('A2').value = ['时间', '周期'] if C else ['Time', 'Cycle']
        slope_sheet.range('A2').value = '相对高度(cm)'

        # 写入时间和周期数据，并设置时间列的显示格式
        cycleNoarr = np.arange(1, len(timearr) + 1).reshape(-1, 1)
        gradient_sheet.range('A3').value = timearr
        gradient_sheet.range('B3').value = cycleNoarr
        gradient_sheet.api.Columns("A:A").NumberFormatLocal = "[$-x-systime]h:mm:ss AM/PM"

        # 创建带分隔空列的双行表头
        header_offset = 2  # A, B列已被时间和序号占用
        for cs, diff_pair in enumerate(diffwords):
            # 计算每个差分数据块前分隔列的列号
            separator_col = header_offset + 1 + cs * (len(wave) + 1)
            # 在分隔列的第一行写入差分对名称
            gradient_sheet.range(1, separator_col).value = diff_pair
            # 在数据列的第二行写入对应的波长名称
            gradient_sheet.range(2, separator_col + 1).value = wave

            # 在分隔列的第一行写入差分对名称
            slope_sheet.range(1, separator_col - header_offset + 1).value = diff_pair
            # 在数据列的第二行写入对应的波长名称
            slope_sheet.range(2, separator_col - header_offset + 2).value = wave

        # --- 2. 加载光谱数据并准备计算矩阵 ---
        self.GuiRefresh(self.Status, 'Reading diff abs data')
        # 从“差分吸光度”表中一次性读取所有相关数据，提高效率
        try:
            last_data_col = 4 + 7 * (len(diffwords) - 1) + (len(wave) - 1)
            all_diff_abs_data = diff_abs_sheet.range((3, 4), (2 + len(timearr), last_data_col)).options(np.array,
                                                                                                        ndim=2).value
        except Exception as e:
            self.GuiRefresh(self.ErrorText, f'读取差分吸光度数据时出错: {e}')
            return

        # 初始化用于存储计算结果的Numpy矩阵，默认填充NaN
        num_diff_cols = len(diffwords) * len(wave)
        relative_height = np.full((len(timearr), 1), np.nan)
        gradient_matrix = np.full((len(timearr), num_diff_cols), np.nan)

        # 定义用于匹配纯数字活动的正则表达式
        pattern = r'^-?\d+(\.\d+)?$'

        # --- 3. 核心计算：遍历实验事件，填充矩阵 ---
        self.GuiRefresh(self.Status, 'Calculating slope...')
        for item in expInfo['schedule']:
            time_data = item.get('time')
            activity = item.get('activity')

            # 只处理有明确起止时间的事件段
            if not isinstance(time_data, (list, tuple)) or len(time_data) != 2:
                continue

            # 获取当前时间段在总时间数组中的布尔索引
            t = [time_data[0] + np.floor(timearr[0]), time_data[1] + np.floor(timearr[0])]
            index = np.logical_and(timearr >= t[0], timearr < t[1]).flatten()

            # 如果在时间段内无数据点，则跳过
            if not np.any(index):
                continue

            # 检查活动是否为纯数字，如果是，则填充“相对高度”矩阵
            height = float(activity) if re.fullmatch(pattern, activity) else np.nan
            if not np.isnan(height):
                relative_height[index] = height

            # 仅当活动为纯数字（如高度值）且数据点大于等于2时，才进行梯度计算
            if not np.isnan(height) and np.sum(index) >= 2:
                time_slice = timearr[index].reshape(-1, 1)
                # 对每一个差分-波长组合进行线性回归
                for cs in range(len(diffwords)):
                    for w_idx in range(len(wave)):
                        # 在原始数据和结果矩阵中找到对应的列索引
                        numpy_col_idx = (7 * cs) + w_idx
                        matrix_col_idx = cs * len(wave) + w_idx

                        spec_slice = all_diff_abs_data[index, numpy_col_idx].reshape(-1, 1)
                        slope, _ = self.calculate_gradient(time_slice, spec_slice)

                        # 将计算出的斜率填充到结果矩阵的相应位置
                        gradient_matrix[index, matrix_col_idx] = slope

        # --- 4. 结果写入Excel ---
        self.GuiRefresh(self.Status, 'Writing slope data...')
        # 去除重复数据，写入sheet中
        def remove_same_data(matrix):
            if np.isnan(matrix).any():
                matrix[np.isnan(matrix)] = 999  # 替换nan为999, 因为999一定不会出现在实验中

            _, unique_index = np.unique(matrix, axis=0, return_index=True)
            unique_index = np.sort(unique_index)
            matrix[matrix == 999] = np.nan
            return matrix[unique_index, :]

        unique_data = remove_same_data(gradient_matrix)
        unique_height = remove_same_data(relative_height)
        slope_sheet.range(3, 1).value = unique_height

        # 因存在分隔列，需将梯度矩阵分块写入
        data_offset = 2  # A, B列已被占用
        for cs in range(len(diffwords)):
            # 计算当前数据块在Excel中应写入的起始列
            start_col = data_offset + 1 + cs * (len(wave) + 1) + 1

            # 从完整的梯度矩阵中切出当前差分对的数据块
            matrix_start_col = cs * len(wave)
            matrix_end_col = (cs + 1) * len(wave)
            data_chunk = gradient_matrix[:, matrix_start_col:matrix_end_col]
            data_chunk2 = unique_data[:, matrix_start_col:matrix_end_col]

            # 将数据块写入“梯度分析”工作表的正确位置
            gradient_sheet.range((3, start_col)).value = data_chunk
            slope_sheet.range((3, start_col - data_offset + 1)).value = data_chunk2

        self.GuiRefresh(self.Status, 'Writing complete')

    def calculate_gradient(self, time, spec, method='lsq', alpha=1.0):
        """
        仅使用 NumPy 计算光谱变化的梯度（斜率）和 R²。

        支持普通最小二乘法、岭回归和 Lasso 回归。
        假设模型为 spec = slope * time + intercept。

        :param time: 时间序列 (1D array)
        :param spec: 光谱值序列 (1D array)
        :param method: 拟合方法。可选 'lsq', 'ridge', 或 'lasso'。
        :param alpha: 岭回归或 Lasso 回归的正则化强度。
        :return: slope (斜率), r2 (决定系数 R-squared)
        """
        assert len(time) == len(spec), "时间和光谱数组的长度必须相等。"
        n_samples = len(time)

        # --- 核心计算 ---

        if method == 'lsq':
            # 转换为1D数组
            time = time.reshape(-1)
            spec = spec.reshape(-1)
            w = np.polyfit(time, spec, 1)
            slope, intercept = w[0], w[1]

        elif method in ('ridge', 'lasso'):
            # 对于岭回归和 Lasso，为避免对截距进行正则化，我们先中心化数据
            time_mean = np.mean(time)
            spec_mean = np.mean(spec)
            time_c = time - time_mean
            spec_c = spec - spec_mean

            if method == 'ridge':
                # 应用岭回归公式计算斜率
                # slope = Σ(x_c * y_c) / (Σ(x_c²) + α)
                slope = np.sum(time_c * spec_c) / (np.sum(time_c ** 2) + alpha)

            elif method == 'lasso':
                # 对单变量问题，Lasso 的解可以通过软阈值函数得到
                # rho = Σ(x_c * y_c)
                rho = np.sum(time_c * spec_c)

                # scikit-learn 的 Lasso 目标函数是 (1/(2n)) * ||y-Xw||² + α||w||₁
                # 这导致阈值为 n * α
                lambda_ = n_samples * alpha

                sum_sq_time_c = np.sum(time_c ** 2)
                if sum_sq_time_c == 0:
                    # 如果所有时间点都相同，斜率无法定义
                    slope = 0.0
                elif rho > lambda_:
                    slope = (rho - lambda_) / sum_sq_time_c
                elif rho < -lambda_:
                    slope = (rho + lambda_) / sum_sq_time_c
                else:
                    # 如果 rho 的绝对值小于阈值，则斜率被压缩为 0
                    slope = 0.0

            # 根据中心化前的均值计算截距
            intercept = spec_mean - slope * time_mean

        else:
            raise NotImplementedError(
                "方法 '{}' 未实现。请选择 'lsq', 'ridge', 或 'lasso'。".format(method)
            )

        # --- R² (决定系数) 计算 ---

        # spec_pred = slope * time + intercept
        # ss_res = np.sum((spec - spec_pred) ** 2)
        # ss_tot = np.sum((spec - np.mean(spec)) ** 2)
        #
        # if ss_tot == 0:
        #     r2 = 1.0 if ss_res == 0 else 0.0
        # else:
        #     r2 = 1 - (ss_res / ss_tot)
        r2 = 1
        return slope, r2

    def create_refresh_button(self, wb, sheetnames, expInfo):
        """
        添加刷新按键
        默认添加到固定位置，调用固定的vba

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
        """
        if expInfo is None:
            return

        self.GuiRefresh(self.Status, "Adding refresh button... ")

        button_text = '刷新标注框'
        macro_name = "Updater.xlam!ChartAnnotationUpdater.RefreshAllChartAnnotations"  # 绑定的宏（如需模块前缀，加上 ModuleName.）

        # 按钮位置与大小（像素）
        left, top, width, height = self.CHART_LEFT + self.CHART_WIDTH*4, self.CHART_TOP, 160, 100
        ws = wb.sheets[sheetnames[4]]

        # 添加一个“表单控件按钮”
        btn = ws.api.Buttons().Add(left, top, width, height)
        btn.Characters.Text = button_text
        btn.OnAction = macro_name

        btn.ShapeRange.Fill.ForeColor.RGB = xw.utils.rgb_to_int(self.TITLE_BOX_COLOR)
        btn.Characters.Font.Name = "Times New Roman"
        btn.Characters.Font.Size = self.MAIN_TITLE_FONT_SIZE
        btn.Characters.Font.Bold = 1

    def create_charts(self, wb, sheetnames, timearr, wave, Ch, wn, rng_lcol, expInfo, Chpath, C):
        """
        创建图表

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            timearr: 时间数组
            wave: 波长列表
            Ch: 环数
            wn: 波长数量
            rng_lcol: 血糖数据列索引
            expInfo: 实验信息
            Chpath: 数据文件路径
            C: 是否为中文文件名
        """
        if not self.PLTCheckBox.isChecked():
            return

        # 设置时间格式
        for i in range(0, len(sheetnames)):
            wb.sheets[sheetnames[i]].api.Columns("A:A").NumberFormatLocal = "[$-x-systime]h:mm:ss AM/PM"

        # 重置图表列表
        self.charts = []

        # 图表配置信息
        charttitles = ['12环差分信号vs.室温', '23环差分信号vs.测头旁皮肤温度',
                       '1050nm单环吸光度vs.测头下实际温度', '1219nm单环吸光度',
                       '34环差分信号vs.加热功率', '45环差分信号vs.测头相对扶手高度(cm)',
                       '1314nm单环吸光度', '1409nm单环吸光度',
                       'Diff1550-Diff1050', '1050nm差分吸光度vs.测头下实际温度',
                       '1550nm单环吸光度', '1609nm单环吸光度']

        ringsindex = ['Diff12', 'Diff23', '1050', '1219',
                      'Diff34', 'Diff45', '1314', '1409',
                      'Diff1550-Diff1050', 'Diff1050', '1550', '1609']

        if self.TempCheckBox.isChecked():
            tempindex = ['4', '5', '12', '0',
                         '15', '0', '0', '0',
                         '0', '12', '0', '0']  # 对应sheet中的列，设置为0则不设置副坐标轴
        else:
            tempindex = ['0', '0', '0', '0',
                         '0', '0', '0', '0',
                         '0', '0', '0', '0']  # 对应sheet中的列，设置为0则不设置副坐标轴

        infoindex = [False, False, False, False,
                     True, True, True, True,
                     False, False, False, False]

        if self.OGTTCheckBox.isChecked():  # OGTT时的血糖值绘制准备
            # tempindex[5] = str(rng_lcol + 2)
            tempindex[8] = str(rng_lcol + 2)
            # charttitles[5] = '45环差分信号vs.血糖真值'
            charttitles[8] = charttitles[8] + ' vs.血糖值'

        if self.tempCorrelationCheckBox.isChecked():
            charttitles[11] = '温度校正后的波长差分'
            ringsindex[11] = 'Diff1550-Diff1050-temp'
            tempindex[11] = '0'

        pltN = len(charttitles)
        SRRange = 'A1:ZZ2'
        diffSheet = wb.sheets[sheetnames[4]]
        sglSheet = wb.sheets[sheetnames[1]]
        tempSheet = wb.sheets[sheetnames[7]]

        # 创建图表的详细实现
        self._create_individual_charts(wb, sheetnames, charttitles, ringsindex, tempindex, infoindex,
                                       wave, Ch, wn, timearr, expInfo, Chpath, C, rng_lcol, pltN, SRRange,
                                       diffSheet, sglSheet, tempSheet)

    # 数据处理主函数
    def DataProcess(self):
        """
        数据处理主函数

        执行完整的数据处理流程，包括：
        1. 数据加载和验证
        2. 温度数据处理
        3. 实验信息解析
        4. 基础数据计算
        5. Excel文件生成
        6. 图表绘制
        """
        try:
            self.Process.setEnabled(False)
            self.starttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, '')

            # 1. 加载和验证数据
            Chvalues, Ch, C, Chpath = self.load_and_validate_data()
            Chvalues = self.preprocess_chvalues(Chvalues)
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

            # 2. 处理温度数据
            Tempvalue, Temptitle = self.process_temperature_data(Chpath, C)

            # 3. 解析实验信息
            expInfo = self.parse_experiment_info(Chpath)

            # 4. 计算基本数据（时间、序列、波长等）
            n = len(Chvalues[0])  # 单环数据行数
            m = len(Chvalues[0][0])  # 每次测量数加一（最后一列是时间）
            datarange = [i for i in range(n)]

            # 根据LD选项确定波长
            if self.LDCheckBox.isChecked():
                wave = ['1064', '1310', '1390', '1550', '1625']
                # for i in range(len(datarange) - 1, -1, -1):
                #     if (datarange[i] % 6 == 5):
                #         del datarange[i]
            else:
                wave = ['1050', '1219', '1314', '1409', '1550', '1609']

            n = len(datarange)
            wn = len(wave)
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

            # 生成环和差分标签
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
                # timepop = self.popele(Chvalues[0][6 * l:6 * l + wn][:, -1], 0)
                # timeele = sum(timepop) / len(timepop)
                timepop = np.nanmax(Chvalues[0, 6 * l:6 * l + wn, :])
                timeele = np.nanmean(timepop)
                timearr[l] = timeele
                cycleNoarr[l] = l + 1

            # 5. 计算基准周期数据（支持多基准周期）
            basesingle_dict, basediff_dict = self.calculate_base_data(Chvalues, Ch, wn, m, expInfo)

            # 6. 创建Excel工作簿和工作表
            wb, sheetnames = self.create_excel_workbook(Chpath, C)
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

            # 7. 写入时间和标题
            self.write_time_and_headers(wb, sheetnames, timearr, cycleNoarr)
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

            # 8. 处理温度数据插值
            yinterp = self.process_temperature_interpolation(wb, sheetnames, timearr, cycleNoarr, Tempvalue, Temptitle)
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

            # 9. 写入汇总数据标题
            self.write_summary_data(wb, sheetnames, wave, ringwords, diffwords, Ch, C)

            # 10. 处理单环数据（支持多基准周期）
            self.process_single_ring_data(wb, sheetnames, Chvalues, Ch, ringwords, wave, datarange, basesingle_dict, n, wn, m, timearr, expInfo)

            # 11. 处理差分数据（支持多基准周期）
            self.process_differential_data(wb, sheetnames, Chvalues, Ch, diffwords, wave, datarange, basediff_dict, n, wn, m, timearr, expInfo)

            # 12. 处理动态基准周期
            self.handle_dynamic_base_cycle(wb, sheetnames, Ch, n, wn, C)

            # 13. 添加备注数据并计算梯度
            self.add_info_data(wb, sheetnames, expInfo, timearr)
            self.add_gradient_data(wb, sheetnames, expInfo, timearr, C, diffwords, wave)

            # 14. 添加血糖数据
            rng_lcol = self.add_glucose_data(wb, sheetnames, timearr, yinterp, C, expInfo=expInfo)

            # 15. 创建图表
            self.create_charts(wb, sheetnames, timearr, wave, Ch, wn, rng_lcol, expInfo, Chpath, C)

            # 16. 最终保存和清理
            wb.close()
            self.GuiRefresh(self.Status, 'Process Finished')
            self.Process.setEnabled(True)
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

        # 错误处理
        except Exception as ex:
            import traceback, sys
            self.GuiRefresh(self.ErrorText, str(ex))

            # 打印完整堆栈信息
            self.GuiRefresh(self.ErrorText, traceback.format_exc())

            # 只打印行号信息
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = exc_tb.tb_frame.f_code.co_filename
            lineno = exc_tb.tb_lineno
            self.GuiRefresh(self.ErrorText, f"出错位置: 文件 {fname}, 第 {lineno} 行")

            self.Process.setEnabled(True)
            try:
                wb.save()
                wb.close()
            except Exception:
                pass

    def _create_individual_charts(self, wb, sheetnames, charttitles, ringsindex, tempindex, infoindex,
                                  wave, Ch, wn, timearr, expInfo, Chpath, C, rng_lcol, pltN, SRRange,
                                  diffSheet, sglSheet, tempSheet):
        """
        创建具体的图表实现

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            charttitles: 图表标题列表
            ringsindex: 环标识列表
            tempindex: 温度数据索引列表
            infoindex: 信息索引列表
            wave: 波长列表
            Ch: 环数
            wn: 波长数量
            timearr: 时间数组
            expInfo: 实验信息
            Chpath: 数据文件路径
            C: 是否为中文文件名
            rng_lcol: 血糖数据列索引
            pltN: 图表数量
            SRRange: 搜索范围
            diffSheet, sglSheet, tempSheet: 各工作表对象
        """
        for p, each in enumerate(ringsindex):
            if each == '0':
                self.charts.append(None)
                continue

            # 确定数据源范围
            PltRangeS = 'A:A'  # 时间列

            # 副座标轴Series个数
            secondary_axis_series_count = 0

            if len(each) == 6 and '-' not in each:  # 差分数据，如'Diff12'
                datasheet = diffSheet
                addr = self.FindRowColRange(datasheet, 'Col', each, SRRange)
                addrstr = xw.utils.col_name(int(addr) + 1) + ':' + xw.utils.col_name(int(addr) + wn)
                PltRangeS = PltRangeS + ', ' + addrstr
                ytitle = 'ΔAd'

            elif len(each) == 6 and '-' in each:  # 提取特定的单环数据
                datasheet = sglSheet
                addr = self.FindRowColRange(datasheet, 'Col', each[0:4], SRRange)
                ring_num = int(each[-1]) - 1
                addrstr = xw.utils.col_name(int(addr) + 7 * ring_num) + ':' + xw.utils.col_name(int(addr) + 7 * ring_num)
                PltRangeS = PltRangeS + ', ' + addrstr
                ytitle = 'ΔA'

            elif len(each) == 4:  # 单环数据，如'1050'
                datasheet = sglSheet
                addr = self.FindRowColRange(datasheet, 'Col', each, SRRange)
                for i in range(0, Ch):
                    addrstr = xw.utils.col_name(int(addr) + 7 * i) + ':' + xw.utils.col_name(int(addr) + 7 * i)
                    PltRangeS = PltRangeS + ', ' + addrstr
                ytitle = 'ΔA'

            elif len(each) == 8:  # 特定波长的差分数据，如'Diff1050'
                datasheet = diffSheet
                waveIndex = wave.index(each[-4:]) + 1
                ytitle = 'ΔAd'
                pairs = [(i, i + 1) for i in range(1, Ch)]  # 相邻：12,23,34...
                pairs.append((3, 5))  # 添加35环

                for a, b in pairs:
                    target = f'Diff{a}{b}'
                    addr = self.FindRowColRange(datasheet, 'Col', target, SRRange)
                    addrstr = xw.utils.col_name(int(addr) + waveIndex) + ':' + xw.utils.col_name(int(addr) + waveIndex)
                    PltRangeS = PltRangeS + ', ' + addrstr

            elif len(each) >= 17:  # 差分计算，如'Diff1550-Diff1050'
                datasheet = diffSheet
                ytitle = 'ΔAd'
                wave1, wave2 = each[4:8], each[13:17]
                waveIndex1, waveIndex2 = wave.index(wave1) + 1, wave.index(wave2) + 1

                targets = ['Diff12', 'Diff23', 'Diff34', 'Diff35', 'Diff45']
                sheet_target = wb.sheets[sheetnames[4]]

                # 动态获取最后一列索引，并计算新的列名
                last_col = sheet_target.used_range.last_cell.column
                indices = [xw.utils.col_name(last_col + 2 + i) for i in range(len(targets))]
                # sheet_target.range(xw.utils.col_name(last_col + 1) + '1').value = wave1 + ' - ' + wave2
                sheet_target.range(xw.utils.col_name(last_col + 1) + '1').value = ' '   # 占位符，保证图例为 Diffxx XXXX - XXXX

                # 在 Excel 中插入计算公式
                for target, indice in zip(targets, indices):
                    base_addr = int(self.FindRowColRange(datasheet, 'Col', target, SRRange))
                    addr1, addr2 = xw.utils.col_name(base_addr + waveIndex1), xw.utils.col_name(base_addr + waveIndex2)

                    # 绝对引用公式
                    if len(each) == 17:
                        formula = f"=${addr1}2-${addr2}2"
                    elif len(each) == 22:
                        """
                        设置温度校正后的波长差分公式，并将1050温度校正系数和使用的温度列写入文件中
                        公式: $ y = Diff1550 - (Diff1050 - \phi*T) $
                        """
                        base_cycle = int(self.BaseCycle.value()) + 1
                        temp_corr_addr = xw.utils.col_name(last_col + len(targets) + 2) # 设置矫正配置项所在列
                        temp_addr = f'INDEX(温度数据!$L:$Q,ROW()-1,MATCH(${temp_corr_addr}$4,温度数据!$L$1:$Q$1,0))' # 列匹配公式
                        base_temp_addr = (
                            f'INDEX(温度数据!$L:$Q,{base_cycle},'
                            f'MATCH(${temp_corr_addr}$4,温度数据!$L$1:$Q$1,0))'
                        )

                        # 填充配置
                        sheet_target.range(xw.utils.col_name(last_col + 1) + '1').value = '温度校正'
                        sheet_target.range(temp_corr_addr + '1').value = 'Diff1050温度校正系数👇'
                        sheet_target.range(temp_corr_addr + '2').value = self.TempCorr1050
                        sheet_target.range(temp_corr_addr + '3').value = '温度列👇'
                        sheet_target.range(temp_corr_addr + '4').value = 'TC1实际温度'

                        # 组合公式
                        formula = f"=${addr1}2-${addr2}2+${temp_corr_addr}$2*({temp_addr}-{base_temp_addr})"

                    # 填充整个列（从 2 到最后一行）
                    sheet_target.range(f"{indice}2:{indice}{datasheet.used_range.last_cell.row}").formula = formula

                    # 设置表头信息
                    sheet_target.range(f"{indice}1").value = None
                    sheet_target.range(f"{indice}2").value = target + ' ' + wave1 + '-' + wave2

                PltRangeS += f", {indices[0]}1:{indices[-1]}{datasheet.used_range.last_cell.row}"
                if not self.waveDiffCheckBox.isChecked():
                    self.charts.append(None)
                    continue  # 仅处理数据但是不画图了

            elif len(each) <= 2:  # 温度数据
                datasheet = tempSheet
                addr = xw.utils.col_name(int(each))
                addrstr = addr + ':' + addr
                PltRangeS = PltRangeS + ', ' + addrstr
                ytitle = wb.sheets[sheetnames[7]].range(1, int(each)).value

            # 获取数据范围
            pltrange = datasheet.range(PltRangeS)

            # 设置副坐标轴数据范围
            secrange = None
            try:
                temp_col = int(tempindex[p])
            except TypeError:
                # 设置副坐标轴为光谱数据
                temp_col = -1
                secaddrstr = None
                if 'Diff' in tempindex[p][0]:
                    datasheet = diffSheet
                else:
                    datasheet = sglSheet

                addr_list = []
                for wave_target in tempindex[p]:
                    # 差分数据
                    wave_index = wave.index(wave_target[-4:]) + 1
                    addr = self.FindRowColRange(datasheet, 'Col', wave_target[:6], SRRange)
                    addrstr = xw.utils.col_name(int(addr) + wave_index) + ':' + xw.utils.col_name(
                        int(addr) + wave_index)
                    if secaddrstr is not None:
                        secaddrstr = secaddrstr + ', ' + addrstr
                    else:
                        secaddrstr = addrstr

                    addr_list.append(addrstr)

                secrange = datasheet.range(secaddrstr)

            if temp_col > 0:
                if temp_col > rng_lcol:  # 血糖数据
                    secaddr = xw.utils.col_name(temp_col)
                    secaddrstr = secaddr + ':' + secaddr
                    SecRangeS = secaddrstr
                    secrange = tempSheet.range(SecRangeS)
                else:  # 温度数据
                    secaddr = xw.utils.col_name(temp_col)
                    secaddrstr = secaddr + ':' + secaddr
                    SecRangeS = secaddrstr
                    secrange = tempSheet.range(SecRangeS)

            self.GuiRefresh(self.Status, 'Plotting ' + str(p + 1) + '/' + str(pltN))

            # 计算图表位置 - 为了解决diff1050和diff1550-diff1050绘图位置的问题
            # if len(each) <= 8:
            #     figure_lft = self.CHART_LEFT + self.CHART_WIDTH * int(p % 4)
            #     figure_top = self.CHART_TOP + self.CHART_HEIGHT * int(p / 4)
            #     figure_height = self.CHART_HEIGHT
            # else:
            #     # len(each) >= 8的情况，图表位置需要特殊处理
            #     figure_lft = self.CHART_LEFT + self.CHART_WIDTH * int(p % 4)
            #     figure_top = self.CHART_TOP + self.CHART_HEIGHT * int(p / 4)
            #     figure_height = self.CHART_HEIGHT

            figure_lft = (self.CHART_LEFT + self.CHART_WIDTH * int(p % 4))
            figure_top = (self.CHART_TOP + self.CHART_HEIGHT * int(p / 4))

            if self.waveDiffCheckBox.isChecked() and len(each) == 17:
                # 特殊处理Diff1550-Diff1050的位置
                figure_lft = self.CHART_LEFT
                figure_top = self.CHART_TOP + self.CHART_HEIGHT * 2

            # 创建图表
            chart = wb.sheets[sheetnames[4]].charts.add(
                left=figure_lft,
                top=figure_top,
                width=self.CHART_WIDTH,
                height=self.CHART_HEIGHT,
            )
            self.charts.append(chart)

            # 设置图表基本属性
            chart.chart_type = 'xy_scatter_lines'
            chart.set_source_data(pltrange)
            chart.api[0].Placement = xw.constants.Placement.xlFreeFloating
            chartApi = chart.api[1]
            chartApi.Legend.Position = -4107  # xlLegendPositionBottom

            # 设置X轴范围
            ratio = 0.05
            chartApi.Axes(1).MinimumScale = timearr.min() - ratio * (timearr.max() - timearr.min())
            chartApi.Axes(1).MaximumScale = timearr.max() + ratio * (timearr.max() - timearr.min())

            # 添加副坐标轴数据
            if secrange is not None:
                if temp_col > rng_lcol:  # 血糖数据特殊处理
                    chartApi.SeriesCollection().Add(Source=secrange.api, SeriesLabels=True)
                    series_count = chartApi.SeriesCollection().Count
                    chartApi.FullSeriesCollection(series_count).Name = "=" + sheetnames[7] + "!" + xw.utils.col_name(rng_lcol + 2) + "1"
                    chartApi.FullSeriesCollection(series_count).XValues = "=" + sheetnames[7] + "!" + xw.utils.col_name(rng_lcol + 1) + "2:" + xw.utils.col_name(rng_lcol + 1) + str(len(timearr))
                    chartApi.FullSeriesCollection(series_count).Values = "=" + sheetnames[7] + "!" + xw.utils.col_name(rng_lcol + 2) + "2:" + xw.utils.col_name(rng_lcol + 2) + str(len(timearr))
                    chartApi.SeriesCollection(series_count).AxisGroup = 2
                    chartApi.Axes(2, 2).ReversePlotOrder = True
                    chartApi.ChartColor = 10
                    chartApi.FullSeriesCollection(series_count).Format.Line.ForeColor.RGB = 255
                    chartApi.FullSeriesCollection(series_count).MarkerBackgroundColor = 255
                    chartApi.FullSeriesCollection(series_count).MarkerForegroundColor = 255
                    chartApi.SeriesCollection(series_count).Format.Line.Weight = self.LINE_WEIGHT
                    secondary_axis_series_count += 1
                else:  # 普通温度数据
                    if isinstance(tempindex[p], list):
                        last_row = datasheet.used_range.last_cell.row
                        x_range = datasheet.range(f'A3:A{last_row}')

                        for addr, wave_target in zip(addr_list, tempindex[p]):
                            new_series = chartApi.SeriesCollection().NewSeries()
                            new_series.Name = wave_target
                            column_letter = addr.split(':')[0]
                            y_range = datasheet.range(f'{column_letter}3:{column_letter}{last_row}')

                            # 5. 将精确范围的 .api 属性赋值给系列
                            new_series.XValues = x_range.api
                            new_series.Values = y_range.api

                            # 6. 将新系列分配给副坐标轴并设置格式
                            new_series.AxisGroup = 2
                            new_series.Format.Line.Weight = self.LINE_WEIGHT
                            secondary_axis_series_count += 1

                    else:
                        chartApi.SeriesCollection().Add(Source=secrange.api, SeriesLabels=True)
                        chartApi.ChartColor = 10
                        series_count = chartApi.SeriesCollection().Count
                        chartApi.SeriesCollection(series_count).AxisGroup = 2
                        chartApi.SeriesCollection(series_count).Format.Line.Weight = self.LINE_WEIGHT
                        secondary_axis_series_count += 1

            # 设置系列标记样式
            for i in range(1, chartApi.SeriesCollection().Count + 1):
                series = chartApi.SeriesCollection(i)
                series.MarkerStyle = 8  # 圆形标记
                series.MarkerSize = 5

            self._configure_chart_appearance(chartApi, charttitles[p], ytitle, temp_col,wb, sheetnames, rng_lcol)

            # 添加实验信息标注
            if expInfo and (self.expInfoCheckBox.isChecked() or infoindex[p]):
                # with_subaxis = False if temp_col <= 0 else True
                self._add_experiment_annotations(chartApi, expInfo, timearr, p, secondary_axis_series_count)

            # 图表复制并放大
            if self.duplicateCheckBox.isChecked() and p in self.Duplicate_Target["index"]:
                duplicate_index = self.Duplicate_Target["index"].index(p)
                duplicated_chart = chartApi.Parent.Duplicate()
                duplicated_chart.Top = 200 + self.CHART_TOP + self.CHART_HEIGHT*3 + duplicate_index*1.2*self.CHART_HEIGHT
                duplicated_chart.Left = self.CHART_LEFT + self.CHART_WIDTH
                duplicated_chart.Width = 2.5 * self.CHART_WIDTH
                duplicated_chart.Height = 1.2 * self.CHART_HEIGHT

                for ignore_index in self.Duplicate_Target["ignore_series"][duplicate_index]:
                    duplicated_chart.Chart.FullSeriesCollection(ignore_index).IsFiltered = True

                self.AXIS_FONT_SIZE += 6
                self.AXIS_TITLE_FONT_SIZE += 6
                self._configure_chart_appearance(duplicated_chart.Chart, charttitles[p], ytitle, temp_col, wb, sheetnames, rng_lcol)
                self.AXIS_FONT_SIZE -= 6
                self.AXIS_TITLE_FONT_SIZE -= 6

            # 更新UI
            self.currenttime = datetime.datetime.now()
            self.GuiRefresh(self.ErrorText, 'Process time: ' + str(self.currenttime - self.starttime).split('.')[0])

        # 添加备注和标题
        self._add_chart_annotations(wb, sheetnames, expInfo, Chpath)

        # 添加刷新按键函数
        # self.create_refresh_button(wb, sheetnames, expInfo)
        self.GuiRefresh(self.Status, 'Saving...')
        wb.save()

    def _configure_chart_appearance(self,
                                    chartApi,
                                    chart_title,
                                    ytitle,
                                    temp_index,
                                    wb,
                                    sheetnames,
                                    rng_lcol,
                                    title_font_size=18,
                                    ):
        """
        配置图表外观样式

        Args:
            chartApi: 图表API对象
            chart_title: 图表标题
            ytitle: Y轴标题
            temp_index: 温度数据索引
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            rng_lcol: 血糖数据列索引
        """
        # 图表整体样式
        chartApi.ChartArea.Format.Line.ForeColor.RGB = self.CHART_BORDER_COLOR

        # 坐标轴格式
        chartApi.Axes(1).MajorTickMark = xw.constants.Constants.xlCross
        chartApi.Axes(1).Format.Line.Weight = self.AXIS_WEIGHT
        chartApi.Axes(1).Format.Line.ForeColor.RGB = 0
        chartApi.Axes(2, 1).MajorTickMark = xw.constants.Constants.xlInside
        chartApi.Axes(2, 1).Format.Line.Weight = self.AXIS_WEIGHT
        chartApi.Axes(2, 1).Format.Line.ForeColor.RGB = 0
        chartApi.Axes(2, 1).MinorUnit = 0.001 if ytitle == 'ΔAd' else 0.01

        if int(temp_index) != 0:
            chartApi.Axes(2, 2).MajorTickMark = xw.constants.Constants.xlInside
            chartApi.Axes(2, 2).Format.Line.Weight = self.AXIS_WEIGHT
            chartApi.Axes(2, 2).Format.Line.ForeColor.RGB = 0

        # X轴格式
        chartApi.Axes(1).TickLabels.NumberFormatLocal = "h:mm;@"

        # 网格线
        chartApi.SetElement(334)  # 主要网格线
        chartApi.SetElement(330)  # 次要网格线
        chartApi.Axes(1).MajorGridlines.Format.Line.ForeColor.RGB = self.GRID_COLOR
        chartApi.Axes(1).MajorGridlines.Format.Line.Weight = self.GRID_WEIGHT
        chartApi.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = self.GRID_COLOR
        chartApi.Axes(2).MajorGridlines.Format.Line.Weight = self.GRID_WEIGHT

        # 图表标题
        chartApi.SetElement(2)
        chartApi.ChartTitle.Format.TextFrame2.TextRange.Characters.Text = chart_title
        chartApi.ChartTitle.Format.TextFrame2.TextRange.Font.Name = "Calibri"
        chartApi.ChartTitle.Format.TextFrame2.TextRange.Characters.Font.Size = self.TITLE_FONT_SIZE
        chartApi.ChartTitle.Format.TextFrame2.TextRange.Characters.Font.Bold = 1

        # Y轴标题
        chartApi.Axes(2, 1).HasTitle = True
        chartApi.Axes(2, 1).AxisTitle.Characters.Text = ytitle
        chartApi.Axes(2, 1).AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Calibri"
        chartApi.Axes(2, 1).AxisTitle.Format.TextFrame2.TextRange.Font.Size = self.AXIS_TITLE_FONT_SIZE
        chartApi.Axes(2, 1).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = 1

        if int(temp_index) != 0:
            # Y2轴标题
            chartApi.Axes(2, 2).HasTitle = True
            if temp_index == -1:
                chartApi.Axes(2, 2).AxisTitle.Characters.Text = 'ΔAd'
            else:
                chartApi.Axes(2, 2).AxisTitle.Characters.Text = wb.sheets[sheetnames[7]].range(1, int(temp_index)).value
            chartApi.Axes(2, 2).AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Calibri"
            chartApi.Axes(2, 2).AxisTitle.Format.TextFrame2.TextRange.Font.Size = self.AXIS_TITLE_FONT_SIZE
            chartApi.Axes(2, 2).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = 1

        # 坐标轴字体
        chartApi.Axes(1).TickLabels.Font.Size = self.AXIS_FONT_SIZE
        chartApi.Axes(1).TickLabels.Font.Bold = 1
        chartApi.Axes(2, 1).TickLabels.Font.Size = self.AXIS_FONT_SIZE
        chartApi.Axes(2, 1).TickLabels.Font.Bold = 1
        if int(temp_index) != 0:
            chartApi.Axes(2, 2).TickLabels.Font.Size = self.AXIS_FONT_SIZE
            chartApi.Axes(2, 2).TickLabels.Font.Bold = 1

        # 图例
        chartApi.Legend.Format.TextFrame2.TextRange.Font.Size = self.LEGEND_FONT_SIZE - 1
        chartApi.Legend.Format.TextFrame2.TextRange.Font.Bold = 1

        # 统一设置线宽
        series_count = chartApi.SeriesCollection().Count
        for count in range(1, series_count):
            chartApi.FullSeriesCollection(count).Format.Line.Weight = self.LINE_WEIGHT

    def _add_experiment_annotations(self, chartApi, expInfo, timearr, chart_index, secondary_axis_series_count):
        """
        添加实验信息标注到图表

        Args:
            chartApi: 图表API对象
            expInfo: 实验信息字典
            timearr: 时间数组
            chart_index: 图表索引
        """
        # 获取当前系列数量，用于后续图例删除
        initial_series_count = chartApi.SeriesCollection().Count + 1
        # initial_series_count -= secondary_axis_series_count if secondary_axis_series_count <= 0 else 0
        initial_series_count -= secondary_axis_series_count
        # axis_index = 2 if secondary_axis_series_count > 0 else 1
        axis_index = 1
        target_axis = chartApi.Axes(2, axis_index)

        # 定义正则表达式，检测activity是否为纯数字
        pattern = r'^-?\d+(\.\d+)?$'

        ratio = 0.15 # 文字框缩放比例

        for nitem, item in enumerate(expInfo['schedule']):
            # 处理时间数据
            time_data = item.get('time')
            activity = item.get('activity')
            if time_data is None:
                continue

            # 确保时间数据是可迭代的
            if isinstance(time_data, (int, float)):
                t = [time_data + np.floor(timearr[0])]
            elif isinstance(time_data, (list, tuple)) and len(time_data) == 2:
                t = [time_data[0] + np.floor(timearr[0]), time_data[1] + np.floor(timearr[0])]
            else:
                continue

            # 添加新的数据序列
            y_min = target_axis.MinimumScale
            y_max = target_axis.MaximumScale

            if re.fullmatch(pattern, activity):
                p_min = y_min
                p_max = y_max
            else:
                scale = y_max - y_min
                p_min = y_min + scale * ratio
                p_max = y_max - scale * ratio

            series = chartApi.SeriesCollection().NewSeries()
            series.AxisGroup = axis_index
            itemColor = self.hexColor2Int(self.CHART_COLORS[nitem % len(self.CHART_COLORS)])

            if len(t) == 2:  # 时间段 - 画框
                pointIdx = 3 if nitem % 2 == 0 else 6
                t_point = t[0] + (t[1] - t[0]) / 2

                series.XValues = [t[0], t[0], t_point, t[1], t[1], t_point, t[0]]
                series.Values = [p_min, p_max, p_max, p_max, p_min, p_min, p_min]

            elif len(t) == 1:  # 时间点 - 画线
                pointIdx = 2 if nitem % 2 == 0 else 1

                series.XValues = [t[0], t[0]]
                series.Values = [p_min, p_max]
            else:
                # 未知类型，跳过
                continue

            target_axis.MinimumScale = y_min
            target_axis.MaximumScale = y_max

            # 设置序列格式
            series.ChartType = 75  # xlLine
            series.Format.Line.Weight = self.ANNOTATION_LINE_WEIGHT
            series.Format.Line.DashStyle = 4  # 虚线
            series.Format.Line.ForeColor.RGB = itemColor
            series.Name = f"_ANNOTATION_BOX_{nitem}"

            # 在特定图表上添加标签
            if chart_index == 5:  # 第6个图表
                series.Points(pointIdx).ApplyDataLabels()
                series.Points(pointIdx).DataLabel.Text = item["activity"]
                series.Points(pointIdx).DataLabel.Font.Size = self.LABEL_FONT_SIZE
                series.Points(pointIdx).DataLabel.Font.Bold = 1
                series.Points(pointIdx).DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = itemColor
                series.Points(pointIdx).DataLabel.Position = xw.constants.DataLabelPosition.xlLabelPositionCenter

            # 隐藏图例 - 删除标注线的图例项
            current_series_count = chartApi.SeriesCollection().Count
            leg = chartApi.Legend.LegendEntries(initial_series_count)
            leg.Delete()


    def _add_chart_annotations(self, wb, sheetnames, expInfo, Chpath):
        """
        添加图表标注和备注

        Args:
            wb: Excel工作簿对象
            sheetnames: 工作表名称列表
            expInfo: 实验信息
            Chpath: 数据文件路径
        """
        # 添加备注文本框
        if expInfo:
            self.GuiRefresh(self.Status, 'Adding remarks')
            textbox = wb.sheets[sheetnames[4]].shapes.api.AddTextbox(
                Orientation=1,  # 水平方向
                Left=self.CHART_LEFT + self.CHART_WIDTH * 2,
                Top=self.CHART_TOP + self.CHART_HEIGHT * 1,
                Width=self.CHART_WIDTH / 2,
                Height=self.CHART_HEIGHT,
            )
            # 显示备注内容，如果没有备注则显示默认文本
            remark_text = "备注:\n" + (expInfo.get('remark', '') if expInfo.get('remark') else "")
            textbox.TextFrame2.TextRange.Characters.Text = remark_text
            textbox.TextFrame2.TextRange.ParagraphFormat.Alignment = 2  # 居中
            textbox.TextFrame2.TextRange.Characters.Font.Name = "Times New Roman"
            textbox.TextFrame2.TextRange.Characters.Font.Size = self.ANNOTATION_FONT_SIZE
            textbox.TextFrame2.TextRange.Characters.Font.Bold = 1
            textbox.Placement = xw.constants.Placement.xlFreeFloating

        # 添加标题文本框
        titlebox = wb.sheets[sheetnames[4]].shapes.api.AddTextbox(
            Orientation=1,
            Left=self.CHART_LEFT,
            Top=self.CHART_TOP - 100,
            Width=self.CHART_WIDTH * 3,
            Height=100,
        )
        titlebox.Fill.ForeColor.RGB = xw.utils.rgb_to_int(self.TITLE_BOX_COLOR)
        titlebox.TextFrame2.TextRange.Characters.Text = Chpath.split('\\')[-2]
        titlebox.TextFrame2.VerticalAnchor = 3  # 垂直居中
        titlebox.TextFrame2.TextRange.Characters.Font.Name = "Times New Roman"
        titlebox.TextFrame2.TextRange.Characters.Font.Size = self.MAIN_TITLE_FONT_SIZE
        titlebox.TextFrame2.TextRange.Characters.Font.Bold = 1
        titlebox.Placement = xw.constants.Placement.xlFreeFloating

        if self.duplicateCheckBox.isChecked():
            duplicated_titlebox = titlebox.Duplicate()
            duplicated_titlebox.Top = 200 + self.CHART_TOP + self.CHART_HEIGHT * 3 - 100
            duplicated_titlebox.Left = self.CHART_LEFT + self.CHART_WIDTH
            duplicated_titlebox.Width = 2.5 * self.CHART_WIDTH

    # ==================== 窗口事件处理 ====================

    def closeEvent(self, event):
        """
        重写窗口关闭事件

        Args:
            event: 关闭事件对象
        """
        reply = QtWidgets.QMessageBox.question(self,
                                               'Exit',
                                               "Confirm Exit?",
                                               QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                                               QtWidgets.QMessageBox.StandardButton.No)
        if reply == QtWidgets.QMessageBox.StandardButton.Yes:
            event.accept()
        else:
            event.ignore()


if __name__ == "__main__":
    """
    程序入口点

    创建QApplication实例，显示主窗口，并启动事件循环
    """
    app = QApplication(sys.argv)
    form = GUI_Dialog()
    app.processEvents()
    form.show()

    app.exec()

    # 清理Excel应用程序
    form.xwapp.quit()
    try:
        form.xwapp.kill()
    except Exception as e:
        pass
