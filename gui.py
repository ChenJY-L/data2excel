# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'GUI.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_Data_Processing(object):
    def setupUi(self, Data_Processing):
        Data_Processing.setObjectName("Data_Processing")
        # Data_Processing.resize(455, 250)
        Data_Processing.setFixedSize(480, 250)
        font = QtGui.QFont()
        font.setFamily("Calibri")
        
        # 环的选择 
        self.Rings = QtWidgets.QComboBox(Data_Processing)
        self.Rings.setGeometry(QtCore.QRect(20, 20, 151, 31))
        font.setPointSize(12)
        self.Rings.setFont(font)
        self.Rings.setObjectName("Rings")
        items = ["5 Rings", "7 Rings"]
        self.Rings.addItems(items)
        
        # 文件路径及其标签
        self.Path = QtWidgets.QTextEdit(Data_Processing)
        self.Path.setGeometry(QtCore.QRect(200, 20, 160, 31))
        font.setPointSize(9)
        self.Path.setFont(font)
        self.Path.setObjectName("Path")

        self.FileSelect = QtWidgets.QPushButton(Data_Processing)
        self.FileSelect.setGeometry(QtCore.QRect(360, 20, 71, 31))
        font.setPointSize(12)
        self.FileSelect.setFont(font)
        self.FileSelect.setObjectName("FileSelect")
        

        # 数据校准与否的选择
        self.Original = QtWidgets.QComboBox(Data_Processing)
        self.Original.setGeometry(QtCore.QRect(20, 70, 151, 31))
        font.setPointSize(12)
        self.Original.setFont(font)
        self.Original.setObjectName("Original")
        self.Original.addItem("")
        self.Original.addItem("")
        
        # 状态及其标签
        self.Status = QtWidgets.QTextEdit(Data_Processing)
        self.Status.setGeometry(QtCore.QRect(200, 70, 161, 31))
        font.setPointSize(9)
        self.Status.setFont(font)
        self.Status.setObjectName("Status")
        
        self.StatusLabel = QtWidgets.QLabel(Data_Processing)
        self.StatusLabel.setGeometry(QtCore.QRect(370, 70, 81, 30))
        font.setPointSize(12)
        self.StatusLabel.setFont(font)
        self.StatusLabel.setObjectName("StatusLabel")
        
        # 计算按钮
        self.Process = QtWidgets.QPushButton(Data_Processing)
        self.Process.setGeometry(QtCore.QRect(330, 116, 101, 31))
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.Process.setFont(font)
        self.Process.setAutoFillBackground(False)
        self.Process.setStyleSheet("color: rgb(255, 0, 0);")
        self.Process.setObjectName("Process")
        font.setBold(False)
        
        # 温度勾选框
        self.TempCheckBox = QtWidgets.QCheckBox(Data_Processing)
        self.TempCheckBox.setGeometry(QtCore.QRect(20, 110, 201, 41))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Maximum, QtWidgets.QSizePolicy.Policy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.TempCheckBox.sizePolicy().hasHeightForWidth())
        self.TempCheckBox.setSizePolicy(sizePolicy)
        self.TempCheckBox.setMaximumSize(QtCore.QSize(16777215, 16777215))
        font.setPointSize(12)
        self.TempCheckBox.setFont(font)
        self.TempCheckBox.setIconSize(QtCore.QSize(100, 100))
        self.TempCheckBox.setChecked(True)
        self.TempCheckBox.setAutoRepeatInterval(100)
        self.TempCheckBox.setObjectName("TempCheckBox")
        
        # 基准周期及其标签
        self.BaseCycle = QtWidgets.QSpinBox(Data_Processing)
        self.BaseCycle.setGeometry(QtCore.QRect(250, 120, 71, 22))
        font.setPointSize(12)
        self.BaseCycle.setFont(font)
        self.BaseCycle.setMinimum(1)
        self.BaseCycle.setMaximum(99999999)
        self.BaseCycle.setProperty("value", 1)
        self.BaseCycle.setObjectName("BaseCycle")
        
        self.BaseCycleLabel = QtWidgets.QLabel(Data_Processing)
        self.BaseCycleLabel.setGeometry(QtCore.QRect(170, 120, 91, 21))
        font.setPointSize(12)
        self.BaseCycleLabel.setFont(font)
        self.BaseCycleLabel.setObjectName("BaseCycleLabel")

        # Plot勾选框
        self.PLTCheckBox = QtWidgets.QCheckBox(Data_Processing)
        self.PLTCheckBox.setGeometry(QtCore.QRect(20, 140, 201, 41))
        # sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Maximum, QtWidgets.QSizePolicy.Policy.Maximum)
        # sizePolicy.setHeightForWidth(self.PLTCheckBox.sizePolicy().hasHeightForWidth())
        # self.PLTCheckBox.setSizePolicy(sizePolicy)
        font.setPointSize(12)
        self.PLTCheckBox.setFont(font)
        # self.PLTCheckBox.setIconSize(QtCore.QSize(100, 100))
        self.PLTCheckBox.setChecked(True)
        self.PLTCheckBox.setAutoRepeatInterval(100)
        self.PLTCheckBox.setObjectName("PLTCheckBox")

        # OGTT勾选框
        self.OGTTCheckBox = QtWidgets.QCheckBox(Data_Processing)
        self.OGTTCheckBox.setGeometry(QtCore.QRect(90, 140, 141, 41))
        font.setPointSize(12)
        self.OGTTCheckBox.setFont(font)
        self.OGTTCheckBox.setChecked(False)
        # self.OGTTCheckBox.setAutoRepeatInterval(100)
        self.OGTTCheckBox.setObjectName("OGTTCheckBox")

        # DyBC勾选框
        self.DyBCCheckBox = QtWidgets.QCheckBox(Data_Processing)
        self.DyBCCheckBox.setGeometry(QtCore.QRect(135, 140, 161, 41))
        self.DyBCCheckBox.setLayoutDirection(QtCore.Qt.LayoutDirection.RightToLeft)
        font.setPointSize(12)
        self.DyBCCheckBox.setFont(font)
        self.DyBCCheckBox.setChecked(False)
        self.DyBCCheckBox.setObjectName("DyBCCheckBox")

        # LD勾选框
        self.LDCheckBox = QtWidgets.QCheckBox(Data_Processing)
        self.LDCheckBox.setGeometry(QtCore.QRect(316, 140, 191, 41))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Maximum, QtWidgets.QSizePolicy.Policy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.LDCheckBox.sizePolicy().hasHeightForWidth())
        self.LDCheckBox.setSizePolicy(sizePolicy)
        self.LDCheckBox.setMaximumSize(QtCore.QSize(16777215, 16777215))
        font.setPointSize(12)
        self.LDCheckBox.setFont(font)
        self.LDCheckBox.setIconSize(QtCore.QSize(100, 100))
        self.LDCheckBox.setChecked(False)
        self.LDCheckBox.setAutoRepeatInterval(100)
        self.LDCheckBox.setObjectName("LDCheckBox")

        # 部分标注勾选框
        self.expInfoCheckBox = QtWidgets.QCheckBox(Data_Processing)
        self.expInfoCheckBox.setGeometry(QtCore.QRect(380, 140, 221, 41))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Maximum, QtWidgets.QSizePolicy.Policy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.expInfoCheckBox.sizePolicy().hasHeightForWidth())
        self.expInfoCheckBox.setSizePolicy(sizePolicy)
        self.expInfoCheckBox.setMaximumSize(QtCore.QSize(16777215, 16777215))
        font.setPointSize(12)
        self.expInfoCheckBox.setFont(font)
        self.expInfoCheckBox.setIconSize(QtCore.QSize(100, 100))
        self.expInfoCheckBox.setChecked(False)
        self.expInfoCheckBox.setAutoRepeatInterval(100)
        self.expInfoCheckBox.setObjectName("expInfoCheckBox")

        # 错误及其标签
        self.ErrorText = QtWidgets.QTextEdit(Data_Processing)
        self.ErrorText.setGeometry(QtCore.QRect(20, 190, 361, 51))
        font.setPointSize(9)
        self.ErrorText.setFont(font)
        self.ErrorText.setObjectName("ErrorText")
        
        self.ErrorLabel = QtWidgets.QLabel(Data_Processing)
        self.ErrorLabel.setGeometry(QtCore.QRect(390, 190, 71, 22))
        font.setPointSize(12)
        self.ErrorLabel.setFont(font)
        self.ErrorLabel.setObjectName("ErrorLabel")

        # 设置功能 —— 连接功能区
        self.retranslateUi(Data_Processing)
        QtCore.QMetaObject.connectSlotsByName(Data_Processing)
    
    # 程序自带转译区
    def retranslateUi(self, Data_Processing):
        _translate = QtCore.QCoreApplication.translate
        # Data_Processing.setWindowTitle(_translate("Data_Processing", "Data_Processing 20240705"))
        self.Original.setItemText(0, _translate("Data_Processing", "Calibrated"))
        self.Original.setItemText(1, _translate("Data_Processing", "Original"))
        self.Process.setText(_translate("Data_Processing", "Process"))
        self.TempCheckBox.setText(_translate("Data_Processing", "Temperature?"))
        self.ErrorLabel.setText(_translate("Data_Processing", "Error"))
        self.StatusLabel.setText(_translate("Data_Processing", "Status"))
        self.FileSelect.setText(_translate("Data_Processing", "File"))
        self.PLTCheckBox.setText(_translate("Data_Processing", "Plot?"))
        self.OGTTCheckBox.setText(_translate("Data_Processing", "OGTT?"))
        self.BaseCycleLabel.setText(_translate("Data_Processing", "BaseCycle"))
        self.DyBCCheckBox.setText(_translate("Data_Processing", "Dyna Basecycle?"))
        self.LDCheckBox.setText(_translate("Data_Processing", "LD?"))
        self.expInfoCheckBox.setText(_translate("Data_Processing", "Info for all?"))
