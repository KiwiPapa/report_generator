# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'report_generator.ui'
#
# Created by: PyQt5 UI code generator 5.14.2
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(629, 379)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.radioButton = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton.setGeometry(QtCore.QRect(200, 130, 89, 16))
        self.radioButton.setObjectName("radioButton")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(200, 80, 54, 12))
        self.label.setObjectName("label")
        self.radioButton_2 = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_2.setGeometry(QtCore.QRect(200, 150, 89, 16))
        self.radioButton_2.setObjectName("radioButton_2")
        self.radioButton_3 = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_3.setGeometry(QtCore.QRect(60, 100, 89, 16))
        self.radioButton_3.setObjectName("radioButton_3")
        self.radioButton_4 = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_4.setGeometry(QtCore.QRect(60, 120, 89, 16))
        self.radioButton_4.setObjectName("radioButton_4")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(60, 80, 54, 12))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(320, 90, 201, 16))
        self.label_3.setObjectName("label_3")
        self.radioButton_5 = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_5.setGeometry(QtCore.QRect(330, 120, 89, 16))
        self.radioButton_5.setObjectName("radioButton_5")
        self.radioButton_6 = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_6.setGeometry(QtCore.QRect(330, 140, 89, 16))
        self.radioButton_6.setObjectName("radioButton_6")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(400, 140, 101, 16))
        self.label_4.setObjectName("label_4")
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(500, 130, 81, 31))
        self.textEdit.setAcceptDrops(True)
        self.textEdit.setAcceptRichText(True)
        self.textEdit.setObjectName("textEdit")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(60, 160, 101, 16))
        self.label_5.setObjectName("label_5")
        self.radioButton_7 = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_7.setGeometry(QtCore.QRect(60, 200, 89, 16))
        self.radioButton_7.setObjectName("radioButton_7")
        self.radioButton_8 = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_8.setGeometry(QtCore.QRect(60, 180, 89, 16))
        self.radioButton_8.setObjectName("radioButton_8")
        self.radioButton_9 = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_9.setGeometry(QtCore.QRect(200, 110, 89, 16))
        self.radioButton_9.setObjectName("radioButton_9")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(40, 30, 131, 31))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(170, 260, 81, 21))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(400, 260, 91, 41))
        self.pushButton_2.setObjectName("pushButton_2")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(50, 290, 231, 20))
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(50, 260, 81, 21))
        self.pushButton_3.setObjectName("pushButton_3")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "固井水泥胶结测井评价报告自动生成系统"))
        self.radioButton.setText(_translate("MainWindow", "威远"))
        self.label.setText(_translate("MainWindow", "模板选择"))
        self.radioButton_2.setText(_translate("MainWindow", "长宁"))
        self.radioButton_3.setText(_translate("MainWindow", "有储层"))
        self.radioButton_4.setText(_translate("MainWindow", "无储层"))
        self.label_2.setText(_translate("MainWindow", "有无储层"))
        self.label_3.setText(_translate("MainWindow", "液面高度是否和开始评价深度一致？"))
        self.radioButton_5.setText(_translate("MainWindow", "一致"))
        self.radioButton_6.setText(_translate("MainWindow", "不一致"))
        self.label_4.setText(_translate("MainWindow", "若不一致请输入："))
        self.label_5.setText(_translate("MainWindow", "有无固井质量差层"))
        self.radioButton_7.setText(_translate("MainWindow", "没有"))
        self.radioButton_8.setText(_translate("MainWindow", "有"))
        self.radioButton_9.setText(_translate("MainWindow", "普通"))
        self.label_6.setText(_translate("MainWindow", "程序参数选择"))
        self.pushButton.setText(_translate("MainWindow", "完整性检测"))
        self.pushButton_2.setText(_translate("MainWindow", "报告生成"))
        self.pushButton_3.setText(_translate("MainWindow", "登记卡选择"))
