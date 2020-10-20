# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'win.ui'
#
# Created by: PyQt5 UI code generator 5.14.2
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(486, 112)
        Form.setStyleSheet("#pushButton_4{\n"
"padding:2px 2px 2px 2px;\n"
"background-color: rgba(255,255,255,0);\n"
"border-style:none;\n"
"border-radius: 5px;\n"
"image:url(:/newPrefix/关闭1.png);\n"
"width:15px;height:15px;\n"
"}\n"
"\n"
"#pushButton_4:hover{\n"
"padding:2px 2px 2px 2px;\n"
"image:url(:/newPrefix/关闭.png);\n"
"background-color: rgb(220, 220, 220,100);\n"
"border-radius: 5px;\n"
"width:15px;height:15px;\n"
"}\n"
"#pushButton_4:pressed {\n"
"padding:2px 2px 2px 2px;\n"
"    image:url(:/newPrefix/关闭1.png);\n"
"    /* 改变背景色 */  \n"
"    background-color:rgb(180, 180, 180,100);  \n"
"    /* 改变边框风格 */  \n"
"    border-style:inset;  \n"
"    /* 使文字有一点移动 */  \n"
"    padding-left:1px;  \n"
"    padding-top:1px;  \n"
"border-radius: 5px;\n"
"}\n"
"QLabel{\n"
"color:rgb(255, 0, 0)\n"
"}")
        self.verticalLayout = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setSpacing(0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.widget = QtWidgets.QWidget(Form)
        self.widget.setMinimumSize(QtCore.QSize(0, 30))
        self.widget.setMaximumSize(QtCore.QSize(16777215, 30))
        self.widget.setObjectName("widget")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.widget)
        self.horizontalLayout_2.setContentsMargins(10, 0, 10, 0)
        self.horizontalLayout_2.setSpacing(0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label = QtWidgets.QLabel(self.widget)
        font = QtGui.QFont()
        font.setFamily("楷体")
        font.setPointSize(11)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout_2.addWidget(self.label, 0, QtCore.Qt.AlignHCenter)
        self.pushButton_4 = QtWidgets.QPushButton(self.widget)
        self.pushButton_4.setMinimumSize(QtCore.QSize(20, 20))
        self.pushButton_4.setMaximumSize(QtCore.QSize(20, 20))
        self.pushButton_4.setStyleSheet("")
        self.pushButton_4.setText("")
        self.pushButton_4.setObjectName("pushButton_4")
        self.horizontalLayout_2.addWidget(self.pushButton_4)
        self.verticalLayout.addWidget(self.widget)
        self.widget_2 = QtWidgets.QWidget(Form)
        self.widget_2.setObjectName("widget_2")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.widget_2)
        self.horizontalLayout.setContentsMargins(20, -1, 20, -1)
        self.horizontalLayout.setSpacing(10)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.pushButton = QtWidgets.QPushButton(self.widget_2)
        self.pushButton.setMinimumSize(QtCore.QSize(50, 0))
        self.pushButton.setMaximumSize(QtCore.QSize(50, 16777215))
        self.pushButton.setObjectName("pushButton")
        self.horizontalLayout.addWidget(self.pushButton)
        self.lineEdit = QtWidgets.QLineEdit(self.widget_2)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)
        self.pushButton_2 = QtWidgets.QPushButton(self.widget_2)
        self.pushButton_2.setObjectName("pushButton_2")
        self.horizontalLayout.addWidget(self.pushButton_2)
        self.verticalLayout.addWidget(self.widget_2)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "生成word机"))
        self.pushButton.setText(_translate("Form", "选择"))
        self.pushButton_2.setText(_translate("Form", "生成word"))
import img_rc
