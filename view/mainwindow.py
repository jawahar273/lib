# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'mainwindow.ui'
#
# Created by: PyQt4 UI code generator 4.11.4
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName(_fromUtf8("MainWindow"))
        MainWindow.resize(826, 745)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(_fromUtf8("view/ring-alt.gif")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setIconSize(QtCore.QSize(500, 500))
        MainWindow.setToolButtonStyle(QtCore.Qt.ToolButtonFollowStyle)
        self.centralWidget = QtGui.QWidget(MainWindow)
        self.centralWidget.setObjectName(_fromUtf8("centralWidget"))
        self.gridLayout = QtGui.QGridLayout(self.centralWidget)
        self.gridLayout.setMargin(11)
        self.gridLayout.setSpacing(6)
        self.gridLayout.setObjectName(_fromUtf8("gridLayout"))
        self.tabWidget = QtGui.QTabWidget(self.centralWidget)
        self.tabWidget.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.tabWidget.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.tabWidget.setObjectName(_fromUtf8("tabWidget"))
        self.tab = QtGui.QWidget()
        self.tab.setObjectName(_fromUtf8("tab"))
        self.gridLayout_2 = QtGui.QGridLayout(self.tab)
        self.gridLayout_2.setMargin(11)
        self.gridLayout_2.setSpacing(6)
        self.gridLayout_2.setObjectName(_fromUtf8("gridLayout_2"))
        self.frame_2 = QtGui.QFrame(self.tab)
        self.frame_2.setFrameShape(QtGui.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtGui.QFrame.Raised)
        self.frame_2.setObjectName(_fromUtf8("frame_2"))
        self.gridLayout_6 = QtGui.QGridLayout(self.frame_2)
        self.gridLayout_6.setMargin(11)
        self.gridLayout_6.setSpacing(6)
        self.gridLayout_6.setObjectName(_fromUtf8("gridLayout_6"))
        self.Dialog_title = QtGui.QLabel(self.frame_2)
        self.Dialog_title.setStyleSheet(_fromUtf8("font: italic 26pt \"Monotype Corsiva\";"))
        self.Dialog_title.setObjectName(_fromUtf8("Dialog_title"))
        self.gridLayout_6.addWidget(self.Dialog_title, 0, 1, 1, 1)
        spacerItem = QtGui.QSpacerItem(40, 20, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        self.gridLayout_6.addItem(spacerItem, 0, 0, 1, 1)
        spacerItem1 = QtGui.QSpacerItem(40, 20, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        self.gridLayout_6.addItem(spacerItem1, 0, 2, 1, 1)
        self.gridLayout_2.addWidget(self.frame_2, 0, 2, 1, 1)
        self.frame = QtGui.QFrame(self.tab)
        self.frame.setStyleSheet(_fromUtf8("font: italic 22pt \"Monotype Corsiva\";"))
        self.frame.setFrameShape(QtGui.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtGui.QFrame.Raised)
        self.frame.setObjectName(_fromUtf8("frame"))
        self.gridLayout_3 = QtGui.QGridLayout(self.frame)
        self.gridLayout_3.setMargin(11)
        self.gridLayout_3.setSpacing(6)
        self.gridLayout_3.setObjectName(_fromUtf8("gridLayout_3"))
        self.label_2 = QtGui.QLabel(self.frame)
        self.label_2.setObjectName(_fromUtf8("label_2"))
        self.gridLayout_3.addWidget(self.label_2, 3, 0, 1, 1)
        self.label_5 = QtGui.QLabel(self.frame)
        self.label_5.setObjectName(_fromUtf8("label_5"))
        self.gridLayout_3.addWidget(self.label_5, 3, 2, 1, 1)
        self.label_3 = QtGui.QLabel(self.frame)
        self.label_3.setObjectName(_fromUtf8("label_3"))
        self.gridLayout_3.addWidget(self.label_3, 6, 0, 1, 1)
        self.pushButton_3 = QtGui.QPushButton(self.frame)
        self.pushButton_3.setObjectName(_fromUtf8("pushButton_3"))
        self.gridLayout_3.addWidget(self.pushButton_3, 7, 1, 1, 1)
        self.label = QtGui.QLabel(self.frame)
        self.label.setObjectName(_fromUtf8("label"))
        self.gridLayout_3.addWidget(self.label, 1, 0, 1, 1)
        self.lineEdit = QtGui.QLineEdit(self.frame)
        self.lineEdit.setObjectName(_fromUtf8("lineEdit"))
        self.gridLayout_3.addWidget(self.lineEdit, 6, 1, 1, 2)
        self.open_result_file = QtGui.QCheckBox(self.frame)
        self.open_result_file.setObjectName(_fromUtf8("open_result_file"))
        self.gridLayout_3.addWidget(self.open_result_file, 8, 1, 1, 2)
        self.pushButton_2 = QtGui.QPushButton(self.frame)
        self.pushButton_2.setMaximumSize(QtCore.QSize(193, 44))
        self.pushButton_2.setToolTip(_fromUtf8(""))
        self.pushButton_2.setWhatsThis(_fromUtf8(""))
        self.pushButton_2.setObjectName(_fromUtf8("pushButton_2"))
        self.gridLayout_3.addWidget(self.pushButton_2, 3, 1, 1, 1)
        self.label_4 = QtGui.QLabel(self.frame)
        self.label_4.setObjectName(_fromUtf8("label_4"))
        self.gridLayout_3.addWidget(self.label_4, 1, 2, 1, 1)
        self.label_6 = QtGui.QLabel(self.frame)
        self.label_6.setObjectName(_fromUtf8("label_6"))
        self.gridLayout_3.addWidget(self.label_6, 9, 1, 1, 1)
        self.pushButton = QtGui.QPushButton(self.frame)
        self.pushButton.setMaximumSize(QtCore.QSize(193, 44))
        self.pushButton.setToolTip(_fromUtf8(""))
        self.pushButton.setObjectName(_fromUtf8("pushButton"))
        self.gridLayout_3.addWidget(self.pushButton, 1, 1, 1, 1)
        self.label_7 = QtGui.QLabel(self.frame)
        self.label_7.setStyleSheet(_fromUtf8("font: italic 8pt \"MS Shell Dlg 2\";"))
        self.label_7.setObjectName(_fromUtf8("label_7"))
        self.gridLayout_3.addWidget(self.label_7, 2, 1, 1, 1)
        self.groupBox = QtGui.QGroupBox(self.frame)
        self.groupBox.setObjectName(_fromUtf8("groupBox"))
        self.gridLayout_5 = QtGui.QGridLayout(self.groupBox)
        self.gridLayout_5.setMargin(11)
        self.gridLayout_5.setSpacing(6)
        self.gridLayout_5.setObjectName(_fromUtf8("gridLayout_5"))
        self.radioButton_2 = QtGui.QRadioButton(self.groupBox)
        self.radioButton_2.setObjectName(_fromUtf8("radioButton_2"))
        self.gridLayout_5.addWidget(self.radioButton_2, 1, 0, 1, 1)
        self.radioButton = QtGui.QRadioButton(self.groupBox)
        self.radioButton.setObjectName(_fromUtf8("radioButton"))
        self.gridLayout_5.addWidget(self.radioButton, 0, 0, 1, 1)
        self.radioButton_3 = QtGui.QRadioButton(self.groupBox)
        self.radioButton_3.setObjectName(_fromUtf8("radioButton_3"))
        self.gridLayout_5.addWidget(self.radioButton_3, 2, 0, 1, 1)
        self.gridLayout_3.addWidget(self.groupBox, 5, 0, 1, 3)
        self.label_8 = QtGui.QLabel(self.frame)
        self.label_8.setStyleSheet(_fromUtf8("font: italic 8pt \"MS Shell Dlg 2\";"))
        self.label_8.setObjectName(_fromUtf8("label_8"))
        self.gridLayout_3.addWidget(self.label_8, 4, 1, 1, 1)
        self.gridLayout_2.addWidget(self.frame, 2, 2, 1, 1)
        spacerItem2 = QtGui.QSpacerItem(40, 20, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        self.gridLayout_2.addItem(spacerItem2, 0, 0, 1, 1)
        spacerItem3 = QtGui.QSpacerItem(40, 20, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        self.gridLayout_2.addItem(spacerItem3, 2, 3, 1, 1)
        spacerItem4 = QtGui.QSpacerItem(20, 40, QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Expanding)
        self.gridLayout_2.addItem(spacerItem4, 1, 2, 1, 1)
        spacerItem5 = QtGui.QSpacerItem(40, 20, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        self.gridLayout_2.addItem(spacerItem5, 2, 0, 1, 1)
        spacerItem6 = QtGui.QSpacerItem(20, 40, QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Expanding)
        self.gridLayout_2.addItem(spacerItem6, 3, 2, 1, 1)
        self.tabWidget.addTab(self.tab, _fromUtf8(""))
        self.tab_3 = QtGui.QWidget()
        self.tab_3.setObjectName(_fromUtf8("tab_3"))
        self.tabWidget.addTab(self.tab_3, _fromUtf8(""))
        self.tab_2 = QtGui.QWidget()
        self.tab_2.setObjectName(_fromUtf8("tab_2"))
        self.gridLayout_4 = QtGui.QGridLayout(self.tab_2)
        self.gridLayout_4.setMargin(11)
        self.gridLayout_4.setSpacing(6)
        self.gridLayout_4.setObjectName(_fromUtf8("gridLayout_4"))
        self.tableWidget = QtGui.QTableWidget(self.tab_2)
        self.tableWidget.setObjectName(_fromUtf8("tableWidget"))
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.gridLayout_4.addWidget(self.tableWidget, 0, 0, 1, 1)
        self.tabWidget.addTab(self.tab_2, _fromUtf8(""))
        self.gridLayout.addWidget(self.tabWidget, 0, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralWidget)
        self.menuBar = QtGui.QMenuBar(MainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 826, 21))
        self.menuBar.setObjectName(_fromUtf8("menuBar"))
        self.menu_Help = QtGui.QMenu(self.menuBar)
        self.menu_Help.setObjectName(_fromUtf8("menu_Help"))
        MainWindow.setMenuBar(self.menuBar)
        self.mainToolBar = QtGui.QToolBar(MainWindow)
        self.mainToolBar.setObjectName(_fromUtf8("mainToolBar"))
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.mainToolBar)
        self.statusBar = QtGui.QStatusBar(MainWindow)
        self.statusBar.setObjectName(_fromUtf8("statusBar"))
        MainWindow.setStatusBar(self.statusBar)
        self.action_Rules = QtGui.QAction(MainWindow)
        self.action_Rules.setObjectName(_fromUtf8("action_Rules"))
        self.menu_Help.addAction(self.action_Rules)
        self.menuBar.addAction(self.menu_Help.menuAction())

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(2)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(_translate("MainWindow", "Libary system", None))
        self.Dialog_title.setText(_translate("MainWindow", "Library Stock Managment System", None))
        self.label_2.setText(_translate("MainWindow", "Slave file", None))
        self.label_5.setText(_translate("MainWindow", "file name", None))
        self.label_3.setText(_translate("MainWindow", "Type Result File Name", None))
        self.pushButton_3.setText(_translate("MainWindow", "Generate Report", None))
        self.label.setText(_translate("MainWindow", "Masterfile", None))
        self.open_result_file.setText(_translate("MainWindow", "open the result file", None))
        self.pushButton_2.setText(_translate("MainWindow", "Browes", None))
        self.label_4.setText(_translate("MainWindow", "<html><head/><body><p>file name</p></body></html>", None))
        self.label_6.setText(_translate("MainWindow", "Result of file", None))
        self.pushButton.setText(_translate("MainWindow", "Browes", None))
        self.label_7.setText(_translate("MainWindow", "[ click on browes button to select master file  ]", None))
        self.groupBox.setTitle(_translate("MainWindow", "Select the report file to be generated", None))
        self.radioButton_2.setText(_translate("MainWindow", "Find same Accession Numbers", None))
        self.radioButton.setText(_translate("MainWindow", "Find Missing Accession Numbers", None))
        self.radioButton_3.setText(_translate("MainWindow", "Remove same Accession from Master file", None))
        self.label_8.setText(_translate("MainWindow", "[ click on browes button to select slave file  ]", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "Main Lib", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), _translate("MainWindow", "Dept.Lib", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "View Report", None))
        self.menu_Help.setTitle(_translate("MainWindow", "&Help", None))
        self.action_Rules.setText(_translate("MainWindow", "&Rules", None))

