'''
Created on 03-Jun-2016

@author: jon
'''
from sys import getsizeof
import sys

try:

    from PyQt4.QtGui import QMainWindow, QApplication, QWidget, QFileDialog, QMessageBox
    from PyQt4 import QtCore

    PyQt_version = 4
except ImportError:
    from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QFileDialog, QMessageBox

    PyQt_version = 5
from view.mainwindow import *
from logic.logic import *

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

class ThreadLib(QtCore.QThread):
    def __init__(self):
        QtCore.QThread.__init__(self)
        
        

    def run(self):
        self.start()
       
        

    def __del__(self):
    	self.exiting = True
    	self.wait()

class MainForm(QMainWindow):

    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.logic = Main()
        self.uiWindow = Ui_MainWindow()
        self.uiWindow.setupUi(self)
        # self.uiWindow.progressBar.setMinimum(24)
        # self.uiWindow.progressBar.setMaximum(90)
        self.uiWindow.tableWidget.setRowCount(5)
        self.uiWindow.tableWidget.setColumnCount(3)
        self.save = None

        self.uiWindow.open_result_file.hide()

        if PyQt_version == 4:
           
                QtCore.QObject.connect(self.uiWindow.pushButton, QtCore.SIGNAL(_fromUtf8("clicked()")), self.dialog_show)
                QtCore.QObject.connect(self.uiWindow.pushButton_2, QtCore.SIGNAL(_fromUtf8("clicked()")), self.dialog_show_2)
                QtCore.QObject.connect(self.uiWindow.pushButton_3, QtCore.SIGNAL(_fromUtf8("clicked()")), self.save_file_on_gen)
        else:
            self.uiWindow.pushButton.clicked.connect(self.dialog_show)
            self.uiWindow.pushButton_2.clicked.connect(self.dialog_show_2)

    def dialog_show(self):

        dia = QFileDialog.getOpenFileName(self, _fromUtf8("open excel"), directory=__file__)
        self.uiWindow.pushButton.setStyleSheet("color:green")

            
        #print(dia)
        self.uiWindow.label_4.setText((dia))
        rp = (self.logic.ok_text_1())
        self.uiWindow.pushButton.setText(rp)
        


        if (rp == "fails"):
            self.uiWindow.pushButton.setStyleSheet("color:red")
        try:
            try:
               thread = ThreadLib()
               self.connect(thread, QtCore.SIGNAL("thread.start()"), self.fack_function_1)	
               

               print(getsizeof(dia))
               if (dia):
               	 del dia

            except RuntimeError as e:
            	print(e)
        except FileNotFoundError:
                self.uiWindow.pushButton.setStyleSheet("color:black")
                self.uiWindow.label_4.setText("Browes")
                if (dia):
                	del dia

    def fack_function_1(self):
        self.logic.load_file_1(dia)
   
    def dialog_show_2(self):


        try:
            
            dia = QFileDialog.getOpenFileName(self, _fromUtf8("open excel"), directory=__file__)
        except FileNotFoundError:
           self.uiWindow.pushButton.setStyleSheet("color:black")
            
        self.uiWindow.label_5.setText(dia)
        rp = self.logic.ok_text_2()
        self.uiWindow.pushButton_2.setText(rp)
        self.uiWindow.pushButton_2.setStyleSheet("color:green")
        if (rp == "fails"):
            self.uiWindow.pushButton_2.setStyleSheet("color:red")

        
        try:
            try:
               thread = ThreadLib()
               self.connect(thread, QtCore.SIGNAL("thread.start()"), self.fack_function_2)	
               
               if (dia):
               	 del dia
            except  RuntimeError as e:
            	print(e)
        except FileNotFoundError:
                self.uiWindow.pushButton.setStyleSheet("color:black")
                self.uiWindow.label_4.setText("Browes")
                if (dia):
                	del dia
    def fack_function_2(self):
        self.logic.load_file_2(dia)

    def fack_function_in_save(self):
    	self.logic.save_file(self.save)

    def save_file_on_gen(self):
        try:
            self.save = self.uiWindow.lineEdit.text()+".xlsx"
            try:
               thread = ThreadLib()
               self.connect(thread, QtCore.SIGNAL("thread.start()"), self.fack_function_in_save)	
               

            except  RuntimeError as e:
            	print(e)
            
            self.error = "success"
            self.uiWindow.open_result_file.show()
          #self.uiWindow
            QtCore.QObject.connect(self.uiWindow.open_result_file, QtCore.SIGNAL(_fromUtf8("pressed()")), self.__pop_file)

        except Exception as e:
            self.error = "fails"
            print(e.__traceback__())

        finally:
             self.__j()

    def __j(self):
        self.uiWindow.label_6.setText(self.error)
        self.uiWindow.label_6.setStyleSheet("color:green")

        if self.error == "fails":
          self.uiWindow.label_6.setStyleSheet("color:red")

    def __pop_file(self):
          from os import popen
          popen(self.save)
    
    def __delect_store(self, dele):
    	'''
    	:param dele: this is a variable to be delected.
    	'''
    	del dele


if '__main__' == __name__:
    app = QApplication(sys.argv)
    main = MainForm()
    main.show()
    sys.exit(app.exec_())
