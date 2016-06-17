'''
Created on 03-Jun-2016

@author: jon
'''
from sys import getsizeof
import sys
from time import sleep
try:

    from PyQt4.QtGui import QMainWindow, QApplication, QWidget, QFileDialog, QMessageBox, QButtonGroup
    from PyQt4 import QtCore

    PyQt_version = 4
except ImportError:
    from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QFileDialog, QMessageBox, QButtonGroup

    PyQt_version = 5

from logic.logic import *

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

import  openpyxl.utils.exceptions
from PyQt4.QtGui import QMessageBox



class MainForm(QMainWindow):

    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.logic = Main()
        self.uiWindow = self.logic.ui_window
        self.uiWindow.setupUi(self)
        # self.uiWindow.progressBar.setMinimum(24)
        # self.uiWindow.progressBar.setMaximum(90)
        self.uiWindow.tableWidget.setRowCount(5)
        self.uiWindow.tableWidget.setColumnCount(3)
        self.__count_1 = False
        self.__count_2 = False
        self.__select_radio = self.logic.missing_check()
        self.save = None
        self.dia = None
        self.uiWindow.open_result_file.hide()
        self.__radio_indication = open("read.txt",'w')



        if PyQt_version == 4:
           
                QtCore.QObject.connect(self.uiWindow.pushButton, QtCore.SIGNAL(_fromUtf8("clicked()")), self.dialog_show)
                QtCore.QObject.connect(self.uiWindow.pushButton_2, QtCore.SIGNAL(_fromUtf8("clicked()")), self.dialog_show_2)
                QtCore.QObject.connect(self.uiWindow.pushButton_3, QtCore.SIGNAL(_fromUtf8("clicked()")), self.save_file_on_gen)
                QtCore.QObject.connect(self.uiWindow.radioButton, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.p)
                QtCore.QObject.connect(self.uiWindow.radioButton_2, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.p)
                QtCore.QObject.connect(self.uiWindow.radioButton_3, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.p)
                QtCore.QObject.connect(self.uiWindow.open_result_file, QtCore.SIGNAL(_fromUtf8("pressed()")), self.__pop_file)
        else:
            self.uiWindow.pushButton.clicked.connect(self.dialog_show)
            self.uiWindow.pushButton_2.clicked.connect(self.dialog_show_2)
            self.uiWindow.open_result_file.pressed.connect( self.__pop_file)
    
    def p(self,e):
       if (self.uiWindow.radioButton.isChecked() or self.uiWindow.radioButton_3.isChecked()):
         self.__radio_indication.write(1)
       else:
       	 self.__radio_indication.write(2)




    def dialog_show(self):
        """
        :self: this inherted of class
        this method is used for getting the file name and passed into the logic class for selecting the
        list.
        :return:
        """
        del self.dia
        if(self.__count_1):
          try:
            self.logic.master_set.remove()
          except TypeError:
           pass

        self.__count_1 = True
        self.dia = QFileDialog.getOpenFileName(self, _fromUtf8("open excel"), directory=__file__, filter = "Excel Files(*.xlsx *.xlsm *.xltx *.xltm)" )
        self.uiWindow.pushButton.setStyleSheet("color:green")

            
        #print(dia)
        self.uiWindow.label_4.setText((self.dia))
        rp = (self.logic.ok_text_1())
        self.uiWindow.pushButton.setText(rp)
        


        if (rp == "fails"):
            self.uiWindow.pushButton.setStyleSheet("color:red")
        try:
            try:
               
               #self.connect(thread, QtCore.SIGNAL("thread.start()"), self.fack_function_1)	
               self.fack_to_load_file_1()
               
               print("Master's size:",getsizeof(self.dia), type(self.dia))
               if (self.dia):

                 del self.dia
                 self.dia = None
            except  openpyxl.utils.exceptions.InvalidFileException as e:
                #openpyxl.utils.exceptions.InvalidFileException
                print(e)
                self.uiWindow.pushButton.setText(rp)


        except FileNotFoundError:
                self.uiWindow.pushButton.setStyleSheet("color:black")
                self.uiWindow.pushButton.setText("Browse")
                self.uiWindow.label_4.setText("Browse")
                if (self.dia):
                  del self.dia


    def fack_to_load_file_1(self):
       self.logic.load_file_1(self.dia)


    def dialog_show_2(self):


        """

        :type self: object
        """
        del self.dia
        if(self.__count_2):
            
            try:
               self.logic.slave_set.remove()
            except TypeError:
               pass

        self.__count_2 = True
        try:
            
            self.dia = QFileDialog.getOpenFileName(self, _fromUtf8("open slave excel"), directory=__file__, filter = "Excel Files(*.xlsx *.xlsm *.xltx *.xltm)" ) #
        except FileNotFoundError:
           self.uiWindow.pushButton_2.setStyleSheet("color:black")
            
        self.uiWindow.label_5.setText(self.dia)
        rp = self.logic.ok_text_2()
        self.uiWindow.pushButton_2.setText(rp)
        self.uiWindow.pushButton_2.setStyleSheet("color:green")
        if (rp == "fails"):
            self.uiWindow.pushButton_2.setStyleSheet("color:red")

        
        try:
            try:
              
               
               self.fack_to_load_file_2()
               print("slave size:",getsizeof(self.dia), len(self.dia))
               if (self.dia):

                  del self.dia
                  self.dia = None
            except  openpyxl.utils.exceptions.InvalidFileException as e:

                print(e)
                QMessageBox.information(None, "version", _fromUtf8("invalid..."))
                self.uiWindow.label_5.setText("Invalid format")

        except FileNotFoundError:
                self.uiWindow.pushButton_2.setStyleSheet("color:black")
                self.uiWindow.pushButton_2.setText("Browse")
                self.uiWindow.label_5.setText("Browse")
                if (self.dia):

                  del self.dia
                  self.dia = None

    def fack_to_load_file_2(self):
     self.logic.load_file_2(self.dia)

    def save_file_on_gen(self):
        try:

            self.dia = QFileDialog.getSaveFileName(self, _fromUtf8("Save Excel"), directory=__file__, filter = "Excel Files(*.xlsx *.xlsm *.xltx *.xltm)" ) #
            self.uiWindow.tableWidget.clear()
        except FileNotFoundError:
           self.uiWindow.pushButton_2.setStyleSheet("color:black")
        print(type(self.dia), self.dia)
        try:
            self.save = self.dia
            self.logic.save_file(self.save)
            self.error = "success"
            self.uiWindow.open_result_file.show()
            self.__j()
          #self.uiWindow


        except FileNotFoundError:
             self.error = "fails"
             self.uiWindow.label_6.setText("File not found")

    def __j(self):
        self.uiWindow.label_6.setText(self.error)

        if self.error == "fails":
          self.uiWindow.label_6.setStyleSheet("color:red")

    def __pop_file(self):
          from os import popen
          popen(self.save)

    def __del__(self):
          from os import popen
          try:
            popen("del read.txt")
          except Exception:
           	 pass


if '__main__' == __name__:
    app = QApplication(sys.argv)
    main = MainForm()
    main.show()
    sys.exit(app.exec_())
