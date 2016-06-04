'''
Created on 03-Jun-2016

@author: kar
'''

from openpyxl import Workbook
from openpyxl import load_workbook

from PyQt4.QtGui import QMessageBox
from PyQt4.QtGui import QInputDialog,QWidget
from PyQt4 import QtCore


try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

class PopWidget(QWidget):
    def __init__(self):
        QWidget.__init__(self)
        self.text, self.ok = QInputDialog.getText(self,
                        "Name of the sheet",
                "Enter the correct name of current sheet to access",
                           )

    def text_r(self):
            if self.ok:
              return self.text




class Main(object):
    '''
    classdocs
    '''


    def __init__(self):
        '''
        Constructor
        '''
        self.load_wb = None
        self.load_wb2 = None
        self.select = None
        self.select2 = None
        self.master_set = set()
        self.len_row = None
        self.len_row2 = None
        self.row_data = None
        self.len_col = None
        self.ok_2 = "Success"
        self.ok_1 = "Success"
        #self.Range = None
        self.wb = Workbook()

        self.wb_sheet1_1 = self.wb.active

        #self.select2['A1:A'+self.len_row2]
        self.slave_set = set()
        #row_count_slave = 1
        #row_count_master = 1

       # self.search_in_master()

    def load_file_1(self, a):
        """
        this file is used to load file Master excel file
        :param a: file name
        :return: cuts the file and take the value in Master excel file
        """

            #a = "Master.xlsx"
        self.load_wb = load_workbook(a)


        try:
          self.__u = self.load_wb.get_sheet_names()[0]
          self.select = self.load_wb[self.__u]
        except KeyError:
            try:
              self.select2 = self.load_wb["Sheet"]
            except KeyError:

              self.select2 = self.load_wb[_fromUtf8(PopWidget().text_r())]
            else:
                self.ok_1 = "fails"

        self.len_row = str(len(self.select.rows))
        #start
        self.len_col = str(len(self.select.columns))
        #print("No of columns in Master"+self.len_col)
        #end
        self.l = (self.select['A1:A'+self.len_row])
        count = 1
        for j in self.l:
            self.master_set.add(self.select['A'+str(count)].value)
            #print(self.select['A'+str(count)].value)
            count+=1

    def load_file_2(self, a):
        """
        this file is used to load file slave excel file
        :param a: file name
        :return: cuts the file and take the value in slave excel file
        """

        #a = "Slave1.xlsx"
        self.load_wb2 = load_workbook(a)

        try:
          self.select2 = self.load_wb2["Sheet1"]
        except KeyError:
            try:
              self.select2 = self.load_wb2["Sheet"]
            except KeyError:

              self.select2 = self.load_wb2[_fromUtf8(PopWidget().text_r())]
            else:
                self.ok_2 = "fails"


        self.len_row2 = str(len(self.select2.rows))
        count=1
        for slave in self.select2['A1:A'+self.len_row2]:
            self.slave_set.add(self.select2['A'+str(count)].value)
            #print(">>>>slave",(self.select2['A'+str(count)].value))
            count+=1
    def ok_text_1(self):
        return self.ok_1

    def ok_text_2(self):
        return self.ok_2

    def save_file(self, a):
        """
        This function is used to save the file in .xlsx format
        :param a: file name
        :return: save the excel file with the given name
        """
        self.store = list(self.master_set.difference(self.slave_set))


        #for i in range((self.select["A1:H1"])):
        #self.my_range = self.Range('Sheet1','A1:H1').value
        #self.Range('Sheet2','A1:H1').value = self.my_range
        """
        row_list = []                  # start
        for i in range(1,self.len_col):
           row_data = self.row_values(i)
           row_list.append(row_data)

                                       # end
        """

        for i in self.select["A1:H1"]:
            """ 
            <<<...for inserting header on result excel from master excel 
         """
            coutn =1
            for j in i:
                self.wb_sheet1_1.cell(row=1, column=coutn,value= j.value)
                coutn+=1
        for i in range(0, len(self.store)):
            self.wb_sheet1_1["A"+str(i+2)] =  self.store[i]
        #a = "h.xlsx"

        try:

          self.wb.save(a)


        except PermissionError :
            #print("close the report file please..")
            QMessageBox.information(None, "version", _fromUtf8("Please close the result..."))



