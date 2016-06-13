'''
Created on 03-Jun-2016

@author: jon
'''

from openpyxl import Workbook
from openpyxl import load_workbook

from PyQt4.QtGui import QMessageBox, QMainWindow, QTableWidgetItem , QInputDialog, QWidget, QRadioButton
from PyQt4 import QtCore

from view.mainwindow import *

try:
    _fromUtf8 = QtCore.QString.fromUtf8
    QString = QtCore.QString
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


#################################################################################
#
# maseter unique cell must be set carfully
# :variable:::maseter_head
#
#################################################################################
class Main(QMainWindow):
    '''
    classdocs
    '''

    def __init__(self, parent=None):
        '''
        Constructor
        '''
        QMainWindow.__init__(self, parent)
        self.load_wb = None
        self.load_wb2 = None
        self.select = None
        self.select2 = None
        self.master_set = list()
        self.slave_set = list()
        self.len_row = 0
        self.len_row2 = 0
        self.len_col = 0
        self.len_col2 = 0
        self.row_1 = 2
        self.row_data = None
        self.len_col = None
        self.len_row_str = " "
        self.len_row2_str = " "
        self.len_col_str = " "
        self.master_header = "AccessNo"
        self.len_col2_str = " "
        self.ok_2 = "Success"
        self.ok_1 = "Success"
        self.store_set = list()
        self.finding_missing_feild = list()
        self.j_loop = 1
        self.col_i = 1
        # self.Range = None
        self.ui_window = Ui_MainWindow()
        self.ui_window.setupUi(self)
        self.item = QTableWidgetItem()
        self.ui_window.groupBox.setCheckable(True)

        self.wb = Workbook()

        self.wb_sheet1_1 = self.wb.active
        QtCore.QObject.connect(self.ui_window.radioButton_2, QtCore.SIGNAL(_fromUtf8("toggled(bool)")), self.p)

        # self.select2['A1:A'+self.len_row2]

        # row_count_slave = 1
        # row_count_master = 1

        # self.search_in_master()
    def p(self,e):
       print(83405)

    def load_file_1(self, a):
        """
        this file is used to load file Master excel file
        :param a: file name
        :return: cuts the file and take the value in Master excel file
        """

        # a = "Master.xlsx"
        self.load_wb = load_workbook(a)

        try:
            self.__u = self.load_wb.get_sheet_names()[0]

            self.select = self.load_wb[self.__u]

        except KeyError:
            try:
                self.select = self.load_wb["Sheet"]
            except KeyError:

                self.select = self.load_wb[_fromUtf8(PopWidget().text_r())]
            else:
                self.ok_1 = "fails"

        self.len_row = (len(self.select.rows))
        self.len_row_str = str(self.len_row)
        # start
        self.len_col = len(self.select.columns)
        self.len_col_str = str(self.len_col)
        ####print(("No of columns in Master"+self.len_col)
        # end
        self.l = (self.select['A1:A' + self.len_row_str])
        count = 1
        for j in self.l:
            self.master_set.append(self.select['A' + str(count)].value)
            ####print((self.select['A'+str(count)].value)
            count += 1

    def load_file_2(self, a):
        """
        this file is used to load file slave excel file
        :param a: file name
        :return: cuts the file and take the value in slave excel file
        """

        # a = "Slave1.xlsx"
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

        self.len_row2 = len(self.select2.rows)
        self.len_row2_str = str(self.len_row2)
        count = 1
        for slave in self.select2['A1:A' + self.len_row2_str]:
            self.slave_set.append(self.select2['A' + str(count)].value)
            ####print((">>>>slave",(self.select2['A'+str(count)].value))
            count += 1
            # self.store_set = list(0)

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

        try:

            #    self.store.remove(None)
            #  self.store.remove(self.master_header)
            ##print((self.store)
            self.slave_set.remove(None)
            self.slave_set.remove(self.master_header)
        except ValueError:
            pass

        # self.store.sort()
        # #print((self.store)

        #sorted(self.slave_set)


        ###print((self.store,"\nlen of store:",len_store)

        # for i in range((self.select["A1:H1"])):
        # self.my_range = self.Range('Sheet1','A1:H1').value
        # self.Range('Sheet2','A1:H1').value = self.my_range
        """
        row_list = []                  # start
        for i in range(1,self.len_col):
           row_data = self.row_values(i)
           row_list.append(row_data)

                                       # end
        """

        self.master_set = list(self.master_set)

        # print((self.master_set)
        try:

            self.master_set.remove(None)
            self.master_set.remove(self.master_header)
        except ValueError as v:
            print("V-->\n", v)

        self.len_row = len(self.master_set)
        # print(("master set", self.len_row, "\n row len of file1:", self.len_row, '\n', "====" * 10)

        try:
            self.master_set.sort()
            sorted(self.store_set)
        except TypeError:
            # print(self.master_set,"slave sort:",self.slave_set)
            pass

            # print(("master set", "--" * 20, len(self.master_set), "\n", "--" * 20, "\n len of slave", len(self.slave_set))

            # <-----------------------dont delect this
        try:

            for i in self.select["A1:I1"]:
                """
            --- For inserting header on result excel from master excel
             """
                coutn = 1

                for j in i:
                    self.wb_sheet1_1.cell(row=1, column=coutn, value=j.value)
                    coutn += 1
        except TypeError as t:
            print("T-----------------\n", t)

        i = 1 + 2
        j = []
        self.col_i = 0

        # print(self.select.rows, self.select.columns)
        self.ui_window.tableWidget.setRowCount(len(self.select.rows))
        print((type(self.len_row),self.len_row, "==",len(self.select.rows)))
        enable = False
        self.ui_window.tableWidget.setColumnCount(len(self.select.columns))

        g = 0
        c= 0
        while(g < self.len_row):
            """"
            :: self.col_i == counting slave loop
            """
            c +=1
            toS = str(c)


            #print(("++",self.select["A"+toS+":I"+toS])
            # for j in c:
            #print(("rotation of store:",j)
            """
                  if ( list_value_2 == list_value ):
                ###print((self.master_set[j],"and", self.store[j])
                to = self.select["A"+toS+":I"+toS]
                ###print((to ,"\n rotation:",self.col_i)
                self.col_i+=1

                for r_o in to:
                   col_1 = 1
                   for c_o in r_o:
                      self.wb_sheet1_1.cell(row=self.row_1, column=col_1,value= c_o.value)
                      ###print(("the instered values-->", c_o.value)
                      col_1+=1
                   self.row_1+=1

                 """

            try:

                a1 = self.slave_set[self.col_i]
            except IndexError:
                print("Index error ====G: ", g, "self.col_i:", self.col_i)

            if(enable == True):
               g = g - 1
            b1 = (self.master_set[g])
            g+=1
            if(self.len_row == self.col_i):
                break
            if (a1 == b1):
              print(   "master:", b1, "==", a1,":slave" )
              self.compare_to_check(str(g))
              print("index for master file:",g,"index for slave file",self.col_i,"(toS):", toS)
              self.col_i += 1
            else:
                print(" \t \t \t \t \t \t \t slave:", a1, "!=", b1, ":master")
                enable = False

                if(a1 < b1):
                  print("\t \t \t \t \t \t \t  addition values in slave:", a1)
                  self.col_i += 1
                  enable = True


            '''
            if (a1 != b1):
                print("slave:", a1, "!=", b1, ":master")

                self.compare_to_check(toS)

            else:
                print("\t \t \t \t \t \t \t slave:", a1, "==", b1, ":master")
                self.col_i += 1
            

                # print(("[",self.col_i,"]",c)
            '''

        """
        [
          -- a test for the inserting values
        ]

        for i in range(0, len_store):
            run = str(i+10)
            self.wb_sheet1_1["A"+run] = self.store[i]
        #a = "h.xlsx"
         """""

        print(self.row_1)
        print("cure",self.col_i)
        try:
            self.wb.save(a)

            # print(()
        except PermissionError:
            # #print(("close the report file please..")
            '''
            from os import system
            system("taskkill /IM excel.exe")
            '''
            QMessageBox.information(None, "version", _fromUtf8("closing result  files\n writeing is not allowed on opened file ..."))#
            self.wb.save(a)


    def missing_check(self, a1, b1, toS):

        if (a1 != b1):
            print("slave:", a1, "!=", b1, ":master")
            self.compare_to_check(toS)
        else:
                self.col_i += 1

    def same_check(self, a1, b1, toS):
        if (a1 == b1):
            print("slave:", a1, "==", b1, ":master")

            self.compare_to_check(toS)
            self.col_i += 1

    def compare_to_check(self, toS):

        # #print(("first--->%d )( second --->%d",a,b)
        """

        :type self: object
        :type toS: int(value) to string
        """
        to = self.select["A" + toS + ":I" + toS]
        # print(("row count in master:",toS)
        # print((to ,"\n rotation:",self.col_i)


        for r_o in to:
            col_1 = 1
            #print(("row---->",self.row_1,"\ncol----------------------------------------------------")
            for c_o in r_o:
                #print((col_1,end="\t")
                ''''
                   add the QTableWidgetItem() 
                   for correct code 
                '''

                """
                try:
                    try:
                      c_ovalue1 = QTableWidgetItem(str(c_o.value))
                      self.ui_window.tableWidget.setItem(self.row_1, col_1-1, QTableWidgetItem(c_ovalue1))
                    except ValueError:
                      self.ui_window.tableWidget.setItem(self.row_1, col_1-1, QTableWidgetItem((c_ovalue1)))
                except (TypeError):
                    print("...........")
                    self.ui_window.tableWidget.setItem(self.row_1, col_1-1, QTableWidgetItem(QtCore.QDate.toString(c_o.value)))
                else:
                    self.ui_window.tableWidget.setItem(self.row_1, col_1-1, QTableWidgetItem((c_o.value)))

                """

                self.wb_sheet1_1.cell(row=self.row_1, column=col_1, value=c_o.value)
                ##print((c_o.value)
                col_1 += 1
            ###print(()
            self.row_1 += 1
