# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\Final_App2_v2.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import pyqtSignal, QEvent
from PyQt5.QtWidgets import QLineEdit
from pathlib import Path

import pandas as pd
import numpy as np
import math
from scipy import signal
import xlsxwriter

class initVal:
    def __init__(self):
        self._Freq = []
        self._W = []
        self._V = []
        self._Area = 1 * 1e-3
        self.df1 = pd.DataFrame()

class initParameter:
    def __init__(self):
        self.k = 1.38 * 1e-23
        self.q = 1.60 * 1e-19
        self.sn = 1.00 * 1e-14  # sigma (σ)
        self.nth = 1.90 * 1e+7
        self.Nc = 1.9 * 1e+19
        self.T = 303

global init         # for initVal
global initPara     # for initParameter
initPara = initParameter()

def indices(lst, item):
    return [i for i, x in enumerate(lst) if x == item]

class Change_file(QLineEdit):
    doubleClicked = pyqtSignal()

    def event(self, event):
        if event.type() == QEvent.Type.MouseButtonDblClick:
            self.doubleClicked.emit()
        return super().event(event)

class Change_parameter(QLineEdit):
    doubleClicked = pyqtSignal()

    def event(self, event):
        if event.type() == QEvent.Type.MouseButtonDblClick:
            self.doubleClicked.emit()
        return super().event(event)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(758, 275)
        font = QtGui.QFont()
        font.setPointSize(10)
        MainWindow.setFont(font)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_5.addWidget(self.label_7)
        self.txt_default_file = QtWidgets.QLineEdit(self.centralwidget)
        self.txt_default_file.setEnabled(False)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.txt_default_file.setFont(font)
        self.txt_default_file.setStatusTip("")
        self.txt_default_file.setObjectName("txt_default_file")
        self.horizontalLayout_5.addWidget(self.txt_default_file)

        # Bind doubleClicked on QlineEdit
        self.txt_new_default_file = Change_file(self.txt_default_file)
        self.txt_new_default_file.doubleClicked.connect(self.change_default)

        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")
        self.horizontalLayout_5.addWidget(self.label_8)
        self.txt_deault_T = QtWidgets.QLineEdit(self.centralwidget)
        self.txt_deault_T.setEnabled(True)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.txt_deault_T.setFont(font)
        self.txt_deault_T.setMouseTracking(True)
        self.txt_deault_T.setObjectName("txt_deault_T")
        self.horizontalLayout_5.addWidget(self.txt_deault_T)
        self.verticalLayout.addLayout(self.horizontalLayout_5)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(False)
        font.setWeight(50)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.txt_file = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.txt_file.setFont(font)
        self.txt_file.setMouseTracking(True)
        self.txt_file.setObjectName("txt_file")
        self.horizontalLayout.addWidget(self.txt_file)
        self.btn_load = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.btn_load.setFont(font)
        self.btn_load.setObjectName("btn_load")
        self.horizontalLayout.addWidget(self.btn_load)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(False)
        font.setWeight(50)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.txt_v1 = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.txt_v1.setFont(font)
        self.txt_v1.setObjectName("txt_v1")
        self.horizontalLayout_2.addWidget(self.txt_v1)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(False)
        font.setWeight(50)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_2.addWidget(self.label_3)
        self.txt_v2 = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.txt_v2.setFont(font)
        self.txt_v2.setObjectName("txt_v2")
        self.horizontalLayout_2.addWidget(self.txt_v2)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(False)
        font.setWeight(50)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_3.addWidget(self.label_4)
        self.txt_deg1 = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.txt_deg1.setFont(font)
        self.txt_deg1.setObjectName("txt_deg1")
        self.horizontalLayout_3.addWidget(self.txt_deg1)
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(False)
        font.setWeight(50)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_3.addWidget(self.label_5)
        self.txt_deg2 = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.txt_deg2.setFont(font)
        self.txt_deg2.setObjectName("txt_deg2")
        self.horizontalLayout_3.addWidget(self.txt_deg2)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(False)
        font.setWeight(50)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_4.addWidget(self.label_6)
        self.txt_num = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.txt_num.setFont(font)
        self.txt_num.setObjectName("txt_num")
        self.horizontalLayout_4.addWidget(self.txt_num)
        self.btn_cal = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(19)
        self.btn_cal.setFont(font)
        self.btn_cal.setObjectName("btn_cal")
        self.horizontalLayout_4.addWidget(self.btn_cal)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.gridLayout.addLayout(self.verticalLayout, 0, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        # Bind ui function
        self.btn_load.clicked.connect(self.loadFile)
        self.btn_cal.clicked.connect(self.calculate)
        # self.txt_default_file = LineEdit()
        # self.txt_default_file.doubleClicked.connect(self.unlockTitle)

    def change_default(self):
        global initPara
        initPara = initParameter()

        self.excelPath = QFileDialog.getOpenFileName(filter="Excel File (*.xlsx *.xls)")[0]
        self.txt_new_default_file.setText((Path(self.excelPath).name))
        paraDf = pd.read_excel(self.excelPath, header=None)
        row_list = paraDf[0].tolist()
        sub = ['k', 'q', 'sn', 'nth', 'Nc', 'T']
        for i in range(len(sub)):
            for s in row_list:
                if sub[i] in s:
                    if (i == 0) : initPara.k = paraDf[1][row_list.index(s)]
                    elif (i == 1) : initPara.q = paraDf[1][row_list.index(s)]
                    elif (i == 2): initPara.sn = paraDf[1][row_list.index(s)]
                    elif (i == 3): initPara.nth = paraDf[1][row_list.index(s)]
                    elif (i == 4): initPara.Nc = paraDf[1][row_list.index(s)]
                    elif (i == 5): initPara.T = paraDf[1][row_list.index(s)]
                    else : break
                    break

        print(str(initPara.k) + ", " + str(initPara.q) + ", " + str(initPara.sn) + ", " + str(initPara.nth) + ", "+ str(initPara.Nc) + ", "+ str(initPara.T))


    def getPeakXY(self, df, deg1, deg2, nums, W):
        x = np.array(df["W"])
        y = np.array(df["Gp_W"])
        x_New = np.linspace(min(x), max(x), nums)  # Enhance accuracy
        V_XY_sheet1 = []
        V_XY_sheet2 = []
        V_XY_sheet3 = []

        for k in range(deg1, deg2, 1):
            p = np.polyfit(x, y, k)
            y_New = np.polyval(p, x_New)

            V_XY_sheet1.append(x_New[signal.argrelextrema(y_New, np.greater)[0]][0])  # append V_X-Axis
            V_XY_sheet1.append(y_New[signal.argrelextrema(y_New, np.greater)][0])  # append V_Y-Axis

            V_XY_sheet2.append(1 / x_New[signal.argrelextrema(y_New, np.greater)[0]][0])  # append V_X-Axis
            V_XY_sheet2.append(y_New[signal.argrelextrema(y_New, np.greater)][0] / ((1.6 * 1e-19) / 2))  # append V_Y-Axis

            V_XY_sheet3.append((initPara.k * float(self.txt_deault_T.text()) / initPara.q) * np.log(initPara.sn * initPara.nth * initPara.Nc * (1 / x_New[signal.argrelextrema(y_New, np.greater)[0]][0])))  # append V_X-Axis
            V_XY_sheet3.append(y_New[signal.argrelextrema(y_New, np.greater)][0] / ((1.6 * 1e-19) / 2))  # append V_Y-Axis

        results_df_sheet1.insert(len(results_df_sheet1.columns), str(W), V_XY_sheet1,
                                 True)  # insert by index (diff from task2)
        results_df_sheet2.insert(len(results_df_sheet2.columns), str(W), V_XY_sheet2,
                                 True)  # insert by index (diff from task2)
        results_df_sheet3.insert(len(results_df_sheet3.columns), str(W), V_XY_sheet3,
                                 True)  # insert by index (diff from task2)

    def loadFile(self):
        self.excelPath = QFileDialog.getOpenFileName(filter="Excel File (*.xlsx *.xls)")[0]
        self.txt_file.setText(self.excelPath)

        # Definite Value
        global init
        init = initVal()
        # Get dataframe dataset
        init.df1 = pd.read_excel(self.excelPath, index_col=0)
        init._W = init.df1.index.tolist()  # (6283.185307179586 ~　31415926.535897933)
        init._V = list(map(float, init.df1.columns.tolist()))  # (-10.0 ~ 10.0 )
        init._Area = 1 * 1e-3

        # Set default volt
        self.txt_v1.setText(str(init._V[0]))  # default val (V = -10.0)
        self.txt_v2.setText(str(init._V[-1]))  # default val (V = 10.0)
        # Set default ploy-degree (4-9)
        self.txt_deg1.setText(str(4))
        self.txt_deg2.setText(str(9))
        # Set default samples num (10000)
        self.txt_num.setText(str(10000))

    def calculate(self):
        if (int(self.txt_deg1.text()) < 4):  # limit degree > 3
            QMessageBox.warning(None, 'warning', 'degree < 4')  # limit deg1 < deg2
        elif (int(self.txt_deg1.text()) > int(self.txt_deg2.text())):
            QMessageBox.warning(None, 'warning', 'deg1 > deg2 ')
        else:
            # 0. Get customized value (Volt, degree, samples)
            # Duplicates element & Get minimum (Volt)
            v_start = indices(init._V, float(self.txt_v1.text()))
            v_end = indices(init._V, float(self.txt_v2.text()))
            if ((float(self.txt_v1.text()) < 0) and (float(self.txt_v2.text()) > 0)):  # -5 ~ 5
                min_V1 = min(v_start)
                min_V2 = min(v_end)
            else:  # equal or v1 > v2 ..
                min_V1 = min(v_start)
                min_V2 = max(v_end)
            # update _V
            init._V = np.array(init._V)[min_V1: (min_V2 + 1)].tolist()

            # Get Degree
            deg1 = int(self.txt_deg1.text())
            deg2 = int(self.txt_deg2.text()) + 1
            # Get Samples
            sampleNums = int(self.txt_num.text())

            # 2. cal "Gp_W" & "Peak coordinate
            global results_df_sheet1, results_df_sheet2, results_df_sheet3
            results_df_sheet1 = pd.DataFrame()
            results_df_sheet2 = pd.DataFrame()
            results_df_sheet3 = pd.DataFrame()

            title = []
            for k in range(deg1, deg2, 1):
                title.append(str(k) + "X")
                title.append(str(k) + "Y")
            results_df_sheet1["V"] = title
            results_df_sheet2["V"] = title
            results_df_sheet3["V"] = title

            Gp_w_arr = init.df1.to_numpy()
            for i in range(len(init._V)):
                Gp_W = []
                _each_W = []
                for j in range(len(init._W)):
                    Gp_W.append(float(Gp_w_arr[j, i]))
                    _each_W.append(float(init._W[j]))
                    Gp_W_Results = {
                        "W": _each_W,
                        "Gp_W": Gp_W,
                    }

                # Cal Peak XY-Axis on diff _V & ploy-deg
                self.getPeakXY(pd.DataFrame(Gp_W_Results), deg1, deg2, sampleNums, init._V[i])

            # 3. output results
            try:
                writer = pd.ExcelWriter('poly' + str(deg1) + '-' + str(deg2 - 1) + '_Peak_coordinate.xlsx',
                                        engine='xlsxwriter')
                # results_df.T.to_csv('poly' + str(deg1) + '-' + str(deg2 - 1) + '_Peak_coordinate.csv', encoding='utf-8', header=None)
                results_df_sheet1.T.to_excel(writer, sheet_name='Sheet1', header=None)
                results_df_sheet2.T.to_excel(writer, sheet_name='Sheet2', header=None)
                results_df_sheet3.T.to_excel(writer, sheet_name='Sheet3', header=None)
                writer.save()
                QMessageBox.information(None, 'Message', 'Finished')
            except:
                print("Output error")
                pass

            # # 4. Exit()
            # sys.exit(app.exec_())

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "EC - App2"))
        self.label_7.setText(_translate("MainWindow", "Default Parameter :"))
        self.txt_new_default_file.setText(_translate("MainWindow", "default_value.xlsx"))
        self.label_8.setText(_translate("MainWindow", "T :"))
        self.txt_deault_T.setText(_translate("MainWindow", "303"))
        self.label.setText(_translate("MainWindow", "Excel File :"))
        self.btn_load.setText(_translate("MainWindow", "Load"))
        self.label_2.setText(_translate("MainWindow", "Volt :"))
        self.label_3.setText(_translate("MainWindow", " ~ "))
        self.label_4.setText(_translate("MainWindow", "Degree of Polynomial :"))
        self.label_5.setText(_translate("MainWindow", " ~ "))
        self.label_6.setText(_translate("MainWindow", "Numbers of Samples : "))
        self.btn_cal.setText(_translate("MainWindow", "Calculate"))

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
