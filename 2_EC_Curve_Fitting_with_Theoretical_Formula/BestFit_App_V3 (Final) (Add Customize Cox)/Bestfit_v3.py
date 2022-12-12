#
# Created by NUZEROVI on 2022/12/10.
#

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import pyqtSignal, QEvent
from PyQt5.QtWidgets import QLineEdit, QCheckBox

import pandas as pd
import numpy as np
import math
from pathlib import Path

class initVal:
    def __init__(self):
        self._Freq = []
        self._W = []
        self._V = []
        self._Area = 1 * 1e-3
        self.CoxDefault = True

class pdGroups:
    def __init__(self):
        self.G1 = pd.DataFrame()
        self.G2 = pd.DataFrame()
        self.G3 = pd.DataFrame()

class initParameter: # (For Φ)
    def __init__(self):
        self.k = 1.38 * 1e-23
        self.q = 1.60 * 1e-19
        self.sn = 1.00 * 1e-14  # sigma (σ)
        self.nth = 1.90 * 1e+7
        self.Nc = 1.9 * 1e+19
        self.T = 303

global init
global dfGroups
global initPara     # for initParameter
initPara = initParameter()

def indices(lst, item):
    return [i for i, x in enumerate(lst) if x == item]

# ################################# Goodness of Fit (R^2)######################################
# reference "Goodness of Fit" example : https://blog.csdn.net/qq_43403025/article/details/108285275
def __sst(y_no_fitting):
    y_mean = sum(y_no_fitting) / len(y_no_fitting)
    s_list =[(y - y_mean)**2 for y in y_no_fitting]
    sst = sum(s_list)
    return sst

def __ssr(y_fitting, y_no_fitting):
    y_mean = sum(y_no_fitting) / len(y_no_fitting)
    s_list =[(y - y_mean)**2 for y in y_fitting]
    ssr = sum(s_list)
    return ssr

def __sse(y_fitting, y_no_fitting):
    s_list = [(y_fitting[i] - y_no_fitting[i])**2 for i in range(len(y_fitting))]
    sse = sum(s_list)
    return sse

def goodness_of_fit(y_fitting, y_no_fitting):
    SSR = __ssr(y_fitting, y_no_fitting)
    SST = __sst(y_no_fitting)
    SSE = __sse(y_fitting, y_no_fitting)
    ### Bind range from 0-1 (cause SSE > SST ...)
    r2 = 1 - (SSE / SST)
    if(r2 < 0 or r2 > 1) :
        r2 = SSR / SST
        if(r2 < 0 or r2 > 1) :
            r2 = -1
    return r2
# ########################################################################################### #

class Change_file(QLineEdit):
    doubleClicked = pyqtSignal()

    def event(self, event):
        if event.type() == QEvent.Type.MouseButtonDblClick:
            self.doubleClicked.emit()
        return super().event(event)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(687, 259)
        font = QtGui.QFont()
        font.setFamily("Dubai")
        font.setPointSize(18)
        MainWindow.setFont(font)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setContentsMargins(-1, -1, -1, 9)
        self.gridLayout.setSpacing(6)
        self.gridLayout.setObjectName("gridLayout")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setSpacing(3)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_7.setSpacing(3)
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.groupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox_2.sizePolicy().hasHeightForWidth())
        self.groupBox_2.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.groupBox_2.setFont(font)
        self.groupBox_2.setObjectName("groupBox_2")
        self.gridLayout_4 = QtWidgets.QGridLayout(self.groupBox_2)
        self.gridLayout_4.setContentsMargins(-1, 0, -1, 3)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setSpacing(3)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.groupBox_2)
        font = QtGui.QFont()
        font.setFamily("Dubai")
        font.setPointSize(16)
        font.setBold(False)
        font.setWeight(50)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.txt_file = QtWidgets.QLineEdit(self.groupBox_2)
        font = QtGui.QFont()
        font.setFamily("Dubai Light")
        font.setPointSize(16)
        self.txt_file.setFont(font)
        self.txt_file.setObjectName("txt_file")
        self.horizontalLayout.addWidget(self.txt_file)
        self.btn_load = QtWidgets.QPushButton(self.groupBox_2)
        font = QtGui.QFont()
        font.setFamily("Dubai Medium")
        font.setPointSize(18)
        self.btn_load.setFont(font)
        self.btn_load.setObjectName("btn_load")
        self.horizontalLayout.addWidget(self.btn_load)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.groupBox_2)
        font = QtGui.QFont()
        font.setFamily("Dubai")
        font.setPointSize(16)
        font.setBold(False)
        font.setWeight(50)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.txt_v1 = QtWidgets.QLineEdit(self.groupBox_2)
        font = QtGui.QFont()
        font.setFamily("Dubai Light")
        font.setPointSize(16)
        self.txt_v1.setFont(font)
        self.txt_v1.setObjectName("txt_v1")
        self.horizontalLayout_2.addWidget(self.txt_v1)
        self.label_3 = QtWidgets.QLabel(self.groupBox_2)
        font = QtGui.QFont()
        font.setFamily("Dubai")
        font.setPointSize(16)
        font.setBold(False)
        font.setWeight(50)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_2.addWidget(self.label_3)
        self.txt_v2 = QtWidgets.QLineEdit(self.groupBox_2)
        font = QtGui.QFont()
        font.setFamily("Dubai Light")
        font.setPointSize(16)
        self.txt_v2.setFont(font)
        self.txt_v2.setObjectName("txt_v2")
        self.horizontalLayout_2.addWidget(self.txt_v2)
        self.checkBox = QtWidgets.QCheckBox(self.groupBox_2)
        self.checkBox.setSizeIncrement(QtCore.QSize(0, 0))
        self.checkBox.setBaseSize(QtCore.QSize(0, 0))
        font = QtGui.QFont()
        font.setFamily("Dubai")
        font.setPointSize(16)
        self.checkBox.setFont(font)
        self.checkBox.setIconSize(QtCore.QSize(16, 16))
        self.checkBox.setCheckable(True)
        self.checkBox.setChecked(True)
        self.checkBox.setTristate(False)
        self.checkBox.setObjectName("checkBox")
        self.horizontalLayout_2.addWidget(self.checkBox)

        self.checkBox.clicked.connect(self.onCheckBoxClick)

        self.lineEdit = QtWidgets.QLineEdit(self.groupBox_2)
        self.lineEdit.setEnabled(False)
        font = QtGui.QFont()
        font.setFamily("Dubai Light")
        font.setPointSize(13)
        font.setItalic(False)
        self.lineEdit.setFont(font)
        self.lineEdit.setInputMask("")
        self.lineEdit.setText("")
        self.lineEdit.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout_2.addWidget(self.lineEdit)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.gridLayout_4.addLayout(self.verticalLayout, 0, 0, 1, 1)
        self.horizontalLayout_3.addWidget(self.groupBox_2)
        self.horizontalLayout_7.addLayout(self.horizontalLayout_3)
        self.verticalLayout_2.addLayout(self.horizontalLayout_7)
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setSpacing(3)
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.horizontalGroupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
        self.horizontalGroupBox_2.setEnabled(True)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.horizontalGroupBox_2.setFont(font)
        self.horizontalGroupBox_2.setObjectName("horizontalGroupBox_2")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.horizontalGroupBox_2)
        self.gridLayout_2.setContentsMargins(-1, 0, -1, 3)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setSpacing(3)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setContentsMargins(0, -1, -1, -1)
        self.horizontalLayout_4.setSpacing(6)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_4 = QtWidgets.QLabel(self.horizontalGroupBox_2)
        self.label_4.setEnabled(True)
        font = QtGui.QFont()
        font.setFamily("Dubai")
        font.setPointSize(16)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_4.addWidget(self.label_4)
        self.txt_default_file = QtWidgets.QLineEdit(self.horizontalGroupBox_2)
        self.txt_default_file.setEnabled(False)
        font = QtGui.QFont()
        font.setFamily("Dubai Light")
        font.setPointSize(16)
        self.txt_default_file.setFont(font)
        self.txt_default_file.setObjectName("txt_default_file")
        self.horizontalLayout_4.addWidget(self.txt_default_file)
        self.txt_default_file.setFrame(False)

        # Bind doubleClicked on QlineEdit
        self.txt_new_default_file = Change_file(self.txt_default_file)
        self.txt_new_default_file.doubleClicked.connect(self.change_default)

        self.label_5 = QtWidgets.QLabel(self.horizontalGroupBox_2)
        self.label_5.setEnabled(True)
        font = QtGui.QFont()
        font.setFamily("Dubai")
        font.setPointSize(16)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_4.addWidget(self.label_5)
        self.txt_deault_T = QtWidgets.QLineEdit(self.horizontalGroupBox_2)
        self.txt_deault_T.setEnabled(False)
        font = QtGui.QFont()
        font.setFamily("Dubai Light")
        font.setPointSize(16)
        self.txt_deault_T.setFont(font)
        self.txt_deault_T.setObjectName("txt_deault_T")
        self.horizontalLayout_4.addWidget(self.txt_deault_T)
        self.verticalLayout_3.addLayout(self.horizontalLayout_4)
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.label_6 = QtWidgets.QLabel(self.horizontalGroupBox_2)
        self.label_6.setEnabled(True)
        font = QtGui.QFont()
        font.setFamily("Dubai")
        font.setPointSize(16)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_5.addWidget(self.label_6)
        self.txt_excel = QtWidgets.QLineEdit(self.horizontalGroupBox_2)
        self.txt_excel.setEnabled(False)
        font = QtGui.QFont()
        font.setFamily("Dubai Light")
        font.setPointSize(16)
        self.txt_excel.setFont(font)
        self.txt_excel.setText("")
        self.txt_excel.setObjectName("txt_excel")
        self.horizontalLayout_5.addWidget(self.txt_excel)
        self.verticalLayout_3.addLayout(self.horizontalLayout_5)
        self.gridLayout_2.addLayout(self.verticalLayout_3, 0, 0, 1, 2)
        self.horizontalLayout_6.addWidget(self.horizontalGroupBox_2)
        self.horizontalGroupBox_3 = QtWidgets.QGroupBox(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.horizontalGroupBox_3.setFont(font)
        self.horizontalGroupBox_3.setObjectName("horizontalGroupBox_3")
        self.gridLayout_5 = QtWidgets.QGridLayout(self.horizontalGroupBox_3)
        self.gridLayout_5.setContentsMargins(-1, 0, -1, 3)
        self.gridLayout_5.setObjectName("gridLayout_5")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout()
        self.verticalLayout_5.setSpacing(3)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.btn_fit = QtWidgets.QPushButton(self.horizontalGroupBox_3)
        self.btn_fit.setEnabled(False)
        font = QtGui.QFont()
        font.setFamily("Dubai Medium")
        font.setPointSize(18)
        self.btn_fit.setFont(font)
        self.btn_fit.setObjectName("btn_fit")
        self.verticalLayout_5.addWidget(self.btn_fit)
        self.gridLayout_5.addLayout(self.verticalLayout_5, 0, 0, 1, 1)
        self.horizontalLayout_6.addWidget(self.horizontalGroupBox_3)
        self.verticalLayout_2.addLayout(self.horizontalLayout_6)
        self.gridLayout.addLayout(self.verticalLayout_2, 1, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        self.txt_default_file.textChanged['QString'].connect(self.txt_default_file.update)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        # Bind ui function
        self.btn_load.clicked.connect(self.loadFile)
        self.btn_fit.clicked.connect(self.fit)

    def onCheckBoxClick(self):
        if self.checkBox.isChecked(): # Default
            self.lineEdit.setEnabled(False)
            self.lineEdit.setText("")
            self.CoxDefault = True
        else: #User-Define
            self.lineEdit.setEnabled(True)
            self.CoxDefault = False

    def change_default(self):
        global initPara
        initPara = initParameter()

        self.excelPath = QFileDialog.getOpenFileName(filter="Excel File (*.xlsx *.xls)")[0]
        if((Path(self.excelPath).name) != ''):
            self.txt_new_default_file.setText((Path(self.excelPath).name))
            paraDf = pd.read_excel(self.excelPath, header=None)
            row_list = paraDf[0].tolist()
            sub = ['k', 'q', 'sn', 'nth', 'Nc', 'T']
            for i in range(len(sub)):
                for s in row_list:
                    if sub[i] in s:
                        if (i == 0):
                            initPara.k = paraDf[1][row_list.index(s)]
                        elif (i == 1):
                            initPara.q = paraDf[1][row_list.index(s)]
                        elif (i == 2):
                            initPara.sn = paraDf[1][row_list.index(s)]
                        elif (i == 3):
                            initPara.nth = paraDf[1][row_list.index(s)]
                        elif (i == 4):
                            initPara.Nc = paraDf[1][row_list.index(s)]
                        elif (i == 5):
                            initPara.T = paraDf[1][row_list.index(s)]
                            self.txt_deault_T.setText(str(paraDf[1][row_list.index(s)]))
                        else:
                            break
                        break

            print("k = " +
                str(initPara.k) + ", q = " + str(initPara.q) + ", sn = " + str(initPara.sn) + ", nth = " + str(initPara.nth) + ", Nc = " + str(
                    initPara.Nc) + ", T = " + str(initPara.T))

    def SetUI(self):
        self.txt_excel.setText("Gp_W_Results.xlsx")
        self.txt_excel.setStyleSheet('color: green;')

    def Get_τ_and_D(self, i):
        _each_W = []
        for j in range(len(init._Freq)):
            _each_W.append(float(init._W[j]))

        results = {
            "W": sorted(_each_W),
        }
        df = pd.DataFrame(results)

        init._W = sorted(init._W)
        Gp_W = []

        for j in range(len(init._Freq)):
            Cm = dfGroups.G1.iat[i, j] / init._Area
            Gm = dfGroups.G2.iat[i, j] / init._Area
            if(init.CoxDefault): # default
                Cox = (dfGroups.G1.iloc[:, j].max()) / init._Area
            else:
                Cox = float(self.lineEdit.text()) / init._Area

            Gp_W.append((init._W[j] * math.pow(Cox, 2) * Gm) / (
                    math.pow(Gm, 2) + (math.pow(init._W[j], 2) * math.pow(Cm - Cox, 2))))
        df.insert(len(df.columns), str(init._V[i]), Gp_W, True)
        xx = np.array(df["W"])
        yy = np.array(df.iloc[:, 1])

        # PeakX = (1 / xx[np.argmax(yy)])
        PeakY_D = (yy[np.argmax(yy)] * 2 / initPara.q)
        return df, PeakY_D

    def cal_TH_GP_W(self, df, i, D_val):
        _each_W = []
        for j in range(len(init._Freq)):
            _each_W.append(float(init._W[j]))

        results = {
            "W": sorted(_each_W),
        }

        dfTH_tmp = pd.DataFrame(results)
        init._W = sorted(init._W)

        xx = np.array(df["W"])
        yy = np.array(df.iloc[:, 1])

        τit = (1 / xx[np.argmax(yy)])
        Dit = (yy[np.argmax(yy)] * 2 / initPara.q) - D_val

        Theoretical_Gp_W = []
        for j in range(len(init._Freq)):
            w = init._W[j]
            Theoretical_Gp_W.append((initPara.q * w * τit * Dit) / (1 + (math.pow(w * τit, 2))))

        dfTH_tmp.insert(len(dfTH_tmp.columns), str(init._V[i]), Theoretical_Gp_W, True)

        return dfTH_tmp, τit, Dit

    def cal_BestFit_GP_W(self, maxτ, maxD):

        τit = maxτ
        Dit = maxD

        Theoretical_Gp_W_Best = []
        for j in range(len(init._Freq)):
            w = init._W[j]
            Theoretical_Gp_W_Best.append((initPara.q * w * τit * Dit) / (1 + (math.pow(w * τit, 2))))

        return Theoretical_Gp_W_Best

    def handleStatusMessage(self, txt):
        self.statusbar.showMessage(txt)

    def fit(self):
        GP_W_df, dfTH_Best = self.output_GP_W()
        QtWidgets.QApplication.processEvents()
        self.SetUI()
        global bestFit_df
        bestFit_df = pd.DataFrame(columns=['Volt', 'τit', 'Dit', 'RR', 'Φ'])
        for i in range(len(init._V)):  # len(init._V)
            QtWidgets.QApplication.processEvents()
            self.statusbar.showMessage("Please wait for the best fit ... (" + str(i) + "/" + str(len(init._V)) + ")")
            df, D = self.Get_τ_and_D(i)
            b = np.array(df.iloc[:, 1])  # Gp_W
            bmin = D / 10
            bmax = D * 10

            # initial Value
            tmp = 0
            rr = 0
            maxτ = 0
            maxD = 0
            sheet1 = []
            for j in np.arange(0, bmin, ((bmax - bmin) / 1e+5)):
                dfTH, τit, Dit = self.cal_TH_GP_W(df, i, j)

                a = np.array(dfTH.iloc[:, 1])  # Theoretical_Gp_W
                tmp = goodness_of_fit(a, b)
                if (tmp > rr):
                    maxτ = τit
                    maxD = Dit
                    rr = max(tmp, rr)  # Get Best Fit

            Φ = (initPara.k * float(self.txt_deault_T.text()) / initPara.q) * np.log(
                initPara.sn * initPara.nth * initPara.Nc * (1 / τit))
            sheet1.append(str(init._V[i]))
            sheet1.append(float(maxτ))
            sheet1.append(float(maxD))
            sheet1.append(float(rr))
            sheet1.append(float(Φ))

            bestFit_df.loc[len(bestFit_df)] = sheet1
            dfTH_Best.insert(len(dfTH_Best.columns), str(init._V[i]), self.cal_BestFit_GP_W(maxτ, maxD), True)

        QtWidgets.QApplication.processEvents()
        self.statusbar.showMessage("Finished fitting ...")
        try:
            QtWidgets.QApplication.processEvents()
            self.statusbar.showMessage("Create output file ...")
            with pd.ExcelWriter("BestFit.xlsx") as writer:
                bestFit_df.to_excel(writer, sheet_name="Sheet1", index=False)
                GP_W_df.to_excel(writer, sheet_name="Sheet2", index=False)
                dfTH_Best.to_excel(writer, sheet_name="Sheet3", index=False)
            QMessageBox.information(None, 'Message', 'Finished')
            QtWidgets.QApplication.processEvents()
            self.statusbar.showMessage("Finished.")
        except:
            print("BestFit.xlsx Output error")
            pass

    def loadFile(self):
        global init
        self.excelPaths = QFileDialog.getOpenFileNames(filter="CSV Files (*.csv)")[0]
        if len(self.excelPaths) != 0:
            try:
                self.filenames = ""
                if (self.excelPaths):
                    for self.filename in self.excelPaths:  # 4
                        self.filenames += (Path(self.filename).name + ",")

                self.txt_file.setText(self.filenames)

                # Multiple dataframes
                dfs = []
                init = initVal()
                for i in range(len(self.excelPaths)):
                    dfs.append("df" + str(i))
                    dfs[i] = pd.DataFrame()
                    dfs[i] = pd.read_csv(self.excelPaths[i], encoding='utf-8', sep='\t')
                    dfs[i] = dfs[i].iloc[:, 0].str.split(',', expand=True)

                    # auto retrieve VBias (if dataset change)
                    a = dfs[i].iloc[:, 0].str.strip().tolist()
                    indexes = [index for index in range(len(a)) if a[index] == 'DataName']
                    if len(indexes) == 0:
                        indexes = [index for index in range(len(a)) if a[index] == 'VBias']
                        if len(indexes) == 0:
                            a = dfs[i].iloc[:, 1].str.strip().tolist()
                            indexes = [index for index in range(len(a)) if a[index] == 'VBias']
                            dfs[i] = dfs[i].iloc[:, 1:][indexes[0]:].dropna(axis=1)
                        else:
                            dfs[i] = dfs[i].iloc[:, 0:][indexes[0]:].dropna(axis=1)
                    else:
                        dfs[i] = dfs[i].iloc[:, 1:][indexes[0]:].dropna(axis=1)
                    dfs[i] = dfs[i].rename(columns=dfs[i].iloc[0]).drop(dfs[i].index[0]).reset_index(drop=True)
                    dfs[i].columns = dfs[i].columns.str.strip()  # To strip whitespaces
                    dfs[i] = dfs[i].loc[:, ["VBias", "Freq", "C", "G"]].apply(pd.to_numeric, 1)

                    # Definite Value (_Freq, _V, _Area)
                    init._Freq += sorted(list(set(dfs[i]["Freq"].tolist())))
                    init._V = dfs[0].query("Freq == @init._Freq[0]")[
                        "VBias"].tolist()  # all V
                    init._Area = 1 * 1e-3

                # # 152 rows × 63 columns
                global dfGroups
                dfGroups = pdGroups()
                dfGroups.G1 = pd.DataFrame(columns=init._Freq)  # C
                dfGroups.G2 = pd.DataFrame(columns=init._Freq)  # G
                for i in range(len(init._Freq)):
                    init._W.append(2 * (math.pi) * init._Freq[i])  # Definite Value (_W = 2*pi*_Freq)
                    listC = []
                    listG = []
                    for j in range(len(self.excelPaths)):
                        listC += dfs[j].query("Freq == @init._Freq[@i]")["C"].tolist()
                        listG += dfs[j].query("Freq == @init._Freq[@i]")["G"].tolist()

                    dfGroups.G1[init._Freq[i]] = listC
                    dfGroups.G2[init._Freq[i]] = listG

                # Set default volt
                self.txt_v1.setText(str(init._V[0]))  # default val (V = -15)
                self.txt_v2.setText(str(init._V[-1]))  # default val (V = 15)

                #
                QtWidgets.QApplication.processEvents()
                self.btn_fit.setEnabled(True)

            except:
                print("An error occurred while loading the file")
                pass

    def output_GP_W(self):
        # 0. Update "_V" by windows Form (customized Value)
        # Duplicates element & Get minimum
        v_start = indices(init._V, float(self.txt_v1.text()))
        v_end = indices(init._V, float(self.txt_v2.text()))
        if ((float(self.txt_v1.text()) < 0) and (float(self.txt_v2.text()) > 0)):  # -5 ~ 5
            min_V1 = min(v_start)
            min_V2 = min(v_end)
        else:  # equal or v1 > v2 ..
            min_V1 = min(v_start)
            min_V2 = max(v_end)

        # update _V & dfGroups
        init._V = np.array(init._V)[min_V1: (min_V2 + 1)].tolist()
        dfGroups.G1 = pd.DataFrame(dfGroups.G1.to_numpy()[min_V1: (min_V2 + 1)])
        dfGroups.G2 = pd.DataFrame(dfGroups.G2.to_numpy()[min_V1: (min_V2 + 1)])

        # 1. "W" in colume[0]
        _each_W = []
        for j in range(len(init._Freq)):
            _each_W.append(float(init._W[j]))

        results = {
            "W": sorted(_each_W),
        }
        df = pd.DataFrame(results)
        df_Best = pd.DataFrame(results)

        # 2. cal "Gp_W" (append another column) (behind colume[0])
        init._W = sorted(init._W)

        for i in range(len(init._V)):  # len(init._V)
            Gp_W = []
            for j in range(len(init._Freq)):
                Cm = dfGroups.G1.iat[i, j] / init._Area
                Gm = dfGroups.G2.iat[i, j] / init._Area
                if (init.CoxDefault):  # default
                    Cox = (dfGroups.G1.iloc[:, j].max()) / init._Area
                else:
                    Cox = float(self.lineEdit.text()) / init._Area

                Gp_W.append((init._W[j] * math.pow(Cox, 2) * Gm) / (
                        math.pow(Gm, 2) + (math.pow(init._W[j], 2) * math.pow(Cm - Cox, 2))))
            df.insert(len(df.columns), str(init._V[i]), Gp_W, True)

        return df, df_Best

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "BestFit_App_v3"))
        self.groupBox_2.setTitle(_translate("MainWindow", "Input Files"))
        self.label.setText(_translate("MainWindow", "CSV Files :"))
        self.btn_load.setText(_translate("MainWindow", "Load"))
        self.label_2.setText(_translate("MainWindow", "Volt :"))
        self.label_3.setText(_translate("MainWindow", " ~ "))
        self.checkBox.setText(_translate("MainWindow", "Cox"))
        self.lineEdit.setPlaceholderText(_translate("MainWindow", "default"))
        self.horizontalGroupBox_2.setTitle(_translate("MainWindow", "Experimental Results (With Dafault File)"))
        self.label_4.setText(_translate("MainWindow", "Parameter :"))
        self.txt_new_default_file.setText(_translate("MainWindow", "default_value.xlsx"))
        self.label_5.setText(_translate("MainWindow", "T :"))
        self.txt_deault_T.setText(_translate("MainWindow", "303.0"))
        self.label_6.setText(_translate("MainWindow", "Experimental Results :"))
        self.horizontalGroupBox_3.setTitle(_translate("MainWindow", "With Theoretical Formula"))
        self.btn_fit.setText(_translate("MainWindow", "Best Fit"))

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
