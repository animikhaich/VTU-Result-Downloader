from PyQt5 import QtCore, QtGui, QtWidgets
from main_cbcs import CBCS
from main_non_cbcs import NON_CBCS
from main_mtech_mca import PG
import re


class Ui_MainWindow(object):
    def __init__(self):
        super(Ui_MainWindow, self).__init__()
        self.parent_url_cbcs = None
        self.parent_url_non_cbcs = None
        self.start_usn_cbcs = None
        self.start_usn_non_cbcs = None
        self.end_usn_cbcs = None
        self.end_usn_non_cbcs = None
        self.file_path_cbcs = None
        self.file_path_non_cbcs = None
        self.cbcs = None
        self.non_cbcs = None
        self.regexp_ug = '^\d[A-Z]{2}\d{2}[A-Z]{2}\d{3}$'
        self.regexp_pg = '^\d[A-Z]{2}\d{2}[A-Z]{3}\d{2}$'

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(640, 480)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("Assets/icon_VTU.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setTextFormat(QtCore.Qt.RichText)
        self.label.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.label.setObjectName("label_2")
        self.label.setTextInteractionFlags(QtCore.Qt.LinksAccessibleByMouse)
        self.label.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
        self.label.setOpenExternalLinks(True)
        self.horizontalLayout.addWidget(self.label)

        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)

        self.quitButton = QtWidgets.QToolButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.quitButton.setFont(font)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("Assets/icon_close.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.quitButton.setIcon(icon1)
        self.quitButton.setIconSize(QtCore.QSize(20, 20))
        self.quitButton.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self.quitButton.setAutoRaise(True)
        self.quitButton.setObjectName("quitButton")
        self.horizontalLayout.addWidget(self.quitButton)

        self.gridLayout.addLayout(self.horizontalLayout, 1, 0, 1, 1)

        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName("tabWidget")

        self.non_cbcs_tab = QtWidgets.QWidget()
        self.non_cbcs_tab.setObjectName("non_cbcs_tab")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.non_cbcs_tab)
        self.verticalLayout.setObjectName("verticalLayout")
        self.formLayout = QtWidgets.QFormLayout()
        self.formLayout.setObjectName("formLayout")

        self.non_cbcs_url_label = QtWidgets.QLabel(self.non_cbcs_tab)
        self.non_cbcs_url_label.setObjectName("non_cbcs_url_label")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.non_cbcs_url_label)

        self.non_cbcs_url_lineEdit = QtWidgets.QLineEdit(self.non_cbcs_tab)
        self.non_cbcs_url_lineEdit.setObjectName("non_cbcs_url_lineEdit")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.non_cbcs_url_lineEdit)

        self.non_cbcs_usn_start_label = QtWidgets.QLabel(self.non_cbcs_tab)
        self.non_cbcs_usn_start_label.setObjectName("non_cbcs_usn_start_label")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.non_cbcs_usn_start_label)

        self.non_cbcs_usn_start_lineEdit = QtWidgets.QLineEdit(self.non_cbcs_tab)
        self.non_cbcs_usn_start_lineEdit.setObjectName("non_cbcs_usn_start_lineEdit")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.non_cbcs_usn_start_lineEdit)

        self.non_cbcs_usn_end_label = QtWidgets.QLabel(self.non_cbcs_tab)
        self.non_cbcs_usn_end_label.setObjectName("non_cbcs_usn_end_label")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.non_cbcs_usn_end_label)

        self.non_cbcs_usn_end_lineEdit = QtWidgets.QLineEdit(self.non_cbcs_tab)
        self.non_cbcs_usn_end_lineEdit.setObjectName("non_cbcs_usn_end_lineEdit")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.non_cbcs_usn_end_lineEdit)

        self.non_cbcs_file_path_label = QtWidgets.QLabel(self.non_cbcs_tab)
        self.non_cbcs_file_path_label.setObjectName("non_cbcs_file_path_label")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.LabelRole, self.non_cbcs_file_path_label)

        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")

        self.non_cbcs_file_path_lineEdit = QtWidgets.QLineEdit(self.non_cbcs_tab)
        self.non_cbcs_file_path_lineEdit.setObjectName("non_cbcs_file_path_lineEdit")
        self.horizontalLayout_2.addWidget(self.non_cbcs_file_path_lineEdit)

        self.non_cbcs_file_browse_toolButton = QtWidgets.QToolButton(self.non_cbcs_tab)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("Assets/icon_browse.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.non_cbcs_file_browse_toolButton.setIcon(icon2)
        self.non_cbcs_file_browse_toolButton.setIconSize(QtCore.QSize(50, 25))
        self.non_cbcs_file_browse_toolButton.setAutoRaise(True)
        self.non_cbcs_file_browse_toolButton.setObjectName("non_cbcs_file_browse_toolButton")
        self.horizontalLayout_2.addWidget(self.non_cbcs_file_browse_toolButton)

        self.formLayout.setLayout(3, QtWidgets.QFormLayout.FieldRole, self.horizontalLayout_2)
        self.verticalLayout.addLayout(self.formLayout)
        self.gridLayout_2 = QtWidgets.QGridLayout()
        self.gridLayout_2.setObjectName("gridLayout_2")
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_2.addItem(spacerItem1, 0, 1, 1, 1)
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_2.addItem(spacerItem2, 1, 0, 1, 1)

        self.non_cbcs_download_toolButton = QtWidgets.QToolButton(self.non_cbcs_tab)
        self.non_cbcs_download_toolButton.setEnabled(True)
        font = QtGui.QFont()
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.non_cbcs_download_toolButton.setFont(font)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("Assets/icon_download.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.non_cbcs_download_toolButton.setIcon(icon3)
        self.non_cbcs_download_toolButton.setIconSize(QtCore.QSize(50, 50))
        self.non_cbcs_download_toolButton.setToolButtonStyle(QtCore.Qt.ToolButtonTextUnderIcon)
        self.non_cbcs_download_toolButton.setAutoRaise(True)
        self.non_cbcs_download_toolButton.setObjectName("non_cbcs_download_toolButton")
        self.gridLayout_2.addWidget(self.non_cbcs_download_toolButton, 1, 1, 1, 1)

        spacerItem3 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_2.addItem(spacerItem3, 1, 2, 1, 1)
        spacerItem4 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_2.addItem(spacerItem4, 2, 1, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout_2)

        self.non_cbcs_progressBar = QtWidgets.QProgressBar(self.non_cbcs_tab)
        self.non_cbcs_progressBar.setProperty("value", 0)
        self.non_cbcs_progressBar.setAlignment(QtCore.Qt.AlignCenter)
        self.non_cbcs_progressBar.setTextVisible(True)
        self.non_cbcs_progressBar.setInvertedAppearance(False)
        self.non_cbcs_progressBar.setTextDirection(QtWidgets.QProgressBar.TopToBottom)
        self.non_cbcs_progressBar.setObjectName("non_cbcs_progressBar")
        self.verticalLayout.addWidget(self.non_cbcs_progressBar)

        self.tabWidget.addTab(self.non_cbcs_tab, "")
        self.cbcs_tab = QtWidgets.QWidget()
        self.cbcs_tab.setObjectName("cbcs_tab")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.cbcs_tab)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.formLayout_2 = QtWidgets.QFormLayout()
        self.formLayout_2.setObjectName("formLayout_2")

        self.cbcs_url_label = QtWidgets.QLabel(self.cbcs_tab)
        self.cbcs_url_label.setObjectName("cbcs_url_label")
        self.formLayout_2.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.cbcs_url_label)

        self.cbcs_url_lineEdit = QtWidgets.QLineEdit(self.cbcs_tab)
        self.cbcs_url_lineEdit.setObjectName("cbcs_url_lineEdit")
        self.formLayout_2.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.cbcs_url_lineEdit)

        self.cbcs_usn_start_label = QtWidgets.QLabel(self.cbcs_tab)
        self.cbcs_usn_start_label.setObjectName("cbcs_usn_start_label")
        self.formLayout_2.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.cbcs_usn_start_label)

        self.cbcs_usn_start_lineEdit = QtWidgets.QLineEdit(self.cbcs_tab)
        self.cbcs_usn_start_lineEdit.setObjectName("cbcs_usn_start_lineEdit")
        self.formLayout_2.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.cbcs_usn_start_lineEdit)

        self.cbcs_usn_end_label = QtWidgets.QLabel(self.cbcs_tab)
        self.cbcs_usn_end_label.setObjectName("cbcs_usn_end_label")
        self.formLayout_2.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.cbcs_usn_end_label)

        self.cbcs_usn_end_lineEdit = QtWidgets.QLineEdit(self.cbcs_tab)
        self.cbcs_usn_end_lineEdit.setObjectName("cbcs_usn_end_lineEdit")
        self.formLayout_2.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.cbcs_usn_end_lineEdit)

        self.cbcs_file_path_label = QtWidgets.QLabel(self.cbcs_tab)
        self.cbcs_file_path_label.setObjectName("cbcs_file_path_label")
        self.formLayout_2.setWidget(3, QtWidgets.QFormLayout.LabelRole, self.cbcs_file_path_label)

        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.cbcs_file_path_lineEdit = QtWidgets.QLineEdit(self.cbcs_tab)
        self.cbcs_file_path_lineEdit.setObjectName("cbcs_file_path_lineEdit")
        self.horizontalLayout_3.addWidget(self.cbcs_file_path_lineEdit)

        self.cbcs_file_browse_toolButton = QtWidgets.QToolButton(self.cbcs_tab)
        self.cbcs_file_browse_toolButton.setIcon(icon2)
        self.cbcs_file_browse_toolButton.setIconSize(QtCore.QSize(50, 25))
        self.cbcs_file_browse_toolButton.setAutoRaise(True)
        self.cbcs_file_browse_toolButton.setObjectName("cbcs_file_browse_toolButton")
        self.horizontalLayout_3.addWidget(self.cbcs_file_browse_toolButton)

        self.formLayout_2.setLayout(3, QtWidgets.QFormLayout.FieldRole, self.horizontalLayout_3)
        self.verticalLayout_2.addLayout(self.formLayout_2)
        self.gridLayout_3 = QtWidgets.QGridLayout()
        self.gridLayout_3.setObjectName("gridLayout_3")
        spacerItem5 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_3.addItem(spacerItem5, 0, 1, 1, 1)
        spacerItem6 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_3.addItem(spacerItem6, 1, 0, 1, 1)

        self.cbcs_download_toolButton = QtWidgets.QToolButton(self.cbcs_tab)
        font = QtGui.QFont()
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.cbcs_download_toolButton.setFont(font)
        self.cbcs_download_toolButton.setIcon(icon3)
        self.cbcs_download_toolButton.setIconSize(QtCore.QSize(50, 50))
        self.cbcs_download_toolButton.setToolButtonStyle(QtCore.Qt.ToolButtonTextUnderIcon)
        self.cbcs_download_toolButton.setAutoRaise(True)
        self.cbcs_download_toolButton.setObjectName("cbcs_download_toolButton")
        self.gridLayout_3.addWidget(self.cbcs_download_toolButton, 1, 1, 1, 1)

        spacerItem7 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_3.addItem(spacerItem7, 1, 2, 1, 1)
        spacerItem8 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_3.addItem(spacerItem8, 2, 1, 1, 1)
        self.verticalLayout_2.addLayout(self.gridLayout_3)

        self.cbcs_progressBar = QtWidgets.QProgressBar(self.cbcs_tab)
        self.cbcs_progressBar.setProperty("value", 0)
        self.cbcs_progressBar.setAlignment(QtCore.Qt.AlignCenter)
        self.cbcs_progressBar.setObjectName("cbcs_progressBar")
        self.verticalLayout_2.addWidget(self.cbcs_progressBar)

        self.tabWidget.addTab(self.cbcs_tab, "")
        self.gridLayout.addWidget(self.tabWidget, 0, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        self.quitButton.clicked.connect(MainWindow.close)

        self.cbcs_download_toolButton.clicked.connect(self.cbcs_download)
        self.non_cbcs_download_toolButton.clicked.connect(self.non_cbcs_download)

        self.cbcs_file_browse_toolButton.clicked.connect(self.get_file_path_cbcs)
        self.non_cbcs_file_browse_toolButton.clicked.connect(self.get_file_path_non_cbcs)

        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def get_file_path_cbcs(self):
        cbcs_path = QtWidgets.QFileDialog.getSaveFileName(filter='*xlsx')
        self.cbcs_file_path_lineEdit.setText(cbcs_path[0])

    def get_file_path_non_cbcs(self):
        non_cbcs_path = QtWidgets.QFileDialog.getSaveFileName(filter='*xlsx')
        self.non_cbcs_file_path_lineEdit.setText(non_cbcs_path[0])


    def update_cbcs_progress_bar(self, cbcs_val):
        self.cbcs_progressBar.setValue(cbcs_val)

    def update_non_cbcs_progress_bar(self, non_cbcs_val):
        self.non_cbcs_progressBar.setValue(non_cbcs_val)

    def enable_non_cbcs_download_button(self, non_cbcs_progress_val):
        if non_cbcs_progress_val > 99:
            self.non_cbcs_download_toolButton.setDisabled(False)

    def enable_cbcs_download_button(self, cbcs_progress_val):
        if cbcs_progress_val > 99:
            self.cbcs_download_toolButton.setDisabled(False)

    def cbcs_download(self):
        self.parent_url_cbcs = self.cbcs_url_lineEdit.text()
        self.start_usn_cbcs = self.cbcs_usn_start_lineEdit.text().upper()
        self.end_usn_cbcs = self.cbcs_usn_end_lineEdit.text().upper()
        self.file_path_cbcs = self.cbcs_file_path_lineEdit.text()

        if re.search(self.regexp_ug, self.start_usn_cbcs) and re.search(self.regexp_ug, self.end_usn_cbcs):
            self.cbcs_download_toolButton.setDisabled(True)
            self.cbcs = CBCS(self.parent_url_cbcs, self.start_usn_cbcs, self.end_usn_cbcs, self.file_path_cbcs)
            self.cbcs.start()
            self.cbcs.cbcs_update_progress.connect(self.update_cbcs_progress_bar)
            self.cbcs.cbcs_update_progress.connect(self.enable_cbcs_download_button)

        elif re.search(self.regexp_pg, self.start_usn_cbcs) and re.search(self.regexp_pg, self.end_usn_cbcs):
            self.cbcs_download_toolButton.setDisabled(True)
            print("1")
            self.cbcs = PG(self.parent_url_cbcs, self.start_usn_cbcs, self.end_usn_cbcs, self.file_path_cbcs)
            print("2")
            self.cbcs.start()
            print("3")
            self.cbcs.cbcs_update_progress.connect(self.update_cbcs_progress_bar)
            print("4")
            self.cbcs.cbcs_update_progress.connect(self.enable_cbcs_download_button)

    def non_cbcs_download(self):
        self.parent_url_non_cbcs = self.non_cbcs_url_lineEdit.text()
        self.start_usn_non_cbcs = self.non_cbcs_usn_start_lineEdit.text().upper()
        self.end_usn_non_cbcs = self.non_cbcs_usn_end_lineEdit.text().upper()
        self.file_path_non_cbcs = self.non_cbcs_file_path_lineEdit.text()

        if re.search(self.regexp_ug, self.start_usn_non_cbcs) and re.search(self.regexp_ug, self.end_usn_non_cbcs):
            self.non_cbcs_download_toolButton.setDisabled(True)
            self.non_cbcs = NON_CBCS(self.parent_url_non_cbcs, self.start_usn_non_cbcs, self.end_usn_non_cbcs, self.file_path_non_cbcs)
            self.non_cbcs.start()
            self.non_cbcs.non_cbcs_update_progress.connect(self.update_non_cbcs_progress_bar)
            self.non_cbcs.non_cbcs_update_progress.connect(self.enable_non_cbcs_download_button)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "VTU Result Downloader"))
        self.quitButton.setText(_translate("MainWindow", "Quit"))
        self.non_cbcs_url_label.setText(_translate("MainWindow", "Result URL:"))
        self.non_cbcs_url_lineEdit.setText(_translate("MainWindow", "http://results.vtu.ac.in/vitaviresultnoncbcs18/index.php"))
        self.non_cbcs_usn_start_label.setText(_translate("MainWindow", "Starting USN:"))
        self.non_cbcs_usn_end_label.setText(_translate("MainWindow", "Ending USN:"))
        self.non_cbcs_file_path_label.setText(_translate("MainWindow", "File Path/File Name:"))
        self.non_cbcs_file_path_lineEdit.setText(_translate("MainWindow", 'NON_CBCS_Result.xlsx'))
        self.non_cbcs_file_browse_toolButton.setText(_translate("MainWindow", "Browse"))
        self.non_cbcs_download_toolButton.setText(_translate("MainWindow", "Download"))
        self.non_cbcs_progressBar.setFormat(_translate("MainWindow", "%p%"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.non_cbcs_tab), _translate("MainWindow", "NON-CBCS"))
        self.cbcs_url_label.setText(_translate("MainWindow", "Result URL:"))
        self.cbcs_url_lineEdit.setText(_translate("MainWindow", "http://results.vtu.ac.in/vitaviresultcbcs2018/index.php"))
        self.cbcs_usn_start_label.setText(_translate("MainWindow", "Starting USN:"))
        self.cbcs_usn_end_label.setText(_translate("MainWindow", "Ending USN:"))
        self.cbcs_file_path_label.setText(_translate("MainWindow", "File Path/File Name:"))
        self.cbcs_file_path_lineEdit.setText(_translate("MainWindow", 'CBCS_Result.xlsx'))
        self.cbcs_file_browse_toolButton.setText(_translate("MainWindow", "Browse"))
        self.cbcs_download_toolButton.setText(_translate("MainWindow", "Download"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.cbcs_tab), _translate("MainWindow", "CBCS"))
        self.label.setText(_translate("MainWindow",
                                        "<html><head/><body><p>Designed and Developed by <a href=\"https://www.linkedin.com/in/animikh-aich/\"><span style=\" text-decoration: underline; color:#0000ff;\">Animikh Aich</span></a><br/>Dept. of ECE, <a href=\"http://rnsit.ac.in/\"><span style=\" text-decoration: underline; color:#0000ff;\">RNS Institute of Technology</span></a></p></body></html>"))



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

