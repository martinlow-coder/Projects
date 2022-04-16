import subprocess, os, platform
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog


# Need to sort issue with excel file layout
        # this may just be Ryans Layout
        # does combo box need a limit with a scrollbar?
# Need to create option for selecting all enteries from column
        # if col not "":
# consider splash screen
# remove frame from splash screen

class Ui_FilterWindow(object): # Pass in widget next time to avoid filedialog issue, or set file dialog parent to None instead of self.

# DEFINE PROGRAM FUNCTIONS HERE

    def combo_ls(self):
        '''creates a list of column names in the comboBox'''
        file_n = self.lineEdit_inputFile.text()
        # if sheet name = None, returns sheet names 
        # if sheet name = index, returns col names
        try:
                combo_df = pd.read_excel(file_n, sheet_name= None, keep_default_na=False)
        except:
                return
        
        sheet_qty = len(combo_df)
        global combo_ls
        combo_ls = [] # list for all column names
        combo_ls_clean = [] # list with duplicate column names removed
        # for each sheet
        for i in range(sheet_qty):
                combo_df = pd.read_excel(file_n, sheet_name= i, keep_default_na=False)
                # add each column name adding them to a list
                for j in combo_df.columns:
                        combo_ls.append(j)
        # sort in alphabetical order
        combo_ls = sorted(combo_ls)
        # remove duplicates from list
        for k in combo_ls:
                if k not in combo_ls_clean:
                        combo_ls_clean.append(k)
        combo_ls = combo_ls_clean
        self.comboBox.addItems(combo_ls)
        self.radioButton.setDisabled(True)
        
    def fill_colName(self):
        '''fills the colName lineEdit witth the selected item from the comboBox'''
        self.lineEdit_colName.setText(self.comboBox.currentText())

    def col_index(self, col_n, file_n):
        '''Finds the index of the column
        provided a name by the user'''
        df = pd.read_excel(file_n, keep_default_na=False)
        ls = list(df.columns)
        col_id = ls.index(col_n)
        return col_id

    def create(self):
        '''takes the input data and outputs a new file'''
        file_name = self.lineEdit_inputFile.text()
        try:
                df1 = pd.read_excel(file_name, sheet_name = None, keep_default_na= False)
        except:
                self.lineEdit_inputFile.setText("Don't be a Nitty! Get the fucking file! --->")
                return
        col_name = self.col_index(self.lineEdit_colName.text(), file_name)
        key_word = self.lineEdit_keyWord.text()
        #new_name = self.lineEdit_output.text() # remove later along with lineEdit
        '''Searches the columns of a spread sheet 
        in a file, for defined input,
        outputs a new file with all the found data'''
        sheets = len(df1) # how many sheets in file
        df2 = [] # list for found data to be appended
        # iterate through each sheet and append to list
        for i in range(sheets):
                col_ls = []
                sheet = pd.read_excel(file_name, sheet_name = i, keep_default_na=False)
                # iterate the length of each sheet checking if item in column or letter is in column (case sensitive 1st letter removed)
                for j in range(len(sheet)):
                        var = sheet.iloc[j, col_name]
                        vars = str(var)
                        if key_word in vars:
                                df2.append(sheet.iloc[j])
                        elif len(key_word) == 1 and key_word[0].upper() in vars[0]:
                                df2.append(sheet.iloc[j])
                        elif len(key_word) == 1 and key_word[0].lower() in vars[0]:
                                df2.append(sheet.iloc[j])
        df3 = pd.DataFrame(df2)
        if not df3.empty:
                pass
        else:
                self.lineEdit_keyWord.setText("Word / Letter not in column you Nitty")
                return
        global newFName
        newFName, _ = QFileDialog.getSaveFileName(None, "Save File", "", "Excel File (*.xlsx)")
        if newFName:
                df3.to_excel(newFName) # write to a new excel file give a name and extention added. 
        self.radioButton.setDisabled(False) # disabled to stop duplicate input (reset on create())
        self.comboBox.clear()

    def openFile(self):
        '''opens the file after creating it
        depending on platform, if new file has been named (not cancelled)'''
        try:
                if newFName:
                        if platform.system() == 'Darwin':       # macOS
                                subprocess.call(('open', newFName))
                        elif platform.system() == 'Windows':    # Windows
                                os.startfile(newFName)
                        else:                                   # linux variants
                                subprocess.call(('xdg-open', newFName))
        except:
                self.lineEdit_inputFile.setText("NOPE! try again KnobHead!        ----->")
                return
        
        
    def reset(self):
        ''' reset the input fields to start over'''
        self.lineEdit_inputFile.setText("")
        self.lineEdit_colName.setText("")
        self.comboBox.clear()
        self.lineEdit_keyWord.setText("")
        self.lineEdit_output.setText("")
        self.radioButton.setDisabled(False)

    def search(self):
        #Open the file dialog to select file
        fname, _  = QFileDialog.getOpenFileName(None, "Open File", "", "Excel Files (*.xlsx)")
        # output file to screen if fname is selected
        if fname:
                self.lineEdit_inputFile.setText(fname)

# MAIN USER INTERFACE IS ASSEMBLED HERE

    def setupUi(self, FilterWindow):
        FilterWindow.setObjectName("FilterWindow")
        FilterWindow.resize(658, 470) # set the main window size 
        FilterWindow.setStyleSheet("QMainWindow{\n"
"\n"
"background-color: rgb(56, 58, 91);}")
        self.centralwidget = QtWidgets.QWidget(FilterWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(0, 0, 0);\n"
"    border-radius: 15px;\n"
"}")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.label_title = QtWidgets.QLabel(self.frame)
        self.label_title.setGeometry(QtCore.QRect(0, 100, 621, 41))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(14)
        self.label_title.setFont(font)
        self.label_title.setStyleSheet("color: rgb(245, 109, 255);")
        self.label_title.setAlignment(QtCore.Qt.AlignCenter)
        self.label_title.setObjectName("label_title")
        self.labe_column = QtWidgets.QLabel(self.frame)
        self.labe_column.setGeometry(QtCore.QRect(30, 180, 101, 20))
        self.labe_column.setStyleSheet("color: rgb(255, 255, 255);")
        self.labe_column.setObjectName("labe_column")
        self.label_keyWord = QtWidgets.QLabel(self.frame)
        self.label_keyWord.setGeometry(QtCore.QRect(30, 290, 101, 20))
        self.label_keyWord.setStyleSheet("color: rgb(255, 255, 255);")
        self.label_keyWord.setObjectName("label_keyWord")
        self.lineEdit_colName = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_colName.setGeometry(QtCore.QRect(140, 180, 370, 21))
        self.lineEdit_colName.setStyleSheet("QLineEdit{\n"
"    color: rgb(0, 0, 0);    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 5px;\n"
"}")
        self.lineEdit_colName.setObjectName("lineEdit_colName")
        self.lineEdit_keyWord = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_keyWord.setGeometry(QtCore.QRect(140, 290, 370, 21))
        self.lineEdit_keyWord.setStyleSheet("QLineEdit{\n"
"    color: rgb(0, 0, 0);    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 5px;\n"
"}")
        self.lineEdit_keyWord.setObjectName("lineEdit_keyWord")
        self.label_outputFile = QtWidgets.QLabel(self.frame)
        self.label_outputFile.setGeometry(QtCore.QRect(200, 330, 370, 30))
        #self.label_outputFile.setGeometry(QtCore.QRect(30, 330, 101, 16))
        self.label_outputFile.setStyleSheet("color: rgb(255, 255, 255);")
        self.label_outputFile.setObjectName("label_outputFile")
        #self.label_outputFile.hide()
        self.lineEdit_output = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_output.setGeometry(QtCore.QRect(140, 330, 370, 21))
        self.lineEdit_output.hide()
        self.lineEdit_output.setStyleSheet("QLineEdit{\n"
"    color: rgb(0, 0, 0);    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 5px;\n"
"}")
        self.lineEdit_output.setObjectName("lineEdit_output")
        self.labe_inputFile = QtWidgets.QLabel(self.frame)
        self.labe_inputFile.setGeometry(QtCore.QRect(30, 140, 101, 20))
        self.labe_inputFile.setStyleSheet("color: rgb(255, 255, 255);")
        self.labe_inputFile.setObjectName("labe_inputFile")
        self.lineEdit_inputFile = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_inputFile.setGeometry(QtCore.QRect(140, 140, 261, 21))
        self.lineEdit_inputFile.setStyleSheet("QLineEdit{\n"
"    color: rgb(0, 0, 0);    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 5px;\n"
"}")
        self.lineEdit_inputFile.setObjectName("lineEdit_inputFile")
        self.pushButton_create = QtWidgets.QPushButton(self.frame, clicked= lambda: self.create()) # calls the create function (may change for dialog open file)
        self.pushButton_create.setGeometry(QtCore.QRect(140, 370, 100, 25))
        self.pushButton_create.setStyleSheet("QPushButton{\n"
"    color: rgb(0, 0, 0);\n"
"    background-color: rgb(8, 255, 21);\n"
"    border-radius: 10px;\n"
"}")
        self.pushButton_create.setObjectName("pushButton_create")

        self.comboBox = QtWidgets.QComboBox(self.frame)
        self.comboBox.setGeometry(QtCore.QRect(134, 250, 381, 32))
        self.comboBox.setAutoFillBackground(False)
        # added the abstractItemView in the CSS below for the background-color of the menu and the selection highlight color
        self.comboBox.setStyleSheet("QComboBox{\n"
"    color: rgb(0, 0, 0);\n"
"    \n"
"}"

"QComboBox QAbstractItemView {\n"
"        border: 2px solid darkgray;\n"
"        selection-background-color: rgb(56, 255, 84);\n"
"        background-color: white;\n"
"}")
        self.comboBox.setMaxVisibleItems(100)
        self.comboBox.setObjectName("comboBox")

        self.label_colNames = QtWidgets.QLabel(self.frame)
        self.label_colNames.setGeometry(QtCore.QRect(30, 255, 101, 16))
        self.label_colNames.setStyleSheet("color: rgb(255, 255, 255);")
        self.label_colNames.setObjectName("label_colNames")
        self.pushButton_createOpen = QtWidgets.QPushButton(self.frame, clicked = lambda: [self.create(), self.openFile()]) # does the same as create, but also open the new file
        self.pushButton_createOpen.setGeometry(QtCore.QRect(250, 370, 141, 25))
        self.pushButton_createOpen.setStyleSheet("QPushButton{\n"
"    color: rgb(0, 0, 0);\n"
"    background-color: rgb(8, 255, 21);\n"
"    border-radius: 10px;\n"
"}")
        self.pushButton_createOpen.setObjectName("pushButton_createOpen")

        self.label_credits = QtWidgets.QLabel(self.frame) # this isn't showing for some reason ???
        self.label_credits.setGeometry(QtCore.QRect(-3, 410, 631, 31))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.label_credits.setFont(font)
        self.label_credits.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_credits.setObjectName("label_credits")

        self.radioButton = QtWidgets.QRadioButton(self.frame, clicked = lambda: self.combo_ls()) # creates the list when the radio is selected
        self.comboBox.activated[str].connect(self.fill_colName) #calls the function if an item is chosen from the list)
        self.radioButton.setGeometry(QtCore.QRect(140, 218, 251, 20))
        self.radioButton.setStyleSheet("QRadioButton{\n"
"    color: rgb(255, 255, 255);\n"
"\n"
"}")
        self.radioButton.setObjectName("radioButton")
        self.pushButton_search = QtWidgets.QPushButton(self.frame, clicked = lambda: self.search()) #  opens fileOpenDialog when clicked, then set lineEdit to selection
        self.pushButton_search.setGeometry(QtCore.QRect(410, 140, 100, 20))
        self.pushButton_search.setStyleSheet("QPushButton{\n"
"    color: rgb(0, 0, 0);\n"
"    background-color: rgb(8, 255, 21);\n"
"    border-radius: 10px;\n"
"}")
        self.pushButton_search.setObjectName("pushButton_search")

        self.pushButton_reset = QtWidgets.QPushButton(self.frame, clicked = lambda: self.reset()) # need to reset radio buttton when clicked (can clear lineEdits too)
        self.pushButton_reset.setGeometry(QtCore.QRect(398, 370, 113, 25))
        self.pushButton_reset.setStyleSheet("QPushButton{\n"
"    color: rgb(0, 0, 0);\n"
"    \n"
"    background-color: rgb(255, 34, 20);\n"
"    border-radius: 10px;\n"
"}")
        self.pushButton_reset.setObjectName("pushButton_reset")

        self.label_image = QtWidgets.QLabel(self.frame)
        self.label_image.setGeometry(QtCore.QRect(80, -100, 481, 321))
        self.label_image.setText("")
        self.label_image.setPixmap(QtGui.QPixmap("EverydaySYNCTextinv.png")) # not working ??? Update (removed image file path and gave only file name. works now)
        self.label_image.setScaledContents(True)
        self.label_image.setAlignment(QtCore.Qt.AlignCenter)
        self.label_image.setObjectName("label_image")
        self.label_image.raise_()

        self.labe_column.raise_()
        self.label_keyWord.raise_()
        self.lineEdit_colName.raise_()
        self.lineEdit_keyWord.raise_()
        self.label_outputFile.raise_()
        self.lineEdit_output.raise_()
        self.labe_inputFile.raise_()
        self.lineEdit_inputFile.raise_()
        self.pushButton_create.raise_()

        self.comboBox.raise_()

        self.label_colNames.raise_()
        self.pushButton_createOpen.raise_()
        self.label_credits.raise_()
        self.radioButton.raise_()
        self.pushButton_search.raise_()
        self.pushButton_reset.raise_()
        self.label_title.raise_()
        self.verticalLayout.addWidget(self.frame)
        FilterWindow.setCentralWidget(self.centralwidget)
        self.actionHow_to_use_the_filter = QtWidgets.QAction(FilterWindow)
        self.actionHow_to_use_the_filter.setObjectName("actionHow_to_use_the_filter")
        self.actionSearch_file = QtWidgets.QAction(FilterWindow)
        self.actionSearch_file.setObjectName("actionSearch_file")

        self.retranslateUi(FilterWindow)
        QtCore.QMetaObject.connectSlotsByName(FilterWindow)
        FilterWindow.setTabOrder(self.lineEdit_inputFile, self.lineEdit_colName)
        FilterWindow.setTabOrder(self.lineEdit_colName, self.lineEdit_keyWord)
        FilterWindow.setTabOrder(self.lineEdit_keyWord, self.lineEdit_output)
        FilterWindow.setTabOrder(self.lineEdit_output, self.pushButton_create)

# DEFINE TEXT TO BE TRASLATEABLE

    def retranslateUi(self, FilterWindow):
        _translate = QtCore.QCoreApplication.translate
        FilterWindow.setWindowTitle(_translate("FilterWindow", "FilterWindow"))
        self.label_title.setText(_translate("FilterWindow", "<strong>EXCEL</strong> FILTER"))
        self.labe_column.setText(_translate("FilterWindow", "Name of Column"))
        self.label_keyWord.setText(_translate("FilterWindow", "Search key word"))
        self.lineEdit_colName.setToolTip(_translate("FilterWindow", "<html><head/><body><p>Enter the name of the column you want to filter data from.</p><p><span style=\" color:#ff2600;\">File name must not contain (&quot;/&quot;)</span></p></body></html>"))
        self.lineEdit_colName.setPlaceholderText(_translate("FilterWindow", "eg: Client or choose from the selection below."))
        self.lineEdit_keyWord.setToolTip(_translate("FilterWindow", "Enter the key word you're looking for in the column."))
        self.lineEdit_keyWord.setPlaceholderText(_translate("FilterWindow", "eg: UFC"))
        self.label_outputFile.setText(_translate("FilterWindow", "NOW CLICK A THING TO DO A THING!"))
        self.lineEdit_output.setToolTip(_translate("FilterWindow", "Enter a name for the new xlsx file."))
        self.lineEdit_output.setPlaceholderText(_translate("FilterWindow", "eg: March_Contacted_data"))
        self.labe_inputFile.setText(_translate("FilterWindow", "Input file name"))
        self.lineEdit_inputFile.setToolTip(_translate("FilterWindow", "Enter the name or path of the file you want to filter data from"))
        self.lineEdit_inputFile.setWhatsThis(_translate("FilterWindow", "<html><head/><body><p><br/></p></body></html>"))
        self.lineEdit_inputFile.setPlaceholderText(_translate("FilterWindow", "eg: Client.xlsx or /Users/Name/Desktop/Folder/Client.xlsx"))
        self.pushButton_create.setToolTip(_translate("FilterWindow", "Creates the new file."))
        self.pushButton_create.setText(_translate("FilterWindow", "Create file"))
        self.comboBox.setToolTip(_translate("FilterWindow", "Displays all the column names from your spreadsheets."))
        self.comboBox.setPlaceholderText(_translate("FilterWindow", "Column names will appear here."))
        self.label_colNames.setText(_translate("FilterWindow", "Columns names"))
        self.pushButton_createOpen.setToolTip(_translate("FilterWindow", "Creates a new file and opens it."))
        self.pushButton_createOpen.setText(_translate("FilterWindow", "Create and Open file"))
        self.label_credits.setText(_translate("FilterWindow", "<strong>By</strong>: MC Design"))
        self.radioButton.setText(_translate("FilterWindow", "Generate the column names below"))
        self.pushButton_search.setText(_translate("FilterWindow", "Search"))
        self.pushButton_reset.setToolTip(_translate("FilterWindow", "Rest program"))
        self.pushButton_reset.setText(_translate("FilterWindow", "Reset"))
        self.actionHow_to_use_the_filter.setText(_translate("FilterWindow", "How to use the filter"))
        self.actionSearch_file.setText(_translate("FilterWindow", "Search file"))

# RUN MAIN PROGRAM 

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    FilterWindow = QtWidgets.QMainWindow()
    ui = Ui_FilterWindow()
    ui.setupUi(FilterWindow)
    FilterWindow.show()
    sys.exit(app.exec_())