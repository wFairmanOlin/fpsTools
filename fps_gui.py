# https://pythonprogramminglanguage.com/pyqt-combobox/
import sys
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QLabel, QGridLayout, QWidget, QComboBox, QPushButton, QFileDialog, QLineEdit, QCompleter
from PyQt5.QtCore import QSize, QRect    
import pandas as pd
import time
from docxtpl import DocxTemplate

data = pd.read_csv('fps_data.csv')
doc = DocxTemplate("fps_template.docx")
print(data)


# data = data.append({'company': 'Colgate University'}, ignore_index=True)
companies = data['company'].dropna()
streets =data['street'].dropna()

company = data['company'][2]
index = data.loc[data['company'] == company].index[0]
print('2 == ', index)
# df.loc[df['column_name'] == some_value]

# data.to_csv(path_or_buf='fps_data.csv',index=False)


  # 20 layout = QGridLayout(window)
  # 21 layout.addWidget(nameLabel, 0, 0)
  # 22 layout.addWidget(nameEdit, 0, 1)
  # 23 layout.addWidget(addressLabel, 1, 0)
  # 24 layout.addWidget(addressEdit, 1, 1)
  # 25 layout.setRowStretch(2, 1)

class myQComboBox(QComboBox):
    def __init__(self):
        QComboBox.__init__(self)

    def focusOutEvent(self, event):
        print('lost focus because ', event)  #companyNam
        print('current text: ',mainWin.companyName.currentText())
        indices = data.loc[data['company'] == mainWin.companyName.currentText()].index
        print(indices)
        if len(indices) > 0:
            index = indices[0]
            print('data index: ',index)
            print('street: ', data['street'][index])
            print('service: ', data['service'][index])
            mainWin.companyStreet.setCurrentText(data['street'][index])
        else:
            print('no data found for this name')



        

    def focusInEvent(self, event):
        print('got focus because ', event)


class ExampleWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.window = QWidget()

        self.window.setMinimumSize(QSize(400, 200))    
        self.window.setWindowTitle("FPS Automation (Project Holidays)") 
        self.layout = QGridLayout(self.window)



        # centralWidget = QWidget(self)          
        # self.setCentralWidget(centralWidget)   

        # Create combobox and add items.
        self.companyName = myQComboBox()
        self.companyName.setGeometry(QRect(0,0,100,50))
        self.companyName.setObjectName("company")
        self.companyName.setEditable(True)
        completer = QCompleter(data['company'])
        self.companyName.setCompleter(completer)

        print(self.companyName.isEditable())
        for i in companies:
            self.companyName.addItem(i)
        # self.companyName.setCurrentText('Olin College')

        self.companyStreet = QComboBox()
        self.companyStreet.setGeometry(QRect(0,0,100,100))
        self.companyStreet.setObjectName('street')
        completer = QCompleter(data['street'])
        self.companyStreet.setEditable(True)
        self.companyStreet.setCompleter(completer)

        for i in streets:
            self.companyStreet.addItem(i)
        # self.companyStreet.setCurrentText('1000 Olin way')

        self.button = QPushButton('create')

        self.fileSave = QLineEdit()


        self.button.clicked.connect(self.createDoc)
        # self.companyName.currentTextChanged.connect(self.fillWindow)
        # self.companyName.lostFocus.connect(self.isHighlighted)
        # self.setObjectName('create')

    
        self.layout.addWidget(QLabel('Company Name'),0,0)
        self.layout.addWidget(self.companyName, 1, 0)
        self.layout.addWidget(QLabel('Street Name'),2,0)
        self.layout.addWidget(self.companyStreet, 3, 0)
        self.layout.addWidget(self.fileSave,3,1)
        self.layout.addWidget(self.button,4,1)

        self.layout.setRowStretch(2,3)

    def createDoc(self, b):
        saveLocation =  QFileDialog.getSaveFileName()
        print(saveLocation[0])
        file = str(saveLocation[0]) + '.docx'
        print(file)
        print('Company: ', self.companyName.currentText())
        print('Street: ', self.companyStreet.currentText())
        company = self.companyName.currentText()
        street = self.companyStreet.currentText()
        context = {'company' : company, 'company_street' : street}
        doc.render(context)
        doc.save(file)

        print('document generated!')

    def fillWindow(self, b):
        print(self.companyName.currentIndex())
        print('form filled')

    def isHighlighted(self, b):
        print('highlighted!')

    def keyPressEvent(self, qKeyEvent):
        print(qKeyEvent.key())



if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWin = ExampleWindow()
    mainWin.window.show()

    sys.exit( app.exec_() )
