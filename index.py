from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sqlite3
import sys
from PyQt5.uic import loadUiType
import datetime
from xlrd import *
from xlsxwriter import *

ui,_ = loadUiType('libary.ui')
login,_ = loadUiType('login.ui')

class Login(QWidget, login):
    def __init__(self):
        QWidget.__init__(self)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.Handle_Login)

        style = open('themes/darkorange.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Handle_Login(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        username = self.lineEdit.text()
        password = self.lineEdit_2.text()

        sql = '''SELECT * FROM users'''
        self.cur.execute(sql)
        data = self.cur.fetchall()
        for row in data:
            if username == row[1] and password == row[3]:
                print('user match')
                self.window2 = MainApp()
                self.close()
                self.window2.show()

            else:
                self.label.setText('Please enter username and  password correctly')



class MainApp(QMainWindow, ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.Handle_UI_Changes()
        self.Handle_Buttons()

        self.Show_Category()
        self.Show_Author()
        self.Show_Publisher()

        self.Show_All_Client()
        self.Show_All_Books()
        self.Show_All_Operations()

        self.Show_Category_ComboBox()
        self.Show_Author_ComboBox()
        self.Show_Publisher_ComboBox()
        self.Dark_Orange_Theme()



    def Handle_UI_Changes(self):
        self.Hide_Themes()
        self.tabWidget.tabBar().setVisible(False)

    def Handle_Buttons(self):
        self.pushButton_5.clicked.connect(self.Show_Themes)
        self.pushButton_21.clicked.connect(self.Hide_Themes)


        self.pushButton.clicked.connect(self.Open_Day_To_Day_Tab)
        self.pushButton_2.clicked.connect(self.Open_Books_Tab)
        self.pushButton_28.clicked.connect(self.Open_Clients_Tab)
        self.pushButton_3.clicked.connect(self.Open_Users_Tab)
        self.pushButton_4.clicked.connect(self.Open_Settings_Tab)

        self.pushButton_7.clicked.connect(self.Add_New_Books)
        self.pushButton_10.clicked.connect(self.Search_Books)
        self.pushButton_8.clicked.connect(self.Edit_Books)
        self.pushButton_9.clicked.connect(self.Delete_Books)

        self.pushButton_14.clicked.connect(self.Add_category)
        self.pushButton_15.clicked.connect(self.Add_Author)
        self.pushButton_16.clicked.connect(self.Add_Publisher)

        self.pushButton_24.clicked.connect(self.Add_New_Client)
        self.pushButton_27.clicked.connect(self.Search_Client)
        self.pushButton_25.clicked.connect(self.Edit_Client)
        self.pushButton_26.clicked.connect(self.Delete_Client)

        self.pushButton_11.clicked.connect(self.Add_New_User)
        self.pushButton_12.clicked.connect(self.Login)
        self.pushButton_13.clicked.connect(self.Edit_User)

        self.pushButton_17.clicked.connect(self.Dark_Blue_Theme)
        self.pushButton_18.clicked.connect(self.Light_Blue_Theme)
        self.pushButton_20.clicked.connect(self.Dark_Orange_Theme)
        self.pushButton_19.clicked.connect(self.Dark_Gray_Theme)
        self.pushButton_22.clicked.connect(self.Light_Theme)
        self.pushButton_23.clicked.connect(self.Orange_Theme)

        self.pushButton_6.clicked.connect(self.Handle_Day_Operations)

        self.pushButton_29.clicked.connect(self.Export_Day_Operation)
        self.pushButton_30.clicked.connect(self.Export_Books)
        self.pushButton_31.clicked.connect(self.Export_Clients)


    def Show_Themes(self):
        self.groupBox_3.show()

    def Hide_Themes(self):
        self.groupBox_3.hide()



    ##############################
    ######### Opening tabs #######

    def Open_Day_To_Day_Tab(self):
        self.tabWidget.setCurrentIndex(0)

    def Open_Books_Tab(self):
        self.tabWidget.setCurrentIndex(1)

    def Open_Clients_Tab(self):
        self.tabWidget.setCurrentIndex(2)

    def Open_Users_Tab(self):
        self.tabWidget.setCurrentIndex(3)

    def Open_Settings_Tab(self):
        self.tabWidget.setCurrentIndex(4)

     ###########################
    ######### Day Operations #######

    def Handle_Day_Operations(self):
        book_title = self.lineEdit.text()
        client_name = self.lineEdit_29.text()
        type = self.comboBox.currentText()
        days_number = self.comboBox_2.currentIndex() + 1
        today_date = datetime.date.today()
        to_date = today_date + datetime.timedelta(days=days_number)

        print(today_date)
        print(to_date)

        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute('''
            INSERT INTO day_operations(book_name, client, type, days, date, to_date)
            VALUES(?, ?, ?, ?, ?, ?)
        ''', (book_title,client_name, type, days_number, today_date, to_date))

        self.db.commit()
        self.statusBar().showMessage('New Operation Added')
        self.Show_All_Operations()

    def Show_All_Operations(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute('''
            SELECT book_name, client, type, date, to_date FROM day_operations
        ''')

        data = self.cur.fetchall()
        print((data))

        self.tableWidget.setRowCount(0)
        self.tableWidget.insertRow(0)
        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1

            row_position = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row_position)


    ##############################
    ######### Books #######

    def Show_All_Books(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute('SELECT book_code, book_name, book_description, book_category, book_author, book_publisher, book_price FROM book')
        data = self.cur.fetchall()
        print(data)
        self.tableWidget_5.setRowCount(0)
        self.tableWidget_5.insertRow(0)

        for row_number, row_data in enumerate(data):
            for column, item in enumerate(row_data):
                self.tableWidget_5.setItem(row_number, column, QTableWidgetItem(str(item)))
                column =+ 1

            row_position = self.tableWidget_5.rowCount()
            self.tableWidget_5.insertRow(row_position)

        self.db.close()


    def Add_New_Books(self):

        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        book_title = self.lineEdit_2.text()
        book_description = self.textEdit.toPlainText()
        book_code = self.lineEdit_3.text()
        book_category = self.comboBox_3.currentText()
        book_author = self.comboBox_4.currentText()
        book_publisher = self.comboBox_5.currentText()
        book_price = self.lineEdit_4.text()

        self.cur.execute('''
            INSERT INTO book(book_name, book_description, book_code, book_category, book_author, book_publisher, book_price)
            VALUES(?, ?, ?, ?, ?, ?, ?)
        ''', (book_title, book_description, book_code, book_category, book_author, book_publisher, book_price))

        self.db.commit()
        self.db.close()
        self.statusBar().showMessage('New Book Added')

        self.lineEdit_2.setText('')
        self.textEdit.setPlainText('')
        self.lineEdit_3.setText('')
        self.comboBox_3.setCurrentIndex(0)
        self.comboBox_4.setCurrentIndex(0)
        self.comboBox_5.setCurrentIndex(0)
        self.lineEdit_4.setText('')

        self.Show_All_Books()



    def Search_Books(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        book_title = self.lineEdit_8.text()

        sql = ''' SELECT * FROM book WHERE book_name = ?'''
        self.cur.execute(sql, [book_title])

        data = self.cur.fetchone()
        print(data)

        self.lineEdit_7.setText(data[1])
        self.textEdit_2.setPlainText(data[2])
        self.lineEdit_6.setText(data[3])
        self.comboBox_7.setCurrentText((data[4]))
        self.comboBox_6.setCurrentText(data[5])
        self.comboBox_8.setCurrentText(data[6])
        self.lineEdit_5.setText(str(data[7]))


    def Edit_Books(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        book_title = self.lineEdit_7.text()
        book_description = self.textEdit_2.toPlainText()
        book_code = self.lineEdit_6.text()
        book_category = self.comboBox_7.currentText()
        book_author = self.comboBox_6.currentText()
        book_publisher = self.comboBox_8.currentText()
        book_price = self.lineEdit_5.text()

        search_book_title = self.lineEdit_8.text()

        self.cur.execute(''' 
            UPDATE book SET book_name=?, book_description=?, book_code=?, book_category=?, book_author=?, book_publisher=?, book_price=? WHERE book_name=?
        ''', (book_title, book_description, book_code, book_category, book_author, book_publisher, book_price, search_book_title))

        self.db.commit()
        self.db.close()
        self.statusBar().showMessage('Book updated')
        self.Show_All_Books()


    def Delete_Books(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        book_title = self.lineEdit_8.text()

        warning = QMessageBox.warning(self, 'Delete Book', "Are you sure you want to delete this book", QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes:
            sql = '''DELETE FROM book WHERE book_name=?'''
            self.cur.execute(sql,[(book_title)])
            self.db.commit()
            self.statusBar().showMessage('Book deleted')

            self.lineEdit_8.setText('')
            self.lineEdit_7.setText('')
            self.textEdit_2.setPlainText('')
            self.lineEdit_6.setText('')
            self.comboBox_7.setCurrentIndex(0)
            self.comboBox_6.setCurrentIndex(0)
            self.comboBox_8.setCurrentIndex(0)
            self.lineEdit_5.setText('')

            self.Show_All_Books()

    ######### CLients #######

    def Show_All_Client(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute('SELECT client_name, client_email, client_nationalid FROM clients')
        data = self.cur.fetchall()
        print(data)
        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)

        for row_number, row_data in enumerate(data):
            for column, item in enumerate(row_data):
                self.tableWidget_6.setItem(row_number, column, QTableWidgetItem(str(item)))
                column =+ 1

            row_position = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_position)

        self.db.close()

    def Add_New_Client(self):
        client_name = self.lineEdit_22.text()
        client_email = self.lineEdit_23.text()
        client_nationalid = self.lineEdit_24.text()

        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute('''
            INSERT INTO clients( client_name, client_email, client_nationalid)
            VALUES(?, ?, ?)''', (client_name, client_email, client_nationalid))

        self.db.commit()
        self.db.close()
        self.statusBar().showMessage('New Client Added')

        self.Show_All_Client()

    def Search_Client(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()
        clien_nationalid = self.lineEdit_28.text()

        sql = '''SELECT * FROM clients WHERE client_nationalid=?'''
        self.cur.execute(sql, [(clien_nationalid)])
        data = self.cur.fetchone()
        print(data)

        self.lineEdit_27.setText(data[1])
        self.lineEdit_26.setText(data[2])
        self.lineEdit_25.setText(data[3])

    def Edit_Client(self):
        client_Original_Nationalid = self.lineEdit_28.text()
        client_name = self.lineEdit_27.text()
        client_email = self.lineEdit_26.text()
        client_national_id = self.lineEdit_25.text()

        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute(''' UPDATE clients SET client_name=?, client_email=?, client_nationalid=? WHERE client_nationalid=?
        ''',(client_name, client_email,client_national_id, client_Original_Nationalid))

        self.db.commit()
        self.db.close()
        self.statusBar().showMessage('Client Data Updated')

        self.Show_All_Client()

    def Delete_Client(self):
        client_Original_Nationalid = self.lineEdit_28.text()

        warning = QMessageBox.warning(self, 'Delete Client', 'Are you sure you want to delete this client', QMessageBox.Yes | QMessageBox.Cancel)

        if warning == QMessageBox.Yes:
            self.db = sqlite3.connect('libary.db')
            self.cur = self.db.cursor()

            sql = '''DELETE FROM clients WHERE client_nationalid=?'''
            self.cur.execute(sql, [(client_Original_Nationalid)])

            self.db.commit()
            self.db.close()
            self.statusBar().showMessage('Client Data Deleted')

            self.Show_All_Client()


    ##############################
    ######### Users #######
    def Add_New_User(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        username = self.lineEdit_9.text()
        email = self.lineEdit_10.text()
        password = self.lineEdit_11.text()
        password2 = self.lineEdit_12.text()

        if password  == password2:
            self.cur.execute('''
                INSERT INTO users(user_name, user_email, user_password)
                VALUES( ?, ?, ? )''', (username, email, password))

            self.db.commit()
            self.statusBar().showMessage('New User Added')

        else:
            self.label_30.setText('Please add a valid password twice')


   ###### Users #####

    def Login(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        username = self.lineEdit_13.text()
        password = self.lineEdit_14.text()

        sql = '''SELECT * FROM users'''
        self.cur.execute(sql)
        data = self.cur.fetchall()
        for row in data:
            if username == row[1] and password == row[3]:
                self.statusBar().showMessage('Valid Username & Password')
                self.groupBox_4.setEnabled(True)

                self.lineEdit_15.setText(row[1])
                self.lineEdit_16.setText(row[2])
                self.lineEdit_17.setText(row[3])



    # def showMessage(self, title, message):
    #     msg = QMessageBox()
    #     msg.setIcon(QMessageBox.Information)
    #     msg.setWindowTitle(title)
    #     msg.setText(message)
    #     msg.setStandardButtons(QMessageBox.Ok)
    #     msg.exec_()

    def Edit_User(self):

        username = self.lineEdit_15.text()
        email = self.lineEdit_16.text()
        password = self.lineEdit_17.text()
        password2 = self.lineEdit_18.text()

        original_name = self.lineEdit_13.text()

        if password == password2:
            self.db = sqlite3.connect('libary.db')
            self.cur = self.db.cursor()

            self.cur.execute('''
                UPDATE users SET user_name = ?, user_email= ?, user_password =? WHERE user_name=?
                ''', (username, email, password, original_name))

            self.db.commit()
            self.statusBar().showMessage('User Data Updated Successfully')
        else:
            self.label_32.setText("password doesn't match")


    ##############################
    ######### settings #######
    def Add_category(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        category_name = self.lineEdit_19.text()

        self.cur.execute('''
            INSERT INTO category (category_name) VALUES (?)
            ''', (category_name,))

        self.db.commit()
        self.statusBar().showMessage('New Category Successfully Added')
        self.lineEdit_19.setText('')
        self.Show_Category()
        self.Show_Category_ComboBox()

    def Show_Category(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT category_name FROM category''')
        data = self.cur.fetchall()

        print(data)

        if data:
            self.tableWidget_2.setRowCount(0)
            self.tableWidget_2.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_2.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1

                    row_postion = self.tableWidget_2.rowCount()
                    self.tableWidget_2.insertRow(row_postion)

    def Add_Author(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        author_name = self.lineEdit_20.text()

        self.cur.execute('''
            INSERT INTO authors (author_name) VALUES (?)
            ''', (author_name,))

        self.db.commit()
        self.lineEdit_20.setText('')
        self.statusBar().showMessage('New Author Successfully Added')
        self.lineEdit_20.setText('')
        self.Show_Author()
        self.Show_Author_ComboBox()


    def Show_Author(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT author_name FROM authors''')
        data = self.cur.fetchall()

        print(data)

        if data:
            self.tableWidget_3.setRowCount(0)
            self.tableWidget_3.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_3.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1

                    row_postion = self.tableWidget_3.rowCount()
                    self.tableWidget_3.insertRow(row_postion)

    def Add_Publisher(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        publisher_name = self.lineEdit_21.text()

        self.cur.execute('''
            INSERT INTO publishers (publisher_name) VALUES (?)
            ''', (publisher_name,))

        self.db.commit()
        self.lineEdit_21.setText('')
        self.statusBar().showMessage('New Publisher Successfully Added')
        self.lineEdit_21.setText('')
        self.Show_Publisher()
        self.Show_Publisher_ComboBox()


    def Show_Publisher(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT publisher_name FROM publishers''')
        data = self.cur.fetchall()

        print(data)

        if data:
            self.tableWidget_4.setRowCount(0)
            self.tableWidget_4.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_4.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1

                    row_postion = self.tableWidget_4.rowCount()
                    self.tableWidget_4.insertRow(row_postion)

    ######### Show settings data in UI #########

    def Show_Category_ComboBox(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT category_name FROM category''')
        data = self.cur.fetchall()

        self.comboBox_3.clear()
        for category in data:
            self.comboBox_3.addItem(category[0])
            self.comboBox_7.addItem(category[0])

    def Show_Author_ComboBox(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT author_name FROM authors''')
        data = self.cur.fetchall()

        self.comboBox_4.clear()
        for author in data:
            self.comboBox_4.addItem(author[0])
            self.comboBox_6.addItem(author[0])

    def Show_Publisher_ComboBox(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT publisher_name FROM publishers''')
        data = self.cur.fetchall()
        self.comboBox_5.clear()
        for puplisher in data:
            self.comboBox_5.addItem(puplisher[0])
            self.comboBox_8.addItem(puplisher[0])

    ######### Export Data #######

    def Export_Day_Operation(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute('''
            SELECT book_name, client, type, date, to_date FROM day_operations
        ''')
        data = self.cur.fetchall()

        wb = Workbook('dayoperations.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0,0, 'book title')
        sheet1.write(0,1, 'client name')
        sheet1.write(0,2, 'type')
        sheet1.write(0,3,'from - date')
        sheet1.write(0,4,'to - date')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        print('done')
        self.statusBar().showMessage('Daily Report Created Successfully')

    def Export_Books(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute('SELECT book_code, book_name, book_description, book_category, book_author, book_publisher, book_price FROM book')
        data = self.cur.fetchall()

        wb = Workbook('all_books.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0,0, 'Book Code')
        sheet1.write(0,1, 'Book Name')
        sheet1.write(0,2, 'Book Description')
        sheet1.write(0,3, 'Book Category')
        sheet1.write(0,4, 'Book Author')
        sheet1.write(0,5, 'Book Publisher')
        sheet1.write(0,6, 'Book price')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        print('done')
        self.statusBar().showMessage('Client Report Created Successfully')

    def Export_Clients(self):
        self.db = sqlite3.connect('libary.db')
        self.cur = self.db.cursor()

        self.cur.execute('SELECT client_name, client_email, client_nationalid FROM clients')
        data = self.cur.fetchall()

        wb = Workbook('all_clients.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0,0, 'Clients name')
        sheet1.write(0,1, 'client Email')
        sheet1.write(0,2, 'Client National ID')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        print('done')
        self.statusBar().showMessage('Client Report Created Successfully')

    ######### settings #######
    def Light_Blue_Theme(self):
        style = open('themes/lightblue.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Dark_Blue_Theme(self):
        style = open('themes/darkblue.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Dark_Gray_Theme(self):
        style = open('themes/darkgray.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Light_Theme(self):
        style = open('themes/light.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Dark_Orange_Theme(self):
        style = open('themes/darkorange.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Orange_Theme(self):
        style = open('themes/dark2.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

def main():
    app = QApplication(sys.argv)
    window = Login()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()
