''' 
A simple Payroll Management Desktop App  to manage Salary Details of employees for an Institution.
User can register employees, Create Designations, Create payrolls  for each month, generate Payslips.
Creates payroll logs to monitor the app's activities by a user.
User can backup data and payroll logs by sending to an email.
    
'''

import sqlite3
import os
import sys
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtCore import QDate 
from openpyxl import Workbook, styles
import datetime 
from dateutil.relativedelta import relativedelta
from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import *
from reportlab.lib.utils import ImageReader
from reportlab.lib.pagesizes import A4
from reportlab.lib import pdfencrypt
import configparser
import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from PyQt5.QtCore import QTimer, QTime, QPoint, QRect


# crates a log file to store all logging data
logging.basicConfig(filename='payroll_logs.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

logging.info("System Started")

# get the absolute path to the root directory of the PyInstaller package
if hasattr(sys, '_MEIPASS'):
    # running in a PyInstaller bundle
    root_dir = sys._MEIPASS
else:
    # running in a normal Python environment
    root_dir = os.path.dirname(os.path.abspath(__file__))

# construct the path to the image file relative to the root directory
image_path = os.path.join(root_dir, 'drumvale_logo.jpg')

logo = ImageReader(image_path)

# Account_Number to create placeholder textboxes
class MyLineEdit(QtWidgets.QLineEdit):
    def __init__(self, placeholder_text, parent=None):
        super().__init__(parent)
        self.setPlaceholderText(placeholder_text)

    def focusInEvent(self, event):
        if self.text() == self.placeholderText():
            self.setPlaceholderText("")
        super().focusInEvent(event)


app = QtWidgets.QApplication(sys.argv)
app.setStyle("Fusion")

timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

timelbl = QtWidgets.QLabel()
timelbl.setText(timestamp)

# Create tHE MAIN Window
window = QtWidgets.QWidget()
window.setWindowTitle(f"PAYROLL MANAGEMENT SYSTEM....      Today is {timestamp}")

layout = QtWidgets.QVBoxLayout()
window.setLayout(layout)

## Creating payroll_id for the stated month and the timestamp for payroll creation
now = datetime.datetime.now()
month = now.strftime("%b").upper()
year = now.strftime("%Y")

payroll_id = month + "-" + year
timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

# creating  Validators for QLineEdits
number_validator = QtGui.QIntValidator()
name_validator = QtGui.QRegExpValidator(QtCore.QRegExp("[a-zA-Z]+(\\s[a-zA-Z]+)?"))
phone_number_validator = QtGui.QRegExpValidator(QtCore.QRegExp("^0\d{9}$"))
KRA_PIN_validator = QtGui.QRegExpValidator(QtCore.QRegExp("^[A-Z][A-Z0-9]{9}[A-Z0-9]$"))
bank_accountNumber_validator = QtGui.QRegExpValidator(QtCore.QRegExp("\d{18}$"))

# CREATING QLINEEDITS FOR USER INPUT
Name_txt = MyLineEdit("Enter Full Name")
ID_No_txt = MyLineEdit("Enter ID Number")
Account_Number_txt = MyLineEdit("Bank account number")
Designation_txt = MyLineEdit("Enter the Designation")
Contact_txt = MyLineEdit("Enter Mobile Number")
KRA_PIN_txt = MyLineEdit("Enter KRA PIN")
Search_box_txt = MyLineEdit("Search the system")
Search_payslip_txt = MyLineEdit("Type text to search")
Basic_salary_txt = MyLineEdit("Enter the basic salary")
Commuter_allowance_txt = MyLineEdit("Enter commuter allowance")
House_allowance_txt = MyLineEdit("Enter house allowance")
NSSF_Employee_txt = MyLineEdit("NSSF Employee")
PAYE_txt = MyLineEdit("Enter PAYE to be paid")
NHIF_txt = MyLineEdit("Enter NHIF")
Employer_contribution_txt = MyLineEdit("Enter employer contribution")
Payroll_id_txt = QtWidgets.QLineEdit(payroll_id)

# Resizing QlineEdits
for txt in [ID_No_txt, Name_txt, Account_Number_txt, Designation_txt, Contact_txt, KRA_PIN_txt, Search_payslip_txt, Basic_salary_txt, \
            Commuter_allowance_txt, House_allowance_txt, NSSF_Employee_txt, PAYE_txt, NHIF_txt, Employer_contribution_txt, Payroll_id_txt ]:
    txt.setFixedSize(QtCore.QSize(300, 25))

###  GIVING LINEDITS FUNCTIONALITIES
# Payroll_id_txt.setReadOnly(True)  # makes the text uneditable
ID_No_txt.setMaxLength(8)
Name_txt.setValidator(name_validator)
Contact_txt.setValidator(phone_number_validator)
KRA_PIN_txt.setValidator(KRA_PIN_validator)
Account_Number_txt.setMaxLength(18)
Account_Number_txt.setValidator(bank_accountNumber_validator)

for txt in [ID_No_txt, Basic_salary_txt, Commuter_allowance_txt, House_allowance_txt, NSSF_Employee_txt, PAYE_txt, NHIF_txt, Employer_contribution_txt]:
    txt.setValidator(number_validator) # ensures it only accepts numbers
    
# Create a QTreeWidget
tree = QtWidgets.QTreeWidget()
tree.setHeaderHidden(True)

## CREATING COMBOBOXES
# combo box to display Employee ID
Employee_No_Combo = QtWidgets.QComboBox()
Employee_No_Combo.setFixedSize(QtCore.QSize(300, 25))
Employee_No_Combo.setPlaceholderText("Select an option")

# combo box to display employee designation
Employee_Designation_Combo = QtWidgets.QComboBox()
Employee_Designation_Combo.setFixedSize(QtCore.QSize(300, 25))

# combo box to select Data
Employee_payslip_No_combo = QtWidgets.QComboBox()
Employee_payslip_No_combo.setFixedSize(QtCore.QSize(300, 25))


# Creating labels
name_lbl = QtWidgets.QLabel()
ID_No_lbl = QtWidgets.QLabel()
Designation_lbl = QtWidgets.QLabel()
Account_Number_lbl = QtWidgets.QLabel()
KRA_PIN_lbl = QtWidgets.QLabel()
Gross_pay_lbl = QtWidgets.QLabel()
Taxable_amount_lbl = QtWidgets.QLabel()
Net_pay_lbl = QtWidgets.QLabel()
greet_label = QtWidgets.QLabel()


# Create a QDateEdit widget
date_chooser = QtWidgets.QDateEdit()
date_chooser.setFixedSize(QtCore.QSize(300, 25))
date_chooser.setCalendarPopup(True)
date_chooser.setDate(QtCore.QDate.currentDate())
date_chooser.setMaximumDate(QtCore.QDate.currentDate())  # ensures date chosen is not future date

# CREATING BUTTONS
submit_button = QtWidgets.QPushButton("Submit")
Create_Payroll_btn = QtWidgets.QPushButton("Create New Payroll") # takes you to the payroll form
Regsiter_btn = QtWidgets.QPushButton("Register")
Designation_btn = QtWidgets.QPushButton("Add Designation") # when clicked it pops up the deisgnation form
Add_Designation_btn = QtWidgets.QPushButton("Save Designation")
Save_payroll_btn = QtWidgets.QPushButton("Save payroll")
Generate_payslip_btn = QtWidgets.QPushButton("Generate Payslip")
load_salary_info_btn = QtWidgets.QPushButton("Load Salary Info")
Delete_btn = QtWidgets.QPushButton("Delete")
load_Current_Employees_btn = QtWidgets.QPushButton("Current Employees")
load_Payroll_Data_btn = QtWidgets.QPushButton("View Payroll Data")
load_Designations_btn = QtWidgets.QPushButton("View Designations")
export_button = QtWidgets.QPushButton("Export")
Payslip_btn = QtWidgets.QPushButton("Payslips")
change_theme_btn = QtWidgets.QPushButton("Change Theme")
Update_btn = QtWidgets.QPushButton("Update")
search_btn = QtWidgets.QPushButton("Search")
backup_btn = QtWidgets.QPushButton("Backup")
all_payslips_btn = QtWidgets.QPushButton("All payslips")

# setting the sizes of the buttons and funtionalities
for btn in [submit_button, Create_Payroll_btn, Regsiter_btn, Designation_btn, Add_Designation_btn, \
            Save_payroll_btn, Generate_payslip_btn, load_salary_info_btn, Update_btn, \
            load_Current_Employees_btn, load_Payroll_Data_btn, load_Designations_btn, Payslip_btn, change_theme_btn, backup_btn, all_payslips_btn]:
    btn.setFixedSize(300, 25)

## CREATING FUNCTIONS 
# Load the color configuration file everytime the app loads
config = configparser.ConfigParser()
config.read('config.ini')

# Retrieve the last selected color from the configuration file
last_color = config.get('settings', 'color', fallback='white')

# Set the initial background color of the main window and group box to the last selected color
window.setStyleSheet(f"background-color: {last_color};")

# function to chnage the theme color of the window
def change_Theme():
    # Open the QColorDialog with the last selected color
    color = QtWidgets.QColorDialog.getColor(QtGui.QColor(last_color))

    if color.isValid():
        # Set the background color of the main window to the selected color
        window.setStyleSheet(f"background-color: {color.name()};")
        group_box.setStyleSheet(f"background-color: {color.name()};")

        logging.info(f"User changed the theme application to {color}")

        # Save the selected color to the configuration file
        if not config.has_section('settings'):
            config.add_section('settings')
        config.set('settings', 'color', color.name())
        with open('config.ini', 'w') as configfile:
            config.write(configfile)


# This function loads Employee Number/ID into the combobox for selection
def load_Employee_Number():
    
    conn = sqlite3.connect('Payroll.db')
    cursor = conn.cursor()

    # Retrieve data from the table
    cursor.execute("SELECT Employee_No FROM current_employees")
    records = cursor.fetchall()  

    # Add the data to the combobox
    for item in records:
        Employee_No_Combo.addItem(item[0])
        
    # connect the currentIndexChanged signal to a slot function
    Employee_No_Combo.currentIndexChanged.connect(populate_employee_info)

    #commit changes
    conn.commit()

    # Close the cursor and connection
    cursor.close()
    conn.close() 

# function to load all the employee nummbers that have been employed in that institution both current and those left
# def load_Employee_Number_Payslips():
    
#     conn = sqlite3.connect('Payroll.db')
#     cursor = conn.cursor()

#     # Retrieve data from the table
#     cursor.execute("SELECT employee_No FROM Registration")
#     records = cursor.fetchall()  

#     # Add the data to the combobox
#     for item in records:
#         Employee_payslip_No_combo.addItem(item[0])  # loads current employee numbers into the combobox

#     #commit changes
#     conn.commit()

#     # Close the cursor and connection
#     cursor.close()
#     conn.close() 

# load_Employee_Number_Payslips()

# Function to populate employee info based on Employee No selected
def populate_employee_info():
    selected_employee_No = Employee_No_Combo.currentText() # retrieves the selected employee number from the combobox

    conn = sqlite3.connect('Payroll.db')
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT ID_Number, name, designation, Account_Number, KRA_PIN FROM current_employees WHERE employee_No = ?", (selected_employee_No,))

        selected_info = cursor.fetchone()

        label_list = [(ID_No_lbl, 1), (name_lbl, 0), (Designation_lbl, 2), (Account_Number_lbl, 3), (KRA_PIN_lbl, 4)]

        if selected_info is not None:
            # Loop through the label_list and populate each label with the corresponding value from selected_info
            for label, column_index in label_list:
                value = selected_info[column_index]
                label.setText(value)
        else:
           
            # Clear the text of the labels if no data was retrieved from the database
            for label, column_index in label_list:
                label.setText("")
            # logging.info("User tried to load salary info an employee not in the previous payroll.")

    except Exception as e:
        # Handle the exception by displaying an error message
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText("An error occurred while retrieving employee information.")
        msg.setInformativeText(str(e))
        msg.setWindowTitle("Error")
        msg.exec_()
        logging.info("An error occurred while retrieving employee information.")

    conn.commit()
    cursor.close()
    conn.close()

## Function to submit data into the registration table
def submit_data():
    try:
        if len(ID_No_txt.text()) != 8: # checks if the length of the ID Number is exactly 8 digits
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("ID Number should have 8 digits.")
            msg.setWindowTitle("Invalid")
            msg.exec_()
            logging.error(f"User entered an invalid ID Number {ID_No_txt.text()}.")
            return
        if len(Name_txt.text()) < 5: # checks if the length of the ID Number is exactly 8 digits
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("Name is too short.")
            msg.setWindowTitle("Invalid")
            msg.exec_()
            logging.error(f"User entered invalid Name {Name_txt.text()}.")
            return
        if Employee_Designation_Combo.currentText() == "":
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("Please select a designation.")
            msg.setWindowTitle("Invalid")
            msg.exec_()
            logging.error(f"User did not select designation")
            return
        if len(KRA_PIN_txt.text()) != 11:
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("KRA PIN should be 11 characters long.")
            msg.setWindowTitle("Invalid")
            msg.exec_()
            logging.error(f"User entered an invalid KRA PIN {KRA_PIN_txt.text()} .")
            return
        if len(Account_Number_txt.text()) < 10:
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("the Bank Account Number is not valid.")
            msg.setWindowTitle("Invalid")
            msg.exec_()
            logging.error(f"User entered an invalid Bank Acc Number {Account_Number_txt.text()}.")
            return
        if len(Contact_txt.text()) != 10:
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("Mobile Number should have 10 digits.")
            msg.setWindowTitle("Invalid")
            msg.exec_()
            logging.error(f"User entered an invalid Mobile Number {Contact_txt.text()}.")
            return
        
        # Get the maximum employee number from the Registration table
        conn = sqlite3.connect("Payroll.db")
        c = conn.cursor()
        c.execute("SELECT MAX(employee_NO) FROM Registration")
        max_employee_no = c.fetchone()[0]
        conn.close()

        # Generate the new employee number by adding 1 to the maximum employee number
        new_employee_no = generate_employee_number(max_employee_no)

        logging.info(f"New employee number {new_employee_no} was generated.")

        # Connect to the database
        conn = sqlite3.connect('Payroll.db')
        cursor = conn.cursor()

        # Create the table if it doesn't exist
        cursor.execute('''CREATE TABLE IF NOT EXISTS Registration (Employee_NO, 
                                                                Name text, 
                                                                ID_Number INT,
                                                                Designation text, 
                                                                Account_Number text, 
                                                                Date_of_Joining text,
                                                                KRA_PIN text,
                                                                Contact text
                                                                )''')
       
        # Create the current_employees table if it doesn't exist
        cursor.execute('''CREATE TABLE IF NOT EXISTS current_employees (Employee_NO, 
                                                                        Name text, 
                                                                        ID_Number text PRIMARY KEY,
                                                                        Designation text, 
                                                                        Account_Number text, 
                                                                        Date_of_Joining text,
                                                                        KRA_PIN text,
                                                                        Contact text
                                                                        )''')
        
        # checks if ID Number already exists in the DB
        cursor.execute("SELECT COUNT(*) FROM current_employees WHERE ID_Number=?", (ID_No_txt.text(),))
        result = cursor.fetchone()[0]

        id_result = ID_No_txt.text()

        if result > 0:
            # Id Number already exists, so don't insert it again
            message_box = QtWidgets.QMessageBox()
            message_box.setWindowTitle("Registration Error")
            message_box.setText(f"ID Number <b>{id_result}</b> Already Exists in the System")
            message_box.exec_()
            ID_No_txt.setFocus()
            logging.error(f"User entered ID Number {id_result} that already existt in the system.")
            return
        else:

            # Insert data into the Registration table
            cursor.execute("INSERT INTO Registration VALUES (?, ?, ?, ?, ?, ?, ?, ?)", (new_employee_no, 
                                                                                        Name_txt.text(),
                                                                                        ID_No_txt.text(), 
                                                                                        Employee_Designation_Combo.currentText(),
                                                                                        Account_Number_txt.text(), 
                                                                                        date_chooser.text(),
                                                                                        KRA_PIN_txt.text(),
                                                                                        Contact_txt.text()                                                                    
                                                                                        ))

            # Insert data into the current_employees table
            cursor.execute("INSERT INTO current_employees VALUES (?, ?, ?, ?, ?, ?, ?, ?)", (new_employee_no, 
                                                                                            Name_txt.text(),
                                                                                            ID_No_txt.text(), 
                                                                                            Employee_Designation_Combo.currentText(),
                                                                                            Account_Number_txt.text(), 
                                                                                            date_chooser.text(),
                                                                                            KRA_PIN_txt.text(),
                                                                                            Contact_txt.text()                                                                    
                                                                                            ))

            # Save changes to the database
            conn.commit()

            cursor.close()

            # Close the connection
            conn.close()

            message_box = QtWidgets.QMessageBox()
            message_box.setText(f"Employee Number <b>{new_employee_no}</b> Registered Succefully.")
            message_box.exec_()
            logging.info(f"New Employee Number {new_employee_no}  was inserted in the DB successfully.")


            tree.clear() #clears the Treeview of all the data

            load_current_employees()

            Employee_No_Combo.clear()  #clears the combobox data first before refreshing

            Employee_payslip_No_combo.clear()

            load_Employee_Number() # refreshes the data in the combobox after submitting data

            # Clearing the textboxes after submission
            textboxes =  [ID_No_txt, Name_txt, Account_Number_txt, KRA_PIN_txt, Contact_txt]
            for tb in textboxes:
                tb.clear()

            ID_No_txt.setFocus() # sets the cursor to the ID line edit after submitting

    except sqlite3.Error as error:
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("An error occurred while submitting data")
            msg.setInformativeText(str(error))
            msg.setWindowTitle("Error")
            msg.exec_()
            logging.critical(f"Error occurred while submitting data into The DB {error}")


# function populates selected data when double clicked to their respective lineedits
def double_Clicked():
    selected_items = tree.selectedItems()

    selected_item = selected_items[0] 

    tree.itemDoubleClicked.connect(lambda: Registration_widget.setVisible(True))

    ID_No_txt.setText(selected_item.text(2))
    Name_txt.setText(selected_item.text(1))
    Employee_Designation_Combo.setCurrentText(selected_item.text(3))
    Account_Number_txt.setText(selected_item.text(4))
    selected_date_str = selected_item.text(5)
    selected_date_format = "dd/MM/yyyy" # Replace with the actual format of selected_date_str
    selected_date = QDate.fromString(selected_date_str, selected_date_format)
    date_chooser.setDate(selected_date)
    KRA_PIN_txt.setText(selected_item.text(6))
    Contact_txt.setText(selected_item.text(7))

    Registration_widget.setEnabled(True)

# makes registration form to appear for every double click  
def make_Registration_Form_Appear():
    double_Clicked()
    Registration_widget.close()

tree.itemDoubleClicked.connect(make_Registration_Form_Appear)

# updates data based on selected info
def update_data():
    selected_items = tree.selectedItems() # gets the selected row after double clicking.

    if not selected_items:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText("Please double click on a record to update.")
        msg.setWindowTitle("Invalid")
        msg.exec_()
        return

    selected_item = selected_items[0]  # checks the now selected row after double clicking

    employee_no = selected_item.text(0) # gets the selected item of each column of selected row

    try:
        if len(ID_No_txt.text()) != 8: # checks if the length of the ID Number is exactly 8 digits
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("ID Number should have 8 digits.")
            msg.setWindowTitle("Invalid")
            msg.exec_()
            return
        if len(Name_txt.text()) < 5: # checks if the length of the ID Number is exactly 8 digits
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("Name is too short.")
            msg.setWindowTitle("Invalid")
            msg.exec_()
            logging.error(f"User entered invalid Name {Name_txt.text()}.")
            return
        if len(Account_Number_txt.text()) < 10:
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("the Bank Account Number is not valid.")
            msg.setWindowTitle("Invalid")
            msg.exec_()
            logging.error(f"User entered an invalid Bank Acc Number {Account_Number_txt.text()}.")
            return
        if len(KRA_PIN_txt.text()) != 11:
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("KRA PIN should be 11 characters long.")
            msg.setWindowTitle("Invalid")
            msg.exec_()
            return
        if len(Contact_txt.text()) != 10:
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("Mobile Number should have 10 digits.")
            msg.setWindowTitle("Invalid")
            msg.exec_()
            return

        # Connect to the database
        conn = sqlite3.connect('Payroll.db')
        cursor = conn.cursor()

        # Update data in the Registration table
        cursor.execute("UPDATE Registration SET Name=?, ID=?, Designation=?, Account_Number=?, Date=?, KRA=?, Contact=? WHERE Employee_NO=?", (
                                                                                    Name_txt.text(),
                                                                                    ID_No_txt.text(), 
                                                                                    Employee_Designation_Combo.currentText(),
                                                                                    Account_Number_txt.text(), 
                                                                                    date_chooser.text(),
                                                                                    KRA_PIN_txt.text(),
                                                                                    Contact_txt.text(),
                                                                                    employee_no,
                                                                                    ))

        # Update data in the current_employees table
        cursor.execute("UPDATE current_employees SET Name=?, ID_Number=?, Designation=?, Account_Number=?, Date=?, KRA_PIN=?, Contact=? WHERE Employee_NO=?", (
                                                                                        Name_txt.text(),
                                                                                        ID_No_txt.text(), 
                                                                                        Employee_Designation_Combo.currentText(),
                                                                                        Account_Number_txt.text(), 
                                                                                        date_chooser.text(),
                                                                                        KRA_PIN_txt.text(),
                                                                                        Contact_txt.text(),
                                                                                        employee_no
                                                                                        ))
        
        
        # Save changes to the database
        conn.commit()
        cursor.close()
        conn.close()

        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText(f"{employee_no} record has been updated succesfully.")
        msg.setWindowTitle("Update")
        msg.exec_()
        logging.info(f"record of employee Number {employee_no} was updated")
    
        Registration_widget.close()  # closes the Registration form after updating
        
        load_current_employees()

    except Exception as e:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText("Error occurred while updating the record.")
        msg.setWindowTitle("Error")
        msg.exec_()
        logging.info(f"Error occured while updating record of employee Number {employee_no}.")

# function to generate employee number
def generate_employee_number(max_employee_no):
    if max_employee_no:
        # Extract the numeric portion of the maximum employee number
        max_employee_no = int(max_employee_no[1:])
    else:
        max_employee_no = 4  # Start from E05 if there are no records

    # Generate the new employee number by adding 1 to the maximum employee number
    return "E{:03d}".format(max_employee_no + 1)

# function to crate autocomplete suggestions
def setCompleterForColumn(textbox, column_name):
    conn = sqlite3.connect("Payroll.db")
    cursor = conn.cursor()
    
    # Retrieve distinct values from the specified column
    cursor.execute(f"SELECT DISTINCT {column_name} FROM Payroll_Data")
    values = [row[0] for row in cursor.fetchall()]
    
    # Create a completer using the list of values
    completer = QtWidgets.QCompleter(values)
    completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
    
    # Set the completer for the QLineEdit widget
    textbox.setCompleter(completer)
    
    cursor.close()
    conn.close()

# generates a password protected payslip based on searched month and employee number
def generate_Payslip():
    try:
        selected_employee_No = Employee_payslip_No_combo.currentText()
        search_month = Search_payslip_txt.text()
        
        conn = sqlite3.connect("Payroll.db")
        cursor = conn.cursor()

         # Retrieve the payroll data for the selected employee and month
        cursor.execute("SELECT * FROM Payroll_Data WHERE Employee_Number = ? AND (Payroll_ID) = ?", 
                        (selected_employee_No, search_month))  # Use = operator to search for exact match

        result = cursor.fetchone()
        
        if result is not None and search_month != "":
            # Create a new PDF document using the ReportLab library
            filename = f"{result[5]}_{search_month} payslip.pdf"  # Create filename using employee name and payroll period

            #Retrieve the employee's ID number
            employee_password = str(result[4]) # employees password based on their ID Number

            admin_password2 = "admin"  # administrators password that has access to all the employees payslip
            
            enc=pdfencrypt.StandardEncryption(employee_password, admin_password2, canPrint=1, canModify=0,canCopy=0, strength=40)
            
            x_margin, y_margin = 100, 800  # Set the top-left corner of the payslip
            payslip = canvas.Canvas(filename, pagesize=A4, encrypt=enc)
            payslip.translate(x_margin, y_margin)  # Move the origin to the top-left corner

            # Draw border around the payslip document
            payslip.rect(-20, -0, 250, -300)

            fontSize = 8

            # Define the layout and content of the payslip template
            payslip.setFont('Courier-Bold', 8)  # Set the font and font size for the heading
            
            payslip.drawImage(logo, 80, -30, 25, 25)

            payslip.drawString(50, -40, "DRUMVALE SECONDARY SCHOOL")
            payslip.setFont('Courier-Bold', 7) 
            payslip.drawString(0, -50, "P.O. BOX 99-00520 RUAI-NAIROBI, TEL: 0704 921 291")
            payslip.drawString(0, -60, "EMAIL: info@drumvalesecondary.com")
            payslip.drawString(50, -70, f"{result[2]} PAYSLIP")

            payslip.setFont('Courier', fontSize)  # Set the font and font size for the other text
            payslip.drawString(0, -80, "Employee Number:")
            payslip.drawString(120, -80, result[3])
            payslip.drawString(0, -90, "ID Number:")
            payslip.drawString(120, -90, str(result[4]))
            payslip.drawString(0, -100, "Name:")
            payslip.drawString(120, -100, result[5])
            payslip.drawString(0, -110, "Designation:")
            payslip.drawString(120, -110, result[6])
            payslip.drawString(0, -120, "Account Number:")
            payslip.drawString(120, -120, str(result[7]))
            payslip.drawString(0, -130, "KRA PIN:")
            payslip.drawString(120, -130, result[8])
            payslip.setFont('Courier-Bold', fontSize)
            payslip.drawString(0, -140, "Total Earnings....")
            payslip.setFont('Courier', fontSize)

            payslip.drawString(0, -150, "Basic Salary:")
            basic_salary = result[9]
            if basic_salary:
                basic_salary = float(basic_salary)
                if basic_salary != 0:
                    payslip.drawString(120, -150, "{:,.2f}".format(basic_salary))
                else:
                    payslip.drawString(120, -150, "0.00")

            payslip.drawString(0, -160, "Commuter Allowance:")
            Commuter_allowance = result[10]
            if Commuter_allowance:
                Commuter_allowance = float(Commuter_allowance)
                if Commuter_allowance != 0:
                    payslip.drawString(120, -160, "{:,.2f}".format(Commuter_allowance))
                else:
                    payslip.drawString(120, -160, "0.00")

            
            payslip.drawString(0, -170, "House Allowance:")
            House_allowance = result[11]
            if House_allowance:
                House_allowance = float(House_allowance)
                if House_allowance != 0:
                    payslip.drawString(120, -170, "{:,.2f}".format(House_allowance))
                else:
                    payslip.drawString(120, -170, "0.00")

            payslip.setFont('Courier-Bold', fontSize)
            payslip.drawString(0, -180, "Gross Pay:")
            Gross_Pay = result[12]
            if Gross_Pay:
                Gross_Pay = float(Gross_Pay)
                if Gross_Pay != 0:
                    payslip.drawString(120, -180, "{:,.2f}".format(Gross_Pay))
                else:
                    payslip.drawString(120, -180, "0.00")

            payslip.drawString(0, -190, "Deductions.....")
            payslip.setFont('Courier', fontSize) 

            payslip.drawString(0, -200, "NSSF:")
            NSSF = result[13]
            if NSSF:
                NSSF = float(NSSF)
                if NSSF != 0:
                    payslip.drawString(120, -200, "{:,.2f}".format(NSSF))
                else:
                    payslip.drawString(120, -200, "0.00")

            payslip.setFont('Courier-Bold', fontSize)

            payslip.drawString(0, -210, "Taxable Amount:")
            Taxable_Amount = result[14]
            if Taxable_Amount:
                Taxable_Amount = float(Taxable_Amount)
                if Taxable_Amount != 0:
                    payslip.drawString(120, -210, "{:,.2f}".format(Taxable_Amount))
                else:
                    payslip.drawString(120, -210, "0.00")

            payslip.setFont('Courier', fontSize) 

            payslip.drawString(0, -220, "PAYE:")
            PAYE = result[15]
            if PAYE:
                PAYE = float(PAYE)
                if PAYE != 0:
                    payslip.drawString(120, -220, "{:,.2f}".format(PAYE))
                else:
                    payslip.drawString(120, -220, "0.00")

            payslip.drawString(0, -230, "NHIF:")
            NHIF = result[16]
            if NHIF:
                NHIF = float(NHIF)
                if NHIF != 0:
                    payslip.drawString(120, -230, "{:,.2f}".format(NHIF))
                else:
                    payslip.drawString(120, -230, "0.00")

            payslip.setFont('Courier-Bold', fontSize)

            payslip.drawString(0, -240, "Net Pay:")
            Net_Pay = result[17]
            if Net_Pay:
                Net_Pay = float(Net_Pay)
                if Net_Pay != 0:
                    payslip.drawString(120, -240, "{:,.2f}".format(Net_Pay))
                else:
                    payslip.drawString(120, -240, "0.00")

            payslip.setFont('Courier', fontSize)

            payslip.drawString(0, -250, "Employer NSSF:")
            Employer_NSSF = result[18]
            if Employer_NSSF:
                Employer_NSSF = float(Employer_NSSF)
                if Employer_NSSF != 0:
                    payslip.drawString(120, -250, "{:,.2f}".format(Employer_NSSF))
                else:
                    payslip.drawString(120, -250, "0.00")

            payslip.setLineWidth(0.5)
            payslip.line(-20, -275, 230, -275)    # draws a straightline to separate the time stamp

            payslip.setFont('Courier-Oblique', 7)
            payslip.drawString(20, -285, "~~~~Generated at " + timestamp + "~~~~")
            payslip.drawString(20, -295, "~~~Contact Admin at +254753761880~~~")

            # Save the PDF document and display a success message
            payslip.save()

            # message_box = QtWidgets.QMessageBox()
            # message_box.setText(f"{search_month} Payslip for <b>{selected_employee_No}</b> Generated Successfully")
            # message_box.exec_()
            logging.info(f"{search_month} Payslip generated for {selected_employee_No} ")
        else:
            message_box = QtWidgets.QMessageBox()
            message_box.setWindowTitle("Generate Payslip")
            message_box.setText("No Payslip Found for the selected month")
            message_box.exec_()

            cursor.close()
            conn.close()
    except Exception as error:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText(f"An error {error} occurred while generating Payslip")
        msg.setInformativeText(str(error))
        msg.setWindowTitle("Error")
        msg.exec_()
        logging.critical(f"error occured while generating payslip")

# function generates all payslips at once
def generate_All_Payslips():
    selected_month = Search_payslip_txt.text()

    try:
        if selected_month == "":
            message_box = QtWidgets.QMessageBox()
            message_box.setWindowTitle("Generate Payslip")
            message_box.setText("Please enter month to generate payslips")
            message_box.exec_()
            Search_payslip_txt.setFocus()
            return

        conn = sqlite3.connect("Payroll.db")
        cursor = conn.cursor()

        # Get a list of all employee numbers
        # cursor.execute("SELECT Employee_No FROM current_employees")

        cursor.execute("SELECT Employee_Number FROM Payroll_Data")

        employee_numbers = [row[0] for row in cursor.fetchall()]

        # Generate payslip for each employee and save to a PDF file
        for employee_no in employee_numbers:
            
            # Check if employee has data for the selected month
            # cursor.execute("SELECT * FROM Payroll_Data WHERE Employee_Number = ? AND Payroll_ID = ?", (employee_no, selected_month))

            cursor.execute("SELECT Employee_Number, Payroll_ID FROM Payroll_Data WHERE Employee_Number = ? AND Payroll_ID = ?", (employee_no, selected_month))

            data = cursor.fetchone()
            if not data:
                message_box = QtWidgets.QMessageBox()
                message_box.setWindowTitle("Generate Payslip")
                message_box.setText(f"Employee Number {employee_no} is not in the {selected_month} Payroll Skipping...")
                message_box.exec_()
                continue

            # Set the employee number in the combo box
            Employee_payslip_No_combo.setCurrentText(employee_no)

            # Generate payslip for the selected month
            generate_Payslip()

        cursor.close()
        conn.close()

        message_box = QtWidgets.QMessageBox()
        message_box.setWindowTitle("Generate Payslip")
        message_box.setText(f"Employee Number {employee_no} is not in the {selected_month} Payroll Skipping...")
        message_box.exec_()

    except Exception as error:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText(f"An error {error} occurred while generating Payslips")
        msg.setInformativeText(str(error))
        msg.setWindowTitle("Error")
        msg.exec_()
        logging.critical(f"error occured while generating payslips")


# function to load created designation into the combobox designation
def load_designation_data():
    try:
        conn = sqlite3.connect('Payroll.db')
        cursor = conn.cursor()

        # Retrieve data from the table
        cursor.execute("SELECT Designation FROM Employee_Designation")
        records = cursor.fetchall()  

        # Add the data to the combobox
        for item in records:
            Employee_Designation_Combo.addItem(item[0])

        conn.commit()
        cursor.close()
        conn.close()
    except Exception as error:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText(f"An error {error} occurred while laoding designations")
        msg.setInformativeText(str(error))
        msg.setWindowTitle("Error")
        msg.exec_()
        logging.critical(f"error occured while loading designations")

# Delete function to delete selected items in a QTree   
def delete_Selected_Info():
    global current_table

    try:
        selected_items = tree.selectedItems()

        selected_item = selected_items[0]
        selected_record = selected_item.text(0)
        
        # # Prompt the user for confirmation using a QMessageBox
        reply = QtWidgets.QMessageBox.question(tree, "Confirm Deletion",  
                                            f"Are You Sure You Want To Delete Record {selected_record}?" + "From The System?", 
                                            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)

        if reply == QtWidgets.QMessageBox.Yes:
            # Connect to the database
            conn = sqlite3.connect('Payroll.db')
            cursor = conn.cursor()
            
            # checks which table has been selected for deletion
            if current_table == "Employee_Designation":
                cursor.execute(f"DELETE FROM {current_table} WHERE record_ID = ?", (selected_record,))
            elif current_table == "Payroll_Data":
                cursor.execute(f"DELETE FROM {current_table} WHERE row_ID = ?", (selected_record,))
            elif current_table == "current_employees":
                cursor.execute(f"DELETE FROM {current_table} WHERE employee_NO = ?", (selected_record,))
            
            # Save changes to the database
            conn.commit()
            cursor.close()
            conn.close()

            # Remove the selected item from the QTreeWidget
            parent = selected_item.parent()
            if parent is None:
                tree.takeTopLevelItem(tree.indexOfTopLevelItem(selected_item))
            else:
                parent.takeChild(parent.indexOfChild(selected_item))

            # clears selection after deleting
            tree.clearSelection()

            Employee_No_Combo.clear()

            Employee_payslip_No_combo.clear()

            load_Employee_Number()

            Employee_Designation_Combo.clear()

            load_designation_data()
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Information)
            msg.setText(f"record has been deleted")
            msg.setWindowTitle("Delete")
            msg.exec_()
            logging.critical(f"record {selected_record} was deleted.")

        else:
            tree.clearSelection()
    except Exception as error:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText(f"An error coccured while deleting record")
        msg.setInformativeText(str(error))
        msg.setWindowTitle("Error")
        msg.exec_()
        logging.info(f"error occured while deleting record.")


# Fnction disables and enables the delete button depending on selected record
def on_selection_changed():
    # Get the currently selected items from the QTreeWidget
    selected_items = tree.selectedItems()

    # Enable the delete button if at least one item is selected, otherwise disable it
    Delete_btn.setEnabled(len(selected_items) > 0)

# Connect the itemSelectionChanged signal of the QTreeWidget to the on_selection_changed function
tree.itemSelectionChanged.connect(on_selection_changed)


# function converts a text to upper case
def convert_to_upper(text):
    Name_txt.setText(text.upper())
    Search_payslip_txt.setText(text.upper())
    

Search_payslip_txt.textChanged.connect(convert_to_upper) #converts text to upper case

# def change_row_color():
#     pass

# Function to serch data on the display        
def search():
    search_keyword = Search_box_txt.text().lower()
    if search_keyword:
        for i in range(tree.topLevelItemCount()):
            item = tree.topLevelItem(i)
            for j in range(item.columnCount()):
                if search_keyword in item.text(j).lower():
                    item.setHidden(False)
                    break
            else:
                item.setHidden(True)
    else:
        for i in range(tree.topLevelItemCount()):
            item = tree.topLevelItem(i)
            item.setHidden(False)

# variable to store the current table where data data has been loaded from into the tree
current_table = None

# loads data to the tree depending the table selected or called
def load_data(table_name, search_keyword=None):
    global current_table # to make the variable accessible and not local

    tree.setColumnCount(0)
    tree.clear() # clears the tree
    
    tree.setHeaderHidden(False)

    # Connect to the database
    conn = sqlite3.connect('Payroll.db')
    cursor = conn.cursor()

    # Retrieve column headings from the table
    cursor.execute(f"PRAGMA table_info({table_name})")
    headings = [heading[1] for heading in cursor.fetchall()]

    # Set the header labels for the QTreeWidget
    tree.setHeaderLabels(headings)

    # Retrieve data from the table
    if search_keyword is None:

        cursor.execute(f"SELECT * FROM {table_name}")
    else:
        cursor.execute(f"SELECT * FROM {table_name} WHERE name LIKE ?", ('%'+search_keyword+'%',))
    rows = cursor.fetchall()

    # Add data to the QTreeWidget
    for i, row in enumerate(rows):
        item = QtWidgets.QTreeWidgetItem(tree)
        for j in range(len(headings)):
            item.setText(j, str(row[j]))
            if i % 2 == 0:
                item.setBackground(j, QtGui.QBrush(QtGui.QColor(QtCore.Qt.lightGray)))
            else:        
                item.setBackground(j, QtGui.QBrush(QtGui.QColor(QtCore.Qt.cyan)))

    # Close the cursor and connection
    cursor.close()
    conn.close()

    current_table = table_name

# fucntion to laod current_employees
def load_current_employees():
    load_data("current_employees")

# loads payroll data to the tree
def load_payroll_data():
    try:
       load_data("Payroll_Data")
    except Exception as error:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText(f"An error {error} occurred while loading payroll data.")
        msg.setInformativeText(str(error))
        msg.setWindowTitle("Error")
        msg.exec_()
        logging.critical(f"error occured while retrieving payroll data")

# loads designation data on the screen
def view_designations_data():
    load_data("Employee_Designation")

# function to create designation
def create_designation():
    try:
        if Designation_txt.text() == "":
            msgbox = QtWidgets.QMessageBox()
            msgbox.setWindowTitle("Error!")
            msgbox.setText("Designation Cannot be empty")
            msgbox.exec_()
            return

        conn = sqlite3.connect('Payroll.db')
        cursor = conn.cursor()

         # Create the new table if it doesn't exist
        cursor.execute('''CREATE TABLE IF NOT EXISTS Employee_Designation
                        (record_ID INTEGER PRIMARY KEY AUTOINCREMENT, 
                        Time_stamp TIMESTAMP, 
                        Designation Text 
                        )''')
    
        
        # Check if the designation already exists in the table
        cursor.execute("SELECT COUNT(*) FROM Employee_Designation WHERE UPPER(Designation)=UPPER(?)", (Designation_txt.text(),))
        result = cursor.fetchone()[0]
        if result > 0:
            # Designation already exists, so don't insert it again
            message_box = QtWidgets.QMessageBox()
            message_box.setText(f"The Designation <b>{Designation_txt.text()}</b> Already Exists in the System")
            message_box.exec_()
            
            logging.error(f"User tried to add an existing designation: {Designation_txt.text()}")
        else:
            # Insert data into the table
            cursor.execute("INSERT INTO Employee_Designation (Time_stamp, Designation) VALUES (?, ?)", (timestamp, Designation_txt.text().upper()))


            message_box = QtWidgets.QMessageBox()
            message_box.setText(f"The Designation <b>{Designation_txt.text()}</b> has been added successfully")
            message_box.exec_()

            logging.info(f"User added a new Designation Successfully: {Designation_txt.text()}")

        # Save changes to the database
        conn.commit()
        cursor.close()
        conn.close()

        Designation_txt.clear() # to clear tetxtbox

        Employee_Designation_Combo.clear() # clears the data in the designation combobox

        load_designation_data() # refreshes data in the designation combobox

        view_designations_data()
    except Exception as error:
        # Handle the exception by displaying an error message
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText(f"An error {error} occurred while creating designation")
        msg.setInformativeText(str(error))
        msg.setWindowTitle("Error")
        msg.exec_()
        logging.critical(f"error occured while creating designation")

# This function creates a payroll based on specified month
def Create_Payroll():
    try:
        payroll_period = Payroll_id_txt.text()
        selected_employee_No = Employee_No_Combo.currentText()

        conn = sqlite3.connect("Payroll.db")
        cursor = conn.cursor()

        cursor.execute('''CREATE TABLE IF NOT EXISTS Payroll_Data (
                                                                row_ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                                Time_Stamp TIMESTAMP,
                                                                Payroll_ID,
                                                                Employee_Number,
                                                                ID INTEGER, 
                                                                Name, 
                                                                Designation, 
                                                                Account_Number, 
                                                                KRA_PIN, 
                                                                Salary REAL, 
                                                                Commuter_Allowance INTEGER, 
                                                                House_Allowance REAL, 
                                                                Gross_Pay REAL,
                                                                NSSF REAL,
                                                                Taxable_Amount,
                                                                PAYE,
                                                                NHIF,
                                                                Net_Pay,
                                                                Employer_Contribution
                                                                )''')
        
        # checks if employee Number has been selected by user in the combobox
        if selected_employee_No == "":
            message_box = QtWidgets.QMessageBox()
            message_box.setWindowTitle("Select Employee Number!")
            message_box.setText("Please select an Employee Number first from the dropdown menu.")
            message_box.exec_()
            return
        
        # Check for duplicates
        cursor.execute("SELECT COUNT(*) FROM Payroll_Data WHERE Payroll_ID = ? AND Employee_Number = ?", 
                        (Payroll_id_txt.text(), Employee_No_Combo.currentText()))
        result = cursor.fetchone()
        if result[0] > 0:
            # Record already exists, display error message
            message_box = QtWidgets.QMessageBox()
            message_box.setWindowTitle("Cannot Add User!")
            message_box.setText("Employee number <b>{}</b>.".format(selected_employee_No) + " is Already in the <b>{}</b>.".format(payroll_period)  + " PAYROLL.")
            message_box.exec_()
            logging.info(f"User tried to add employee {selected_employee_No} already in the {payroll_period} payroll ")
        else:
            for line_edit in [Basic_salary_txt, Commuter_allowance_txt, House_allowance_txt ,NSSF_Employee_txt, PAYE_txt, NHIF_txt, Employer_contribution_txt]:
                if line_edit.text() == "":
                    # Set the background color of the empty field to red
                    update_line_edit_color(line_edit)
                    line_edit.setFocus()  # Focus on the empty field
                    return
                else:
                    # Set the background color of the non-empty field to default
                    update_line_edit_color(line_edit)

            # Insert data into the table
            cursor.execute("INSERT INTO Payroll_Data (Time_Stamp,Payroll_ID,Employee_Number, ID, Name, Designation, Account_Number,KRA_PIN, Salary,\
                        Commuter_Allowance, House_Allowance, Gross_Pay,NSSF,Taxable_Amount,\
                        PAYE,NHIF,Net_Pay,Employer_Contribution) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (
                                                                                        timestamp,
                                                                                        Payroll_id_txt.text(),
                                                                                        Employee_No_Combo.currentText(),
                                                                                        name_lbl.text(), 
                                                                                        ID_No_lbl.text(),
                                                                                        Designation_lbl.text(),
                                                                                        Account_Number_lbl.text(),
                                                                                        KRA_PIN_lbl.text(),
                                                                                        Basic_salary_txt.text(),
                                                                                        Commuter_allowance_txt.text(),
                                                                                        House_allowance_txt.text(),
                                                                                        Gross_pay_lbl.text(),
                                                                                        NSSF_Employee_txt.text(),
                                                                                        Taxable_amount_lbl.text(),
                                                                                        PAYE_txt.text(),
                                                                                        NHIF_txt.text(),
                                                                                        Net_pay_lbl.text(),
                                                                                        Employer_contribution_txt.text(),                                                                        
                                                                                                              
                                                                                    ))
            
            message_box = QtWidgets.QMessageBox()  
            message_box.setText("Employee number <b>{}</b>.".format(selected_employee_No) + " succesffuly added to the <b>{}</b>.".format(payroll_period)  + " PAYROLL.")
            message_box.exec_()
            logging.info(f"Employee number {selected_employee_No} successfuly added to {payroll_period} payroll.")
        
        conn.commit()
        cursor.close()
        conn.close()

        # # Clearing the textboxes after submission
        textboxes =  [Basic_salary_txt, Commuter_allowance_txt, House_allowance_txt, NSSF_Employee_txt, PAYE_txt, NHIF_txt, Employer_contribution_txt]
        for tb in textboxes:
            tb.clear()

            palette.setColor(QPalette.Base, default_color) # sets the line edits back to default color afer saving
            tb.setPalette(palette)

        Employee_No_Combo.setFocus()
    except Exception as error:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText(f"An error {error} occurred while creating Payroll for {payroll_period}")
        msg.setInformativeText(str(error))
        msg.setWindowTitle("Error")
        msg.exec_()
        logging.critical(f"error occured while retrieving {payroll_period} payroll")

# creating a palette for line edits background color
palette = QPalette()
empty_color = QColor(255, 150, 150)  # Red color for empty fields
default_color = QColor(255, 255, 255)  # Default color for non-empty fields

# function updates the background color of each textbox/Qlinedit depending on the data entered
def update_line_edit_color(line_edit):
    if line_edit.text() == "":
        # Set the background color of the empty field to red
        palette.setColor(QPalette.Base, empty_color)
        line_edit.setPalette(palette)
        return
    else:
        # Set the background color of the non-empty field to default
        palette.setColor(QPalette.Base, default_color)
        line_edit.setPalette(palette)


# Connects the textChanged signal of each QLineEdit field to the update_line_edit_color function
for txt in [ID_No_txt, Name_txt, Account_Number_txt, Designation_txt, Contact_txt, KRA_PIN_txt, Search_payslip_txt, Basic_salary_txt, \
            Commuter_allowance_txt, House_allowance_txt, NSSF_Employee_txt, PAYE_txt, NHIF_txt, Employer_contribution_txt, Payroll_id_txt ]:
    txt.textChanged.connect(lambda text, le=txt: update_line_edit_color(le))

#  Function computes the grosspay received from allowances
def update_gross_pay():
    basic_salary = Basic_salary_txt.text()
    commuter_allowance = Commuter_allowance_txt.text()
    house_allowance = House_allowance_txt.text()
    if basic_salary and commuter_allowance and house_allowance :
        gross_pay = float(basic_salary) + float(commuter_allowance) + float(house_allowance)

        # rounding the grosspay to two decimal places
        gross_pay_rounded = round(gross_pay, 2)
        Gross_pay_lbl.setText("{:.2f}".format(gross_pay_rounded))
        update_taxable_amount()
        update_net_pay()
    else:
        Gross_pay_lbl.setText("")

# function to compute texable amount
def update_taxable_amount():
    gross_pay = Gross_pay_lbl.text()
    NSSF = NSSF_Employee_txt.text()

    if gross_pay and NSSF:
        taxable_amount = float(gross_pay) - float(NSSF)
        taxable_amount_rounded = round(taxable_amount, 2)
        Taxable_amount_lbl.setText("{:.2f}".format(taxable_amount_rounded))
        update_net_pay()
    else:
        Taxable_amount_lbl.setText("")

def update_net_pay():
    taxable_amount = Taxable_amount_lbl.text()
    PAYE = PAYE_txt.text()
    NHIF = NHIF_txt.text()

    if NHIF and PAYE:
        PAYE = float(PAYE_txt.text())
        NHIF = float(NHIF_txt.text())    
    else:
        NHIF = 0.0
        PAYE = 0.0

    if taxable_amount:
        net_pay = float(taxable_amount) - (PAYE + NHIF)

        net_pay_rounded = round(net_pay, 2)
        Net_pay_lbl.setText("{:.2f}".format(net_pay_rounded))
    else:
        Net_pay_lbl.setText("")


# Connect the textChanged signals of the two widgets to the update_gross_pay function
Basic_salary_txt.textChanged.connect(update_gross_pay)
Commuter_allowance_txt.textChanged.connect(update_gross_pay)
House_allowance_txt.textChanged.connect(update_gross_pay)
NSSF_Employee_txt.textChanged.connect(update_taxable_amount)
PAYE_txt.textChanged.connect(update_net_pay)
NHIF_txt.textChanged.connect(update_net_pay)

Name_txt.textChanged.connect(convert_to_upper) # converts text entered in the name field to upper case

# Function to populate salary info into respective input fields
def populate_salary_info(): 
    selected_employee_No = Employee_No_Combo.currentText()

    if selected_employee_No == "":
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText(f"Please Select an employee Number to load salary information")
        msg.setWindowTitle("Error")
        msg.exec_()
        return

    conn = sqlite3.connect("Payroll.db")
    cursor = conn.cursor()

    try:
        cursor.execute("SELECT MAX(Time_Stamp) FROM Payroll_Data WHERE Employee_Number = ?", (selected_employee_No,)) 
        
        most_recent_payroll_period = cursor.fetchone()[0]
       
        cursor.execute("SELECT Salary, Commuter_Allowance, House_Allowance, NSSF, PAYE, NHIF, Employer_Contribution \
                          FROM Payroll_Data WHERE Employee_Number = ? \
                         AND Time_Stamp = ? ORDER BY Time_Stamp DESC LIMIT 1", (selected_employee_No, most_recent_payroll_period,))
        
        selected_info = cursor.fetchone()
         # Creating a list of tuples containing the line edit objects and their corresponding column indices
        txt_list = [(Basic_salary_txt, 0), (Commuter_allowance_txt, 1), (House_allowance_txt, 2), (NSSF_Employee_txt, 3), (PAYE_txt, 4), (NHIF_txt, 5), (Employer_contribution_txt, 6)]
        if selected_info is not None:
            # Loop through the txt_list and populate each line edit with the corresponding value from selected_info
            for txt, column_index in txt_list:
                value = selected_info[column_index]
                txt.setText(str(value))
        else:
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Information)
            msg.setWindowTitle("Niaje, I Cannot Load Salary Information")
            msg.setText("Employee number <b>{}</b>.".format(selected_employee_No) + " was not in the Previous PAYROLL.")
            msg.exec_()

            # Clear the text of the line edits if no data was retrieved from the database
            for txt, column_index in txt_list:
                txt.setText("")
    except Exception as e:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText("An error occurred while retrieving salary information for employee <b>{}</b>.".format(selected_employee_No))
        msg.setInformativeText(str(e))
        msg.setWindowTitle("Error")
        msg.exec_()

export_button.setEnabled(False)  

# function checks whether the tree has data loaded into it if so the export button is enabled
def has_data_loaded():
   if tree.topLevelItemCount() > 0:
       export_button.setEnabled(True)

# Connecting the itemChanged signal of the QTreeWidget to the has_loaded_fucntion function
tree.itemChanged.connect(has_data_loaded)

# Function to export in EXCEL format
def export_data():
    try:
        if tree.topLevelItemCount() == 0:
            msg = QtWidgets.QMessageBox()
            msg.setWindowTitle("ERROR!")
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            msg.setText("Niaje, Cannot Export Empty Data")
            msg.exec_()
            logging.error("User failed to load data to Export")
            return
        
        # Get the currrent date and time
        now = datetime.datetime.now()

        # Format the date and time for use in the file name
        timestamp = now.strftime("%d-%m-%Y_%H-%M-%S")

        # Create a new workbook
        wb = Workbook()

        # Create the first worksheet
        ws = wb.active
    
        bold_font = styles.Font(bold=True)

        # Write the header row with column names from the QTreeWidget
        header = []    
        for i in range(tree.columnCount()):
            header.append(tree.headerItem().text(i))
        ws.append(header)

        # making the column headers bold
        for cell in ws["1:1"]:
            cell.font = bold_font

        # Write the data from the QTreeWidget
        for i in range(tree.topLevelItemCount()):
            item = tree.topLevelItem(i)
            if item.isHidden() == False:  # check if item is filtered out
                data = []
                for j in range(tree.columnCount()):
                    data.append(item.text(j))
                ws.append(data)

        # Save the workbook with the timestamped filename
        wb.save(f"Excel_Export_{timestamp}.xlsx")
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle("Export")
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText(f"Niaje, <b>Excel_Export_{timestamp}.xlsx</b> has been exported successfully.")
        msg.exec_()
        logging.info(f"Excel_Export_{timestamp}.xlsx was exported.")
    except Exception as error:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setText("An error occurred while exporting Data")
        msg.setInformativeText(str(error))
        msg.setWindowTitle("Error")
        msg.exec_()

def backup():
    try:
        # Replace with your own email credentials and recipient email
        sender_email = "payroll20231@outlook.com"
        sender_password = "lambistic1"
        recipient_email = "antolando231@gmail.com"

        # Create a MIMEMultipart object to construct the email
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = recipient_email
        message['Subject'] = 'Payroll Backup'

        # Get the path of the directory where the app is installed
        app_dir = os.path.dirname(sys.argv[0])

        # Set the payroll file path to the Payroll.db file in the app directory
        payroll_file_path = os.path.join(app_dir, 'Payroll.db')

        # Set the tax file path to the Taxes.xlsx file in the app directory
        payroll_log_file_path = os.path.join(app_dir, 'payroll_logs.log')

        # Check if the files exist
        if not os.path.exists(payroll_file_path) or not os.path.exists(payroll_log_file_path):
            msg = QtWidgets.QMessageBox()
            msg.setWindowTitle("Error")
            msg.setText("One or more files cannot be found.")
            msg.exec_()
            logging.critical("One or more files not found in PATH to be backed up")
            return
            
        # Open the payroll database file and read the contents
        with open(payroll_file_path, 'rb') as f:
            payroll_db = f.read()

        # Open the taxes file and read the contents
        with open(payroll_log_file_path, 'rb') as f:
            payroll_log = f.read()

        # Create MIME application objects to represent the files
        payroll_db_attachment = MIMEApplication(payroll_db, Name=os.path.basename(payroll_file_path))
        payroll_db_attachment['Content-Disposition'] = f'attachment; filename="{os.path.basename(payroll_file_path)}"'

        payroll_log_attachment = MIMEApplication(payroll_log, Name=os.path.basename(payroll_log_file_path))
        payroll_log_attachment['Content-Disposition'] = f'attachment; filename="{os.path.basename(payroll_log_file_path)}"'

        # Attach the files to the email message
        message.attach(payroll_db_attachment)
        message.attach(payroll_log_attachment)

        # Connect to the SMTP server and send the email
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, message.as_string())

        # Display a message box to indicate that the email has been sent
        message_box = QtWidgets.QMessageBox()
        message_box.setText(f"Backup Sent Successfully")
        message_box.exec_()
        logging.info("Backup was sent successfully")
    except Exception as Error:
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle("Internet Connection")
        msg.setText("Error Occurred Sending Backup Might Check Internet Connection")
        msg.exec_()
        logging.critical("Error Occured while Sending Backup")

# function greets user when the app launches
def greet():
    current_time = QTime.currentTime()

    # Set the greeting based on the current time
    if current_time.hour() < 12:
        greeting = "Morning"
    elif current_time.hour() < 17:
        greeting = "Afternoon"
    else:
        greeting = "Evening"
        
    return  f" <b>Good {greeting}  Admin Hope You Doing Great</b> "

def hide_greet(): # will hide the greet label
    greet_label.hide()
    
greet_label.setText(greet())


# Create a timer that will call the hide_label function after 5 seconds
timer = QTimer()
timer.setSingleShot(True)
timer.timeout.connect(hide_greet)
timer.start(5000)

# this checks if the database and tables created exists immediately the app is launched
if not os.path.exists('Payroll.db'):
    logging.info("DB Created") 
    conn = sqlite3.connect('Payroll.db')
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS Registration (employee_No, 
                                                               Name text, 
                                                               ID Number text, 
                                                               Designation, 
                                                               Account_Number text, 
                                                               Date_of_Joining text,
                                                               KRA PIN text,
                                                               COntact text)''')

    logging.info("Registration table created") 
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS Employee_Designation
                        (record_ID INTEGER PRIMARY KEY AUTOINCREMENT, 
                        Time_stamp TIMESTAMP, 
                        Designation Text 
                        )''')
    
    logging.info("Employee_Designation table created") 

    cursor.execute('''CREATE TABLE IF NOT EXISTS Payroll_Data ( row_ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                                Time_Stamp TIMESTAMP,
                                                                Payroll_ID,
                                                                Employee_Number,
                                                                ID INTEGER, 
                                                                Name, 
                                                                Designation, 
                                                                Account_Number, 
                                                                KRA_PIN, 
                                                                Salary REAL, 
                                                                Commuter_Allowance INTEGER, 
                                                                House_Allowance REAL, 
                                                                Gross_Pay REAL,
                                                                NSSF REAL,
                                                                Taxable_Amount,
                                                                PAYE,
                                                                NHIF,
                                                                Net_Pay,
                                                                Employer_Contribution
                                                                )''')
    
    logging.info("Payroll_Data Table Created")

    cursor.execute('''CREATE TABLE IF NOT EXISTS current_employees (Employee_NO, 
                                                                      Name text, 
                                                                      ID_Number text,
                                                                      Designation text, 
                                                                      Account_Number text, 
                                                                      Date_of_Joining text,
                                                                      KRA_PIN text,
                                                                      Contact text
                                                                      )''')
    
    logging.info("current_employees table created")

    conn.commit()
    conn.close()
else:
    conn = sqlite3.connect('Payroll.db')

load_designation_data() #loads designation data into designation combobox

load_Employee_Number() #loads data into the combobox

setCompleterForColumn(Search_payslip_txt, "Payroll_ID") # crreates autoccompletion for a linedit

# splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
splitter = QtWidgets.QSplitter()

# Registration widget to hold the registration form layout
Registration_widget = QtWidgets.QWidget()
Registration_widget.setFixedSize(500, 500)
Registration_widget.setWindowTitle("REGISTRATION FORM")
Registration_layout = QtWidgets.QFormLayout(Registration_widget)
Registration_layout.addWidget(ID_No_txt)
Registration_layout.addWidget(Name_txt)
Registration_layout.addRow("Designation:", Employee_Designation_Combo)
Registration_layout.addWidget(Account_Number_txt)
Registration_layout.addWidget(date_chooser)
Registration_layout.addRow("Date of Joining:", date_chooser)
Registration_layout.addWidget(KRA_PIN_txt)
Registration_layout.addWidget(Contact_txt)
Registration_layout.addWidget(submit_button)
Registration_layout.addWidget(Update_btn)

## HOME widget to create the HOME menu to hold the register and create payroll button 
Home_widget = QtWidgets.QWidget()
Home_layout = QtWidgets.QFormLayout(Home_widget)
Home_layout.addWidget(greet_label)
Home_layout.addWidget(load_Current_Employees_btn)
Home_layout.addWidget(Regsiter_btn)
Home_layout.addWidget(Designation_btn)
Home_layout.addWidget(Create_Payroll_btn)
Home_layout.addWidget(load_Payroll_Data_btn)
Home_layout.addWidget(Payslip_btn)
Home_layout.addWidget(Delete_btn)
Home_layout.addWidget(export_button)
Home_layout.addWidget(Search_payslip_txt)
Home_layout.addWidget(change_theme_btn)
Home_layout.addWidget(backup_btn)

# Payslip widget
payslip_widget = QtWidgets.QWidget()
payslip_widget.setWindowTitle("PAYSLIPS")
payslip_widget.setFixedSize(400, 200)
payslip_layout = QtWidgets.QFormLayout(payslip_widget)
payslip_layout.addWidget(Employee_payslip_No_combo)
payslip_layout.addWidget(Search_payslip_txt)
payslip_layout.addWidget(Generate_payslip_btn)
payslip_layout.addWidget(all_payslips_btn)

## PAYROLL WIDGET to hold the widget layout
Payroll_widget = QtWidgets.QWidget()
Payroll_widget.setWindowTitle("PAYROLL FORM")
Payroll_widget.setFixedSize(600, 700)
payroll_layout = QtWidgets.QFormLayout(Payroll_widget)
payroll_layout.addRow("Select Employee Number:", Employee_No_Combo)
payroll_layout.addWidget(load_salary_info_btn)
payroll_layout.addRow("Payroll ID", Payroll_id_txt)
payroll_layout.addRow("ID:", name_lbl)
payroll_layout.addRow("NAME:", ID_No_lbl)
payroll_layout.addRow("DESIGNATION:",Designation_lbl)
payroll_layout.addRow("BANK ACCOUNT:", Account_Number_lbl)
payroll_layout.addRow("KRA PIN:", KRA_PIN_lbl)
payroll_layout.addRow("BASIC SALARY:", Basic_salary_txt)
payroll_layout.addRow("COMMUTER ALLOWANCE:", Commuter_allowance_txt)
payroll_layout.addRow("HOUSE ALLOWANCE:", House_allowance_txt)
payroll_layout.addRow("GROSS PAY:", Gross_pay_lbl)
payroll_layout.addRow("NSSF:", NSSF_Employee_txt)
payroll_layout.addRow("TAXABLE AMOUNT:", Taxable_amount_lbl)
payroll_layout.addRow("PAYE:", PAYE_txt)
payroll_layout.addRow("NHIF:", NHIF_txt)
payroll_layout.addRow("NET PAY:", Net_pay_lbl)
payroll_layout.addRow("EMPLOYER NSSF:", Employer_contribution_txt)
payroll_layout.addWidget(Save_payroll_btn)

# Employee designation widget
designation_widget = QtWidgets.QWidget()
designation_widget.setWindowTitle("DESIGNATION FORM")
designation_widget.setFixedSize(500, 150)
designation_layout = QtWidgets.QFormLayout(designation_widget)
designation_layout.addWidget(load_Designations_btn)

# contains the designation widgets
designation_layout.addWidget(Designation_txt)
designation_layout.addWidget(Add_Designation_btn)

button_layout = QtWidgets.QHBoxLayout()
button_layout.addWidget(Delete_btn)
button_layout.addWidget(export_button)

search_layout = QtWidgets.QHBoxLayout()
search_layout.addWidget(Search_box_txt)
search_layout.addWidget(search_btn)

# display layout to hold the tree
display_layout = QtWidgets.QVBoxLayout()
display_layout.addLayout(search_layout)
display_layout.addWidget(tree)
display_layout.addLayout(button_layout)

# # groupbox to hold the widgets
group_box = QtWidgets.QGroupBox("Registration")

# group_box.setLayout(Registration_layout)
group_box.setLayout(Home_layout)
# group_box.setStyleSheet("background-color: magenta;")

# frame to hold the display layout
frame = QtWidgets.QFrame()
frame.setLayout(display_layout)
frame.setFrameStyle(QtWidgets.QFrame.Panel | QtWidgets.QFrame.Raised)
frame.setStyleSheet("QFrame {border: 2px solid blue; border-radius: 10px;}")

# Add the QTreeWidget to the right side of the splitter
splitter.addWidget(group_box)

splitter.addWidget(frame)

## GIVING BUTTONS FUNCTIONALITIES
submit_button.clicked.connect(submit_data)
Regsiter_btn.clicked.connect(lambda: Registration_widget.setVisible(True))
Create_Payroll_btn.clicked.connect(lambda: Payroll_widget.setVisible(True))
Designation_btn.clicked.connect(lambda: designation_widget.setVisible(True))
Payslip_btn.clicked.connect(lambda: payslip_widget.setVisible(True))
Add_Designation_btn.clicked.connect(create_designation)
Save_payroll_btn.clicked.connect(Create_Payroll)
load_salary_info_btn.clicked.connect(populate_salary_info)
change_theme_btn.clicked.connect(change_Theme)
backup_btn.clicked.connect(backup)
all_payslips_btn.clicked.connect(generate_All_Payslips)

Delete_btn.clicked.connect(delete_Selected_Info)
Delete_btn.setEnabled(False)
load_Current_Employees_btn.clicked.connect(load_current_employees)
load_Payroll_Data_btn.clicked.connect(load_payroll_data)
load_Designations_btn.clicked.connect(view_designations_data)
Generate_payslip_btn.clicked.connect(generate_Payslip)
export_button.setToolTip("export")
export_button.clicked.connect(export_data)
search_btn.clicked.connect(search)
Update_btn.clicked.connect(update_data)

layout.addWidget(splitter)

window.show()

app.exec_()
logging.info("The system shutdown")