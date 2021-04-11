from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox
import sys
import os
import requests
from wordpress_xmlrpc import Client
from wordpress_xmlrpc.methods.users import GetUsers
import csv


#Global variable for the SKU increment
count = 1
#Global for the Autocomplete
count_placeholder_needed = 0

class Ui_MainWindow(object):
    
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(536, 439)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.TitleInput = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.TitleInput.setGeometry(QtCore.QRect(20, 30, 281, 31))
        self.TitleInput.setObjectName("TitleInput")
        self.VendorInput = QtWidgets.QComboBox(self.centralwidget)
        self.VendorInput.setGeometry(QtCore.QRect(20, 105, 281, 31))
        self.VendorInput.setObjectName("VendorInput")

        for row in open('number_id.csv', 'r'):
            self.VendorInput.addItem("")

        self.PriceInput = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.PriceInput.setGeometry(QtCore.QRect(20, 180, 125, 31))
        self.PriceInput.setObjectName("PriceInput")

        self.DescriptionInput = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.DescriptionInput.setGeometry(QtCore.QRect(20, 250, 481, 81))
        self.DescriptionInput.setObjectName("DescriptionInput")

        self.PriceLabel = QtWidgets.QLabel(self.centralwidget)
        self.PriceLabel.setGeometry(QtCore.QRect(20, 160, 61, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.PriceLabel.setFont(font)
        self.PriceLabel.setObjectName("PriceLabel")

        self.TitleLabel = QtWidgets.QLabel(self.centralwidget)
        self.TitleLabel.setGeometry(QtCore.QRect(20, 10, 101, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.TitleLabel.setFont(font)
        self.TitleLabel.setObjectName("TitleLabel")
        self.VendorLabel = QtWidgets.QLabel(self.centralwidget)
        self.VendorLabel.setGeometry(QtCore.QRect(20, 80, 151, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.VendorLabel.setFont(font)
        self.VendorLabel.setObjectName("VendorLabel")
        self.QuantityLabel = QtWidgets.QLabel(self.centralwidget)
        self.QuantityLabel.setGeometry(QtCore.QRect(150, 160, 61, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.QuantityLabel.setFont(font)
        self.QuantityLabel.setObjectName("QuantityLabel")
        self.DescriptionLabel = QtWidgets.QLabel(self.centralwidget)
        self.DescriptionLabel.setGeometry(QtCore.QRect(20, 230, 141, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.DescriptionLabel.setFont(font)
        self.DescriptionLabel.setObjectName("DescriptionLabel")
        self.VendorIDLabel = QtWidgets.QLabel(self.centralwidget)
        self.VendorIDLabel.setGeometry(QtCore.QRect(90, 160, 61, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.VendorIDLabel.setFont(font)
        self.VendorIDLabel.setObjectName("VendorIDLabel")
        self.ProductTypeDropDown = QtWidgets.QComboBox(self.centralwidget)
        self.ProductTypeDropDown.setGeometry(QtCore.QRect(320, 30, 191, 31))
        self.ProductTypeDropDown.setObjectName("ProductTypeDropDown")
        self.ProductTypeDropDown.addItem("")
        self.ProductTypeDropDown.addItem("")
        self.ProductTypeDropDown.addItem("")
        self.ProductTypeDropDown.addItem("")
        self.ProductTypeDropDown.addItem("")
        self.ProductTypeDropDown.addItem("")
        self.ProductTypeDropDown.addItem("")
        self.ProductTypeLabel = QtWidgets.QLabel(self.centralwidget)
        self.ProductTypeLabel.setGeometry(QtCore.QRect(320, 10, 101, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.ProductTypeLabel.setFont(font)
        self.ProductTypeLabel.setObjectName("ProductTypeLabel")
        self.SaveButton = QtWidgets.QPushButton(self.centralwidget)
        self.SaveButton.setGeometry(QtCore.QRect(20, 350, 481, 41))
        self.SaveButton.setObjectName("SaveButton")

        self.TaxableLabel = QtWidgets.QLabel(self.centralwidget)
        self.TaxableLabel.setGeometry(QtCore.QRect(230, 160, 71, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.TaxableLabel.setFont(font)
        self.TaxableLabel.setObjectName("TaxableLabel")
        self.TaxableInput = QtWidgets.QComboBox(self.centralwidget)
        self.TaxableInput.setGeometry(QtCore.QRect(230, 180, 71, 31))
        self.TaxableInput.setObjectName("TaxableInput")
        self.TaxableInput.addItem("")
        self.TaxableInput.addItem("")
        self.TaxableInput.addItem("")
        self.QuantityInput = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.QuantityInput.setGeometry(QtCore.QRect(150, 180, 75, 31))
        self.QuantityInput.setObjectName("QuantityInput")

        self.VendorEmailInput = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.VendorEmailInput.setGeometry(QtCore.QRect(320, 105, 191, 31))
        self.VendorEmailInput.setObjectName("VendorEmailInput")
        self.VendorEmailLabel = QtWidgets.QLabel(self.centralwidget)
        self.VendorEmailLabel.setGeometry(QtCore.QRect(320, 80, 151, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.VendorEmailLabel.setFont(font)
        self.VendorEmailLabel.setObjectName("VendorEmailLabel")
        self.Logo = QtWidgets.QLabel(self.centralwidget)
        self.Logo.setGeometry(QtCore.QRect(320, 150, 171, 81))
        self.Logo.setText("")
        self.Logo.setPixmap(QtGui.QPixmap("logo.png"))
        self.Logo.setScaledContents(True)
        self.Logo.setObjectName("Logo")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionExit = QtWidgets.QAction(MainWindow)
        self.actionExit.setObjectName("actionExit")
        self.actionSave_Current_Product = QtWidgets.QAction(MainWindow)
        self.actionSave_Current_Product.setObjectName("actionSave_Current_Product")
        self.actionsave = QtWidgets.QAction(MainWindow)
        self.actionsave.setObjectName("actionsave")

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Northern Enviromental Product editor"))
        self.TitleLabel.setText(_translate("MainWindow", "Title"))
        self.VendorLabel.setText(_translate("MainWindow", "Vendor Name"))
        self.QuantityLabel.setText(_translate("MainWindow", "Quantity"))
        self.DescriptionLabel.setText(_translate("MainWindow", "Description"))
        self.VendorIDLabel.setText(_translate("MainWindow", ""))

        self.VendorInput.setCurrentText((_translate("MainWindow", " ")))
        # vendors = []
        # vendor_id = []
        vendors_dict = {}
        temp_count = 0
        global count_placeholder_needed
        with open('number_id.csv', 'r', newline='') as csvfile:
            reader = csv.DictReader(csvfile, delimiter=',')
            header = next(reader)
            for row in reader:
                count_placeholder_needed = count_placeholder_needed + 1
                temp_count = temp_count + 1
                self.VendorInput.setItemText(temp_count, _translate("MainWindow", row['Username']))

        self.ProductTypeDropDown.setCurrentText(_translate("MainWindow", " "))
        self.ProductTypeDropDown.setItemText(0, _translate("MainWindow", "Consignment"))
        self.ProductTypeDropDown.setItemText(1, _translate("MainWindow", "Lending Library"))
        self.ProductTypeDropDown.setItemText(2, _translate("MainWindow", "General Merchandise"))
        self.ProductTypeDropDown.setItemText(3, _translate("MainWindow", "Fresh Produce"))
        self.ProductTypeDropDown.setItemText(4, _translate("MainWindow", "Bulk"))
        self.ProductTypeDropDown.setItemText(5, _translate("MainWindow", "Nourish"))
        self.ProductTypeDropDown.setItemText(6, _translate("MainWindow", "Northern Cohort"))
        self.ProductTypeLabel.setText(_translate("MainWindow", "Product Type"))
        self.PriceLabel.setText(_translate("MainWindow", "Price"))
        self.SaveButton.setText(_translate("MainWindow", "Save"))
        self.TaxableLabel.setText(_translate("MainWindow", "Taxable?"))
        self.TaxableInput.setItemText(0, _translate("mainWindow", " "))
        self.TaxableInput.setItemText(1, _translate("mainWindow", "Yes"))
        self.TaxableInput.setItemText(2, _translate("mainWindow", "No"))
        self.VendorEmailLabel.setText(_translate("MainWindow", "Vendor Email"))
        self.actionExit.setText(_translate("MainWindow", "Exit"))
        self.actionSave_Current_Product.setText(_translate("MainWindow", "Save"))
        self.actionsave.setText(_translate("MainWindow", "save"))
        
        self.SaveButton.clicked.connect(self.product_to_csv)


    def product_to_csv(self):

        #filename = 'UPLOAD.csv'
        #self.init_file(filename)
        #Pull all data from the fields
        #Input feilds
        title = self.TitleInput.toPlainText()
        vendor_email = self.VendorEmailInput.toPlainText()
        quantity = self.QuantityInput.toPlainText()
        description = self.DescriptionInput.toPlainText()
        price = self.PriceInput.toPlainText()
        
        #Combo Boxes
        vendor_name = self.VendorInput.currentText()
        tax = self.TaxableInput.currentText()
        product_type = self.ProductTypeDropDown.currentText()

        self.init_file(str(vendor_name) + '.csv')

        #Grab the Username and ID
        current_id = 0
        input_file = csv.DictReader(open("number_id.csv"))
        for row in input_file:
            if vendor_name == row['Username']:
                current_id = row['ID']

        #Product Type for sku
        type_num = 0
        if product_type == 'Consignment':
            type_num = 7
        if product_type == 'Lending Library':
            type_num = 6
        if product_type == 'General Merchandise':
            type_num = 5 
        if product_type == 'Fresh Produce':
            type_num = 4
        if product_type == 'Bulk':
            type_num = 3
        if product_type == 'Nourish':
            type_num = 2
        if product_type == 'Northern Cohort':
            type_num = 1

        #quantity counter
        global count
        count_str = str(count)
        zero_filled_number = count_str.zfill(3)

        #Supplier code for SKU
        letters = re.findall('[A-Z]', vendor_name)
        supplier_code_full = ''.join(letters)
        supplier_code = supplier_code_full[0:2]

        #Create title and SKU 
        final_product_title = str(vendor_name) + " // " + str(title)
        final_product_sku = str(type_num) + str(supplier_code).upper() + str(current_id) + zero_filled_number

        

        taxable = ''
        tax_rate = ''

        if tax == 'Yes':
            taxable = 'taxable'
            tax_rate = ''
        if tax == 'No':
            taxable = 'none'
            tax_rate = 'zero-rate'
        

        #Define output list
        output_list = ['','simple',final_product_sku,final_product_title,'1','0','visible','',description,'','',taxable,tax_rate,'1',quantity,'','0','0','','','','','1','','',price,'Uncategorized','','','','','','','','','','','','0',current_id,'','','global','global','','','no','','no','','no','no','no','','','','','','','','','','','','','',]

        #Clear all Input and prompt a confirmation of save
        self.clear_labels()
        self.message_box_save()

        file_out = open(str(vendor_name) + '.csv', 'a', newline='')
        writerobj = csv.writer(file_out, delimiter=',')

        writerobj.writerow(output_list)
        #increments count so multiple entries sku qty gets increment
        count = count + 1
        

    def init_file(self, file_name):

        if os.path.exists(file_name):
            pass
        else:
            header = ['ID','Type','SKU','Name','Published',
            'Is featured?','Visibility in catalogue','Short description',
            'Description','Date sale price starts','Date sale price ends',
            'Tax status','Tax class','In stock?','Stock','Low stock amount',
            'Backorders allowed?','Sold individually?','Weight (kg)',
            'Length (cm)','Width (cm)','Height (cm)','Allow customer reviews?','Purchase note','Sale price','Regular price','Categories','Tags','Shipping class','Images','Download limit','Download expiry days','Parent','Grouped products','Upsells','Cross-sells','External URL','Button text','Position','Meta: vendor_id','Meta: _booking_min','Meta: _booking_max','Meta: _number_of_dates','Meta: _booking_duration','Meta: _custom_booking_duration','Meta: _first_available_date','Meta: _bookable','Meta: _wc_memberships_product_viewing_restricted_message','Meta: _wc_memberships_use_custom_product_viewing_restricted_message','Meta: _wc_memberships_product_purchasing_restricted_message','Meta: _wc_memberships_use_custom_product_purchasing_restricted_message','Meta: _wc_memberships_force_public','Meta: _wc_memberships_exclude_discounts','Meta: pv_commission_rate','Meta: woosb_ids','Meta: woosb_disable_auto_price','Meta: woosb_discount','Meta: woosb_discount_amount','Meta: woosb_shipping_fee','Meta: woosb_optional_products','Meta: woosb_manage_stock','Meta: woosb_limit_each_min','Meta: woosb_limit_each_max','Meta: woosb_limit_each_min_default','Meta: woosb_limit_whole_min','Meta: woosb_limit_whole_max','Meta: woosb_custom_price']
            file_out = open(file_name, 'w', newline='')
            writer_out = csv.writer(file_out, delimiter=',')
            writer_out.writerow(header)
            file_out.close()

    def clear_labels(self):

        #Clears all fields with text in it. 
        self.TitleInput.clear()
        self.QuantityInput.clear()
        self.DescriptionInput.clear()
        self.VendorEmailInput.clear()
        self.PriceInput.clear()


    def message_box_save(self):

            product = self.TitleInput.toPlainText()
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText(str(product) + "Product Added.")
            msg.setWindowTitle("Confirm Save")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()


def vendor_autocomplete_startup():

    url = 'https://neat.ca/xmlrpc.php'
    username = ''
    password = ''
    wp = Client(url, username, password)
    users = wp.call(GetUsers())
    subs = []
    for u in users:
        if 'vendor' in u.roles:
            subs.append((u.nickname, u.id))
    
            
    header = ['Username', 'ID']
    fireweed = ['Fireweed', '29']
    with open('number_id.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(header)
        for x in subs:
            writer.writerow(x)
        writer.writerow(fireweed)
    


    


if __name__ == "__main__":

    import sys
    import csv
    import os
    from datetime import datetime
    import calendar
    import re
    from woocommerce import API
    
    vendor_autocomplete_startup()
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
