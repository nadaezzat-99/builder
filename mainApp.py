from PyQt6.QtWidgets import QMainWindow, QApplication, QFileDialog
import sys
from window import Ui_MainWindow
import pandas as pd
from PIL import Image
import os
import img2pdf
import oracledb

class converterWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.show()
        self.transactions = None
        self.required_numbers = []
        self.file_data= None
        self.target_dir = None
        self.browse_button.clicked.connect(self.browse_file)
        self.total_button.clicked.connect(self.get_data)
        self.disable_buttons()
        

               
    def enable_buttons(self):
        self.total_button.setStyleSheet(self.enabled_button_style)
        self.total_button.setEnabled(True)
    
    def disable_buttons(self):
        self.total_button.setStyleSheet(self.disabled_button_style)
        self.total_button.setEnabled(False)


    def save_file(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.FileMode.Directory)
        file_dialog.setOption(QFileDialog.Option.ShowDirsOnly, True)
        
        if file_dialog.exec() == True:
            print(file_dialog.selectedFiles()[0])
            self.target_dir = file_dialog.selectedFiles()[0]
        else:
            return
 
 
    def browse_file(self):
        file_dialog = QFileDialog()
        fname = file_dialog.getOpenFileName(self, "browse file", filter="Excel Files (*.xlsx *.xls)")
        if not fname[0]:
            self.disable_buttons() 
            return
        self.filename.setText(os.path.basename(fname[0]))
        self.enable_buttons()
        self.file_data = pd.read_excel(fname[0])
        self.get_phoneNumbers()
        
      
    def get_phoneNumbers(self):
        for index,row in self.file_data.iterrows():
            self.required_numbers.append({"number":"0"+str(row["Phone Number"]),"identifier":str(row["External Number"])})
            

    def get_data(self):
        try:
            with oracledb.connect(user='ACCEPTANCE', password='acceptance', dsn='10.50.70.5:1521/mwtest') as connection:
                with connection.cursor() as cursor: 
                    for merchant_data in self.required_numbers:
                        query = """SELECT NID_PHOTO1, NID_PHOTO2,BUS_DOC_PHOTO_1 FROM merchant WHERE MOBILE_NUMBER = :condition_value"""
                        cursor.execute(query, condition_value=merchant_data["number"])
                        merchant_data["images"] = cursor.fetchall()
            self.save_file()
            self.convert_to_pdf()
        except :
            print("Connection error")
                        
    def convert_to_pdf(self):
        print(self.target_dir)
        if(self.target_dir):
            os.makedirs(self.target_dir,exist_ok=True) 
            for merchant_data in self.required_numbers:
                path = os.path.join(self.target_dir, merchant_data["identifier"]) 
                os.makedirs(path, exist_ok=True)
                for img in merchant_data["images"][0]:
                    print(img)
                    if(img != None):
                        try:
                            image = Image.open(img)
                            new_pdf = path + '\\'+image.filename.split("\\")[-1].split(".")[0]+'.pdf'
                            file = open(new_pdf, "wb")
                            pdf_bytes = img2pdf.convert(image.filename,rotation=img2pdf.Rotation.ifvalid)
                            file.write(pdf_bytes)
                            image.close()
                            file.close()
                        except:
                            pass
        else:
            print("No Directory selected")

                
app = QApplication(sys.argv)
converter = converterWindow()
sys.exit(app.exec())