#استيراد المكتبات
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QCheckBox, QPushButton, QLabel, QMenuBar, QAction, QFileDialog, QMessageBox
from sys import argv, exit
from os import path, listdir

##
app = QApplication(argv)
labels_font = QtGui.QFont()
labels_font.setPointSize(9)
textbox_font = QtGui.QFont()
textbox_font.setPointSize(12)

#نافذة عني
AboutWindow = QMainWindow()
AboutWindow.setFixedSize(378, 146)
AboutWindow.setWindowTitle("عني")

about_textbox = QTextEdit(AboutWindow)
about_textbox.setGeometry(QtCore.QRect(0, 0, 378, 146))
about_textbox.setFont(labels_font)
about_textbox.setReadOnly(True)
about_textbox.setText("طوّرت هذه الأداة من قبل Asgore_Undertale\nصفحتي على github:\nhttps://github.com/asgore-undertale\nلك كامل الحرية في التعديل والنشر،\nبشرط ذكري وصفحتي.")


#نافذة خيارات التحويل
OptionsWindow_Width = 400
checkbox_size = [OptionsWindow_Width-5, 16]
textedit_size = [30, 26]
def pos_y(line_num, Height = checkbox_size[1], Between_every_y = 20):
    y = (Between_every_y+checkbox_size[1]) * line_num + (Between_every_y-Height//2)
    return y

OptionsWindow = QMainWindow()
OptionsWindow.setFixedSize(OptionsWindow_Width, 344)
OptionsWindow.setWindowTitle("خيارات التحويل")

DDL_check = QCheckBox("حذف الأسطر المكررة", OptionsWindow)
DDL_check.setGeometry(QtCore.QRect(0, pos_y(0), checkbox_size[0], checkbox_size[1]))
DDL_check.setLayoutDirection(QtCore.Qt.RightToLeft)

SSL_check = QCheckBox("ترتيب السطور من الأقصر للأطول", OptionsWindow)
SSL_check.setGeometry(QtCore.QRect(0, pos_y(1), checkbox_size[0], checkbox_size[1]))
SSL_check.setLayoutDirection(QtCore.Qt.RightToLeft)

SLS_check = QCheckBox("ترتيب السطور من الأطول للأقصر", OptionsWindow)
SLS_check.setGeometry(QtCore.QRect(0, pos_y(2), checkbox_size[0], checkbox_size[1]))
SLS_check.setLayoutDirection(QtCore.Qt.RightToLeft)

RA_check = QCheckBox("تجميد النص العربي", OptionsWindow)
RA_check.setGeometry(QtCore.QRect(0, pos_y(3), checkbox_size[0], checkbox_size[1]))
RA_check.setLayoutDirection(QtCore.Qt.RightToLeft)

UA_check = QCheckBox("إلغاء تجميد النص العربي", OptionsWindow)
UA_check.setGeometry(QtCore.QRect(0, pos_y(4), checkbox_size[0], checkbox_size[1]))
UA_check.setLayoutDirection(QtCore.Qt.RightToLeft)

C_check = QCheckBox("تحويل النص", OptionsWindow)
C_check.setGeometry(QtCore.QRect(0, pos_y(5), checkbox_size[0], checkbox_size[1]))
C_check.setLayoutDirection(QtCore.Qt.RightToLeft)

UC_check = QCheckBox("إلغاء تحويل النص", OptionsWindow)
UC_check.setGeometry(QtCore.QRect(0, pos_y(6), checkbox_size[0], checkbox_size[1]))
UC_check.setLayoutDirection(QtCore.Qt.RightToLeft)
UC_database_button = QPushButton(OptionsWindow)
UC_database_button.setGeometry(QtCore.QRect(5, 190, 93, 56))
UC_database_button.setText("قاعدة بيانات")

RT_check = QCheckBox("عكس النص", OptionsWindow)
RT_check.setGeometry(QtCore.QRect(0, pos_y(7), checkbox_size[0], checkbox_size[1]))
RT_check.setLayoutDirection(QtCore.Qt.RightToLeft)
RT_end_command = QTextEdit(OptionsWindow)
RT_end_command.setGeometry(QtCore.QRect(5, 275, 30, 26))
RT_end_label = QLabel(OptionsWindow)
RT_end_label.setGeometry(QtCore.QRect(40, 275, 30, 26))
RT_end_label.setText("بعدها:")
RT_start_command = QTextEdit(OptionsWindow)
RT_start_command.setGeometry(QtCore.QRect(75, 275, 30, 26))
RT_start_label = QLabel(OptionsWindow)
RT_start_label.setGeometry(QtCore.QRect(110, 275, 60, 26))
RT_start_label.setText("قبل الأوامر:")

RAO_check = QCheckBox("عكس العربية في النص", OptionsWindow)
RAO_check.setGeometry(QtCore.QRect(0, pos_y(8), checkbox_size[0], checkbox_size[1]))
RAO_check.setLayoutDirection(QtCore.Qt.RightToLeft)


#النافذة الرئيسية
MainWindow = QMainWindow()
MainWindow.setFixedSize(756, 344)
MainWindow.setWindowTitle("محوّل النصوص ومدخلها 1.0")

result_text = QTextEdit(MainWindow)
result_text.setGeometry(QtCore.QRect(13, 62, 301, 253))
result_text.setFont(textbox_font)
entered_text = QTextEdit(MainWindow)
entered_text.setGeometry(QtCore.QRect(440, 66, 301, 253))
entered_text.setFont(textbox_font)

convert_button = QPushButton(MainWindow)
convert_button.setGeometry(QtCore.QRect(330, 116, 93, 28))
convert_button.setText("تحويل")
openfile_button = QPushButton(MainWindow)
openfile_button.setGeometry(QtCore.QRect(330, 166, 93, 41))
openfile_button.setText("فتح ملف\nنص")

label = QLabel(MainWindow)
label.setGeometry(QtCore.QRect(654, 36, 81, 20))
label.setFont(labels_font)
label.setText("النص الداخل:")
label_2 = QLabel(MainWindow)
label_2.setGeometry(QtCore.QRect(220, 36, 81, 20))
label_2.setFont(labels_font)
label_2.setText("النص الناتج:")

menubar = QMenuBar(MainWindow)
menubar.setGeometry(QtCore.QRect(0, 0, 756, 26))
menubar.setLayoutDirection(QtCore.Qt.RightToLeft)
converting_options = QAction("خيارات التحويل", MainWindow)
entering = QAction("إدخال", MainWindow)
about = QAction("عني", MainWindow)
menubar.addAction(converting_options)
menubar.addAction(entering)
menubar.addAction(about)


#نافذة الإدخال
EnteringWindow = QMainWindow()
EnteringWindow.setFixedSize(756, 344)
EnteringWindow.setWindowTitle("نافذة الإدخال")

translate_text = QTextEdit(EnteringWindow)
translate_text.setGeometry(QtCore.QRect(13, 34, 301, 140))
translate_text.setFont(textbox_font)
original_text = QTextEdit(EnteringWindow)
original_text.setGeometry(QtCore.QRect(440, 40, 301, 140))
original_text.setFont(textbox_font)

enter_button = QPushButton(EnteringWindow)
enter_button.setGeometry(QtCore.QRect(330, 45, 93, 28))
enter_button.setText("إدخال")
convert_enter_button = QPushButton(EnteringWindow)
convert_enter_button.setGeometry(QtCore.QRect(330, 80, 93, 41))
convert_enter_button.setText("تحويل\nوإدخال")
open_database_button = QPushButton(EnteringWindow)
open_database_button.setGeometry(QtCore.QRect(330, 130, 93, 41))
open_database_button.setText("فتح قاعدة\nبيانات")

import_folder = QPushButton(EnteringWindow)
import_folder.setGeometry(QtCore.QRect(130, 200, 93, 41))
import_folder.setText("المجلد الحاوي\nللملفات")
export_folder = QPushButton(EnteringWindow)
export_folder.setGeometry(QtCore.QRect(20, 200, 93, 41))
export_folder.setText("مجلد\nالاستخراج")

label = QLabel(EnteringWindow)
label.setGeometry(QtCore.QRect(654, 10, 81, 20))
label.setFont(labels_font)
label.setText("النص الأصلي:")
label_2 = QLabel(EnteringWindow)
label_2.setGeometry(QtCore.QRect(220, 10, 81, 20))
label_2.setFont(labels_font)
label_2.setText("الترجمة:")

database_check = QCheckBox("استخدام قاعدة البيانات", EnteringWindow)
database_check.setGeometry(QtCore.QRect(570, 200, 150, 16))
database_check.setLayoutDirection(QtCore.Qt.RightToLeft)


##المتغيرات
#[Delete Duplicated lines, Sort short to long, Sort long to short, Reshape Arabic, Unshape Arabic, Convert, Unconvert,
# Reverse whole text, Reverse Arabic only]
converting_options_list = [False, False, False, False, False, False, False, False, False]
database_directory = 'SampleScripts/Un-Converting_Database.xlsx'

#الدوال
def being_ready_to_start():
    ##إلغاء العملية في حال تحقق إحدى هذه الشروط
    if text == '': return
    if converting_options_list[5] or converting_options_list[6]:
        if not path.exists(database_directory):
            QMessageBox.about(MainWindow, "!!خطأ", "قاعدة بيانات التحويل غير موجودة،\nتم إيقاف كل العمليات.")
            return

def check_options(check_box, cell):
    converting_options_list[cell] = check_box.isChecked()

def open_textfile():
    fileName, _ = QFileDialog.getOpenFileName(MainWindow, 'ملف نص', '' , '*')
    if path.exists(fileName):
        entered_text.setPlainText(open(fileName, 'r', encoding='utf-8').read())

def open_convert_database():
    fileName, _ = QFileDialog.getOpenFileName(MainWindow, 'قاعدة بيانات', '' , '*.xlsx')
    if path.exists(fileName):
        global database_directory
        database_directory = fileName

def convert(text):
    being_ready_to_start()
    
    if converting_options_list[0]:#Delete Duplicated lines
        from SampleScripts.Delete_Duplicated_lines import script
        text = script(text)
    
    if converting_options_list[1]:#Sort short to long
        from SampleScripts.Sort_lines import script
        text = script(text)
    
    if converting_options_list[2]:#Sort long to short
        from SampleScripts.Sort_lines import script
        text = script(text, 'long to short')
    
    if converting_options_list[3] or converting_options_list[5]:#Reshape Arabic
        from SampleScripts.Re_Unshape_Arabic import script
        text = script(text)
    
    if converting_options_list[5]:#Convert
        from SampleScripts.Un_Convert import script
        text = script(text, 'convert', database_directory)
    
    if converting_options_list[6]:#Unconvert
        from SampleScripts.Un_Convert import script
        text = script(text, 'unconvert', database_directory)
        
    if converting_options_list[4] or converting_options_list[6]:#Unshape Arabic
        from SampleScripts.Re_Unshape_Arabic import script
        text = script(text, 'unshape')
        
    if converting_options_list[7]:#Reverse whole text
        from SampleScripts.Reverse_text import script
        text = script(text, RT_start_command.toPlainText(), RT_end_command.toPlainText())
        
    if converting_options_list[8]:#‫Reverse Arabic only
        from SampleScripts.Reverse_text import script
        text = script(text, RT_start_command.toPlainText(), RT_end_command.toPlainText(), 'Arabic')
    
    return text

def enter():
    being_ready_to_start()
    

##توصيل الإشارات
convert_button.clicked.connect(lambda: result_text.setPlainText(convert(entered_text.toPlainText())))
openfile_button.clicked.connect(lambda: open_textfile())
converting_options.triggered.connect(lambda: OptionsWindow.show())
entering.triggered.connect(lambda: EnteringWindow.show())
about.triggered.connect(lambda: AboutWindow.show())

UC_database_button.clicked.connect(lambda: open_convert_database())

DDL_check.toggled.connect(lambda: check_options(DDL_check, 0))
SSL_check.toggled.connect(lambda: check_options(SSL_check, 1))
SLS_check.toggled.connect(lambda: check_options(SLS_check, 2))
RA_check.toggled.connect (lambda: check_options(RA_check , 3))
UA_check.toggled.connect (lambda: check_options(UA_check , 4))
C_check.toggled.connect  (lambda: check_options(C_check  , 5))
UC_check.toggled.connect (lambda: check_options(UC_check , 6))
RT_check.toggled.connect  (lambda: check_options(RT_check  , 7))
RAO_check.toggled.connect (lambda: check_options(RAO_check , 8))

#النوافذ
MainWindow.show()
exit(app.exec_())