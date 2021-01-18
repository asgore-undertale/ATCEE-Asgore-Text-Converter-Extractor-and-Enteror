#استيراد المكتبات
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QCheckBox, QPushButton, QLabel, QMenuBar, QAction, QFileDialog, QMessageBox, QRadioButton
from sys import argv, exit
from os import path, listdir
import openpyxl

##
app = QApplication(argv)

##الخطوط
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
OptionsWindow_Width = 440
checkbox_size = [OptionsWindow_Width-5, 16]
textedit_size = [30, 26]
def pos_y(line_num, Height = checkbox_size[1], Between_every_y = 20):
    y = (Between_every_y+checkbox_size[1]) * line_num + (Between_every_y-Height//2)
    return y

OptionsWindow = QMainWindow()
OptionsWindow.setFixedSize(OptionsWindow_Width, 330)
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
##
UC_database_button = QPushButton(OptionsWindow)
UC_database_button.setGeometry(QtCore.QRect(25, 190, 93, 56))
UC_database_button.setText("قاعدة بيانات\nالتحويل")
##

RT_check = QCheckBox("عكس النص", OptionsWindow)
RT_check.setGeometry(QtCore.QRect(0, pos_y(7), checkbox_size[0], checkbox_size[1]))
RT_check.setLayoutDirection(QtCore.Qt.RightToLeft)
RAO_check = QCheckBox("عكس العربية في النص", OptionsWindow)
RAO_check.setGeometry(QtCore.QRect(0, pos_y(8), checkbox_size[0], checkbox_size[1]))
RAO_check.setLayoutDirection(QtCore.Qt.RightToLeft)
##
RT_end_command = QTextEdit(OptionsWindow)
RT_end_command.setGeometry(QtCore.QRect(5, 262, 30, 26))
RT_end_label = QLabel(OptionsWindow)
RT_end_label.setGeometry(QtCore.QRect(40, 262, 35, 26))
RT_end_label.setText("بعدها:")
RT_start_command = QTextEdit(OptionsWindow)
RT_start_command.setGeometry(QtCore.QRect(110, 262, 30, 26))
RT_start_label = QLabel(OptionsWindow)
RT_start_label.setGeometry(QtCore.QRect(145, 262, 60, 26))
RT_start_label.setText("قبل الأوامر:")

RT_line_command = QTextEdit(OptionsWindow)
RT_line_command.setGeometry(QtCore.QRect(5, 295, 30, 26))
RT_line_label = QLabel(OptionsWindow)
RT_line_label.setGeometry(QtCore.QRect(40, 295, 65, 26))
RT_line_label.setText("سطر جديد:")
RT_page_command = QTextEdit(OptionsWindow)
RT_page_command.setGeometry(QtCore.QRect(110, 295, 30, 26))
RT_page_label = QLabel(OptionsWindow)
RT_page_label.setGeometry(QtCore.QRect(145, 295, 95, 26))
RT_page_label.setText("أمر صفحة جديدة:")
##


#النافذة الرئيسية
MainWindow = QMainWindow()
MainWindow.setFixedSize(756, 344)
MainWindow.setWindowTitle("AST 1.0")

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
EnteringWindow.setFixedSize(756, 330)
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
text_database_button = QPushButton(EnteringWindow)
text_database_button.setGeometry(QtCore.QRect(330, 130, 93, 41))
text_database_button.setText("فتح قاعدة\nبيانات النصوص")

input_from_folder = QPushButton(EnteringWindow)
input_from_folder.setGeometry(QtCore.QRect(130, 200, 93, 41))
input_from_folder.setText("المجلد الحاوي\nللملفات")
output_from_folder = QPushButton(EnteringWindow)
output_from_folder.setGeometry(QtCore.QRect(20, 200, 93, 41))
output_from_folder.setText("مجلد\nالاستخراج")

label = QLabel(EnteringWindow)
label.setGeometry(QtCore.QRect(654, 10, 81, 20))
label.setFont(labels_font)
label.setText("النص الأصلي:")
label_2 = QLabel(EnteringWindow)
label_2.setGeometry(QtCore.QRect(220, 10, 81, 20))
label_2.setFont(labels_font)
label_2.setText("الترجمة:")

database_check = QCheckBox("استخدام قاعدة البيانات", EnteringWindow)
database_check.setGeometry(QtCore.QRect(570, 190, 150, 16))
database_check.setLayoutDirection(QtCore.Qt.RightToLeft)

after_text = QTextEdit(EnteringWindow)
after_text.setGeometry(QtCore.QRect(420, 218, 30, 26))
after_text_label = QLabel(EnteringWindow)
after_text_label.setGeometry(QtCore.QRect(450, 218, 60, 26))
after_text_label.setText("ما يلحقه:")
before_text = QTextEdit(EnteringWindow)
before_text.setGeometry(QtCore.QRect(520, 218, 30, 26))
before_text_label = QLabel(EnteringWindow)
before_text_label.setGeometry(QtCore.QRect(550, 218, 160, 26))
before_text_label.setText("ما يسبق النص في الملفات:")

too_long_check = QCheckBox("عدم إدخال ترجمات أطول من النص الأصلي (بقيم الهيكس)", EnteringWindow)
too_long_check.setGeometry(QtCore.QRect(270, 255,450, 16))
too_long_check.setLayoutDirection(QtCore.Qt.RightToLeft)
translation_place_check = QCheckBox(":مكان الترجمة في حال كانت أقصر", EnteringWindow)
translation_place_check.setGeometry(QtCore.QRect(270, 280,450, 16))
translation_place_check.setLayoutDirection(QtCore.Qt.RightToLeft)

first_radio = QRadioButton(EnteringWindow)
first_radio.setGeometry(QtCore.QRect(220, 305,450, 16))
first_radio.setText("أول")
first_radio.setLayoutDirection(QtCore.Qt.RightToLeft)
middle_radio = QRadioButton(EnteringWindow)
middle_radio.setGeometry(QtCore.QRect(150, 305,450, 16))
middle_radio.setText("وسط")
middle_radio.setLayoutDirection(QtCore.Qt.RightToLeft)
last_radio = QRadioButton(EnteringWindow)
last_radio.setGeometry(QtCore.QRect(80, 305,450, 16))
last_radio.setText("آخر")
last_radio.setLayoutDirection(QtCore.Qt.RightToLeft)


##المتغيرات
converting_database_directory = 'Scripts/Un-Converting_Database.xlsx'
text_database_directory = 'جدول النصوص.xlsx'
input_folder, output_folder = 'المجلد الحاوي للملفات/', 'مجلد الاستخراج/'

#الدوال

def open_textfile():
    fileName, _ = QFileDialog.getOpenFileName(MainWindow, 'ملف نص', '' , '*')
    if path.exists(fileName):
        entered_text.setPlainText(open(fileName, 'r', encoding='utf-8').read())
    QMessageBox.about(OptionsWindow, "!!تهانيّ", "تم اختيار ملف النص.")

def open_convert_database():
    fileName, _ = QFileDialog.getOpenFileName(OptionsWindow, 'قاعدة بيانات التحويل', '' , '*.xlsx')
    if path.exists(fileName):
        global converting_database_directory
        converting_database_directory = fileName
    QMessageBox.about(OptionsWindow, "!!تهانيّ", "تم اختيار قاعدة البيانات.")

def open_text_database():
    fileName, _ = QFileDialog.getOpenFileName(EnteringWindow, 'قاعدة بيانات النص', '' , '*.xlsx')
    if path.exists(fileName):
        global text_database_directory
    QMessageBox.about(EnteringWindow, "!!تهانيّ", "تم اختيار قاعدة البيانات.")

def open_folder(case='input'):
    folder = str(QFileDialog.getExistingDirectory(EnteringWindow, "Select Directory"))+'/'
    if path.exists(folder):
        if case == 'input':
            global input_folder
            input_folder = folder
        else:
            global output_folder
            output_folder = folder
    QMessageBox.about(EnteringWindow, "!!تهانيّ", "تم اختيار المجلد.")

def convert(text):
    ##إلغاء العملية في حال تحقق إحدى هذه الشروط
    if text == '': return
    if C_check.isChecked() or UC_check.isChecked():
        if not path.exists(converting_database_directory):
            QMessageBox.about(MainWindow, "!!خطأ", "قاعدة بيانات التحويل غير موجودة،\nتم إيقاف كل العمليات.")
            return
    
    if DDL_check.isChecked():#Delete Duplicated lines
        from Scripts.Delete_Duplicated_lines import script
        text = script(text)
    
    if SSL_check.isChecked():#Sort short to long
        from Scripts.Sort_lines import script
        text = script(text)
    
    if SLS_check.isChecked():#Sort long to short
        from Scripts.Sort_lines import script
        text = script(text, 'long to short')
    
    if RA_check.isChecked() or C_check.isChecked():#Reshape Arabic
        from Scripts.Re_Unshape_Arabic import script
        text = script(text)
    
    if C_check.isChecked():#Convert
        from Scripts.Un_Convert import script
        text = script(text, 'convert', converting_database_directory)
    
    if UC_check.isChecked():#Unconvert
        from Scripts.Un_Convert import script
        text = script(text, 'unconvert', converting_database_directory)
        
    if UA_check.isChecked() or UC_check.isChecked():#Unshape Arabic
        from Scripts.Re_Unshape_Arabic import script
        text = script(text, 'unshape')
        
    if RT_check.isChecked():#Reverse whole text
        from Scripts.Reverse_text import script
        text = script(text, RT_start_command.toPlainText(), RT_end_command.toPlainText(), RT_page_command.toPlainText(), RT_line_command.toPlainText())
        
    if RAO_check.isChecked():#‫Reverse Arabic only
        from Scripts.Reverse_text import script
        text = script(text, RT_start_command.toPlainText(), RT_end_command.toPlainText(), RT_page_command.toPlainText(), RT_line_command.toPlainText(), 'Arabic')
    
    return text

def enter(convert_bool=True):
    ##المتغيرات
    text_dic = {}
    too_long_dic = {}
    no_found_list = []
    found_list = []
    no_found_log = ''
    too_long_log = ''
    
    ##إلغاء العملية في حال تحقق إحدى هذه الشروط
    if C_check.isChecked() or UC_check.isChecked():
        if not path.exists(converting_database_directory):
            QMessageBox.about(EnteringWindow, "!!خطأ", "تم إيقاف كل العمليات،\nقاعدة بيانات التحويل غير موجودة.")
            return
    if not path.exists(input_folder):
        QMessageBox.about(EnteringWindow, "!!خطأ", "تم إيقاف كل العمليات،\nالمجلد الحاوي للملفات غير موجود.")
        return
    if not path.exists(output_folder):
        QMessageBox.about(EnteringWindow, "!!خطأ", "تم إيقاف كل العمليات،\nمجلد الاستخراج غير موجود.")
        return
    files_list = listdir(input_folder)
    if len(files_list) == 0:
        QMessageBox.about(EnteringWindow, "!!خطأ", "تم إيقاف كل العمليات،\nلا توجد أي ملفات للإدخال إليها.")
        return
    if database_check.isChecked():
        if not path.exists(text_database_directory):
            QMessageBox.about(EnteringWindow, "!!خطأ", "تم إيقاف كل العمليات،\nقاعدة بيانات النصوص غير موجودة.")
            return
        text_xlsx = openpyxl.load_workbook(text_database_directory)
        text_table = text_xlsx.get_sheet_by_name("Main")
        for cell in range(2, len(text_table['A'])+1):
            original_cell_value = text_table['A'+str(cell)].value
            translate_cell_value = text_table['B'+str(cell)].value
            
            if original_cell_value in text_dic:
                if text_dic[original_cell_value] == '' or text_dic[original_cell_value] == None:
                    text_dic[original_cell_value] = translate_cell_value
            else:
                text_dic[original_cell_value] = translate_cell_value
        
        new_d = {}
        for k in sorted(text_dic, key=len, reverse=True):
            new_d[k] = text_dic[k]
        text_dic = new_d
    else:
        if original_text.toPlainText() == '': return
        text_dic[original_text.toPlainText()] = translate_text.toPlainText()
    
    for filename in files_list:
        file_content = open(input_folder+filename, 'rb').read()
        
        for text, translation in text_dic.items():
            text = before_text.toPlainText() + text + after_text.toPlainText()
            translation = before_text.toPlainText() + translation + after_text.toPlainText()
            
            if translation_place_check.isChecked() and len(translation.encode('utf-8').hex()) < len(text.encode('utf-8').hex()):
                spaces_count = (len(text.encode('utf-8').hex()) // 2) - (len(translation.encode('utf-8').hex()) // 2)
                if first_radio.isChecked():#first
                    for i in range(spaces_count):
                        translation += ' '
                elif middle_radio.isChecked():#middle
                    for i in range(spaces_count):
                        if i % 2 == 0:
                            translation += ' '
                        else:
                            translation = ' ' + translation
                elif last_radio.isChecked():#last
                    for i in range(spaces_count):
                        translation = ' ' + translation
            
            if bytes(text, 'utf-8') in file_content:
                if convert_bool: translation = convert(translation)
                if too_long_check.isChecked() and len(translation.encode('utf-8').hex()) > len(text.encode('utf-8').hex()):
                    too_long_dic[text] = translation
                else:
                    file_content = file_content.replace(bytes(text, 'utf-8'), bytes(translation, 'utf-8'))
                    found_list.append(text)
                    if text in no_found_list: no_found_list.remove(text)
            else:
                if text not in no_found_log and text not in found_list:
                    no_found_list.append(text)
        
        open(output_folder+filename, 'wb').write(file_content)
    
    for item in no_found_list:
        no_found_log += '> ' + text + '\n'
    for k, v in too_long_dic.items():
        too_long_log += '> ' + k + '\n    ' + k.encode('utf-8').hex() + '\n    ' + v + '\n    ' + v.encode('utf-8').hex() + '\n\n'
    
    if no_found_log == '' and too_long_log == '':
        QMessageBox.about(EnteringWindow, "!!تهانيّ", "انتهى الإدخال.")
    if no_found_log != '':
        msg = QMessageBox()
        msg.setText(no_found_log)
        msg.setWindowTitle("ما لم يتم إيجاده")
        msg.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse);
        msg.exec_()
    if too_long_log != '':
        msg = QMessageBox()
        msg.setText(too_long_log)
        msg.setWindowTitle("أطول من النصوص الأصلية")
        msg.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse);
        msg.exec_()
    

##توصيل الإشارات
convert_button.clicked.connect(lambda: result_text.setPlainText(convert(entered_text.toPlainText())))
openfile_button.clicked.connect(lambda: open_textfile())
converting_options.triggered.connect(lambda: OptionsWindow.show())
entering.triggered.connect(lambda: EnteringWindow.show())
about.triggered.connect(lambda: AboutWindow.show())

UC_database_button.clicked.connect(lambda: open_convert_database())
text_database_button.clicked.connect(lambda: open_text_database())

enter_button.clicked.connect(lambda: enter(False))
convert_enter_button.clicked.connect(lambda: enter())
input_from_folder.clicked.connect(lambda: open_folder('input'))
output_from_folder.clicked.connect(lambda: open_folder('output'))

#تشغيل البرنامج
MainWindow.show()
exit(app.exec_())