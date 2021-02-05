import re
import openpyxl
from openpyxl.styles import Font, Alignment#, PatternFill
from os import path

def BMFont_to_FIB(xml_directory, database_directory):
    if not path.exists(xml_directory) or not path.exists(database_directory): return
    
    with open(xml_directory, 'r') as f: xml_content = f.read()
    chars_list = re.findall('<char id="(.*?)"', xml_content)
    width_list = re.findall('width="(.*?)"', xml_content)
    
    FIB_database = openpyxl.load_workbook(database_directory)
    database = FIB_database.get_sheet_by_name("Database")
    
    database.delete_cols(1, 5)
    database['A1'].value = "الحرف"
    database['A1'].font = Font(size = "14", bold=True)
    database['A1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    database['B1'].value = "أول"
    database['B1'].font = Font(size = "14", bold=True)
    database['B1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    database['C1'].value = "وسط"
    database['C1'].font = Font(size = "14", bold=True)
    database['C1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    database['D1'].value = "آخر"
    database['D1'].font = Font(size = "14", bold=True)
    database['D1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    database['E1'].value = "منفصل"
    database['E1'].font = Font(size = "14", bold=True)
    database['E1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    def put_in_table(column, row, char, Acolumn = True):
        if Acolumn:
            BMFont_to_FIB.row += 1
            database['A'+str(row)].value = char
            database['A'+str(row)].font = Font(size = "14", bold=True)
            database['A'+str(row)].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if database['A'+str(row)].value in '٠١٢٣٤٥٦٧٨٩0123456789':
                database['A'+str(row)].value = int(database['A'+str(row)].value)
        database[column+str(row)].value = int(j)
        database[column+str(row)].font = Font(size = "14", bold=True)
        database[column+str(row)].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    def check_for_arabic(main_char, b_char, column):
        if i == b_char:
            for z in range(2, len(database['A'])+1):
                if database['A'+str(z)].value == main_char:
                    put_in_table(column, z, main_char, False)
                    return True
            put_in_table(column, BMFont_to_FIB.row, main_char)
            return True
        return False
    
    BMFont_to_FIB.row = 2
    for i, j in zip(chars_list, width_list):
        i = chr(int(i))
        if check_for_arabic('ء', 'ء', 'E'): continue
        if check_for_arabic('ء', 'ﺀ', 'E'): continue
        if check_for_arabic('آ', 'آ', 'E'): continue
        if check_for_arabic('آ', 'ﺁ', 'E'): continue
        if check_for_arabic('آ', 'ﺂ', 'D'): continue
        if check_for_arabic('أ', 'أ', 'E'): continue
        if check_for_arabic('أ', 'ﺃ', 'E'): continue
        if check_for_arabic('أ', 'ﺄ', 'D'): continue
        if check_for_arabic('ؤ', 'ؤ', 'E'): continue
        if check_for_arabic('ؤ', 'ﺅ', 'E'): continue
        if check_for_arabic('ؤ', 'ﺆ', 'D'): continue
        if check_for_arabic('إ', 'إ', 'E'): continue
        if check_for_arabic('إ', 'ﺇ', 'E'): continue
        if check_for_arabic('إ', 'ﺈ', 'D'): continue
        if check_for_arabic('ئ', 'ئ', 'E'): continue
        if check_for_arabic('ئ', 'ﺉ', 'E'): continue
        if check_for_arabic('ئ', 'ﺊ', 'D'): continue
        if check_for_arabic('ئ', 'ﺌ', 'C'): continue
        if check_for_arabic('ئ', 'ﺋ', 'B'): continue
        if check_for_arabic('ا', 'ا', 'E'): continue
        if check_for_arabic('ا', 'ﺍ', 'E'): continue
        if check_for_arabic('ا', 'ﺎ', 'D'): continue
        if check_for_arabic('ب', 'ب', 'E'): continue
        if check_for_arabic('ب', 'ﺏ', 'E'): continue
        if check_for_arabic('ب', 'ﺐ', 'D'): continue
        if check_for_arabic('ب', 'ﺒ', 'C'): continue
        if check_for_arabic('ب', 'ﺑ', 'B'): continue
        if check_for_arabic('ة', 'ة', 'E'): continue
        if check_for_arabic('ة', 'ﺓ', 'E'): continue
        if check_for_arabic('ة', 'ﺔ', 'D'): continue
        if check_for_arabic('ت', 'ت', 'E'): continue
        if check_for_arabic('ت', 'ﺕ', 'E'): continue
        if check_for_arabic('ت', 'ﺖ', 'D'): continue
        if check_for_arabic('ت', 'ﺘ', 'C'): continue
        if check_for_arabic('ت', 'ﺗ', 'B'): continue
        if check_for_arabic('ث', 'ث', 'E'): continue
        if check_for_arabic('ث', 'ﺙ', 'E'): continue
        if check_for_arabic('ث', 'ﺚ', 'D'): continue
        if check_for_arabic('ث', 'ﺜ', 'C'): continue
        if check_for_arabic('ث', 'ﺛ', 'B'): continue
        if check_for_arabic('ج', 'ج', 'E'): continue
        if check_for_arabic('ج', 'ﺝ', 'E'): continue
        if check_for_arabic('ج', 'ﺞ', 'D'): continue
        if check_for_arabic('ج', 'ﺠ', 'C'): continue
        if check_for_arabic('ج', 'ﺟ', 'B'): continue
        if check_for_arabic('ح', 'ح', 'E'): continue
        if check_for_arabic('ح', 'ﺡ', 'E'): continue
        if check_for_arabic('ح', 'ﺢ', 'D'): continue
        if check_for_arabic('ح', 'ﺤ', 'C'): continue
        if check_for_arabic('ح', 'ﺣ', 'B'): continue
        if check_for_arabic('خ', 'خ', 'E'): continue
        if check_for_arabic('خ', 'ﺥ', 'E'): continue
        if check_for_arabic('خ', 'ﺦ', 'D'): continue
        if check_for_arabic('خ', 'ﺨ', 'C'): continue
        if check_for_arabic('خ', 'ﺧ', 'B'): continue
        if check_for_arabic('د', 'د', 'E'): continue
        if check_for_arabic('د', 'ﺩ', 'E'): continue
        if check_for_arabic('د', 'ﺪ', 'D'): continue
        if check_for_arabic('ذ', 'ذ', 'E'): continue
        if check_for_arabic('ذ', 'ﺫ', 'E'): continue
        if check_for_arabic('ذ', 'ﺬ', 'D'): continue
        if check_for_arabic('ر', 'ر', 'E'): continue
        if check_for_arabic('ر', 'ﺭ', 'E'): continue
        if check_for_arabic('ر', 'ﺮ', 'D'): continue
        if check_for_arabic('ز', 'ز', 'E'): continue
        if check_for_arabic('ز', 'ﺯ', 'E'): continue
        if check_for_arabic('ز', 'ﺰ', 'D'): continue
        if check_for_arabic('س', 'س', 'E'): continue
        if check_for_arabic('س', 'ﺱ', 'E'): continue
        if check_for_arabic('س', 'ﺲ', 'D'): continue
        if check_for_arabic('س', 'ﺴ', 'C'): continue
        if check_for_arabic('س', 'ﺳ', 'B'): continue
        if check_for_arabic('ش', 'ش', 'E'): continue
        if check_for_arabic('ش', 'ﺵ', 'E'): continue
        if check_for_arabic('ش', 'ﺶ', 'D'): continue
        if check_for_arabic('ش', 'ﺸ', 'C'): continue
        if check_for_arabic('ش', 'ﺷ', 'B'): continue
        if check_for_arabic('ص', 'ص', 'E'): continue
        if check_for_arabic('ص', 'ﺹ', 'E'): continue
        if check_for_arabic('ص', 'ﺺ', 'D'): continue
        if check_for_arabic('ص', 'ﺼ', 'C'): continue
        if check_for_arabic('ص', 'ﺻ', 'B'): continue
        if check_for_arabic('ض', 'ض', 'E'): continue
        if check_for_arabic('ض', 'ﺽ', 'E'): continue
        if check_for_arabic('ض', 'ﺾ', 'D'): continue
        if check_for_arabic('ض', 'ﻀ', 'C'): continue
        if check_for_arabic('ض', 'ﺿ', 'B'): continue
        if check_for_arabic('ط', 'ط', 'E'): continue
        if check_for_arabic('ط', 'ﻁ', 'E'): continue
        if check_for_arabic('ط', 'ﻂ', 'D'): continue
        if check_for_arabic('ط', 'ﻄ', 'C'): continue
        if check_for_arabic('ط', 'ﻃ', 'B'): continue
        if check_for_arabic('ظ', 'ظ', 'E'): continue
        if check_for_arabic('ظ', 'ﻅ', 'E'): continue
        if check_for_arabic('ظ', 'ﻆ', 'D'): continue
        if check_for_arabic('ظ', 'ﻈ', 'C'): continue
        if check_for_arabic('ظ', 'ﻇ', 'B'): continue
        if check_for_arabic('ع', 'ع', 'E'): continue
        if check_for_arabic('ع', 'ﻉ', 'E'): continue
        if check_for_arabic('ع', 'ﻊ', 'D'): continue
        if check_for_arabic('ع', 'ﻋ', 'C'): continue
        if check_for_arabic('ع', 'ﻌ', 'B'): continue
        if check_for_arabic('غ', 'غ', 'E'): continue
        if check_for_arabic('غ', 'ﻍ', 'E'): continue
        if check_for_arabic('غ', 'ﻎ', 'D'): continue
        if check_for_arabic('غ', 'ﻐ', 'C'): continue
        if check_for_arabic('غ', 'ﻏ', 'B'): continue
        if check_for_arabic('ف', 'ف', 'E'): continue
        if check_for_arabic('ف', 'ﻑ', 'E'): continue
        if check_for_arabic('ف', 'ﻒ', 'D'): continue
        if check_for_arabic('ف', 'ﻔ', 'C'): continue
        if check_for_arabic('ف', 'ﻓ', 'B'): continue
        if check_for_arabic('ق', 'ق', 'E'): continue
        if check_for_arabic('ق', 'ﻕ', 'E'): continue
        if check_for_arabic('ق', 'ﻖ', 'D'): continue
        if check_for_arabic('ق', 'ﻘ', 'C'): continue
        if check_for_arabic('ق', 'ﻗ', 'B'): continue
        if check_for_arabic('ك', 'ك', 'E'): continue
        if check_for_arabic('ك', 'ﻙ', 'E'): continue
        if check_for_arabic('ك', 'ﻚ', 'D'): continue
        if check_for_arabic('ك', 'ﻜ', 'C'): continue
        if check_for_arabic('ك', 'ﻛ', 'B'): continue
        if check_for_arabic('ل', 'ل', 'E'): continue
        if check_for_arabic('ل', 'ﻝ', 'E'): continue
        if check_for_arabic('ل', 'ﻞ', 'D'): continue
        if check_for_arabic('ل', 'ﻠ', 'C'): continue
        if check_for_arabic('ل', 'ﻟ', 'B'): continue
        if check_for_arabic('م', 'م', 'E'): continue
        if check_for_arabic('م', 'ﻡ', 'E'): continue
        if check_for_arabic('م', 'ﻢ', 'D'): continue
        if check_for_arabic('م', 'ﻤ', 'C'): continue
        if check_for_arabic('م', 'ﻣ', 'B'): continue
        if check_for_arabic('ن', 'ن', 'E'): continue
        if check_for_arabic('ن', 'ﻥ', 'E'): continue
        if check_for_arabic('ن', 'ﻦ', 'D'): continue
        if check_for_arabic('ن', 'ﻨ', 'C'): continue
        if check_for_arabic('ن', 'ﻧ', 'B'): continue
        if check_for_arabic('ه', 'ه', 'E'): continue
        if check_for_arabic('ه', 'ﻩ', 'E'): continue
        if check_for_arabic('ه', 'ﻪ', 'D'): continue
        if check_for_arabic('ه', 'ﻬ', 'C'): continue
        if check_for_arabic('ه', 'ﻫ', 'B'): continue
        if check_for_arabic('و', 'و', 'E'): continue
        if check_for_arabic('و', 'ﻭ', 'E'): continue
        if check_for_arabic('و', 'ﻮ', 'D'): continue
        if check_for_arabic('ى', 'ى', 'E'): continue
        if check_for_arabic('ى', 'ﻯ', 'E'): continue
        if check_for_arabic('ى', 'ﻰ', 'D'): continue
        if check_for_arabic('ي', 'ي', 'E'): continue
        if check_for_arabic('ي', 'ﻱ', 'E'): continue
        if check_for_arabic('ي', 'ﻲ', 'D'): continue
        if check_for_arabic('ي', 'ﻴ', 'C'): continue
        if check_for_arabic('ي', 'ﻳ', 'B'): continue
        if check_for_arabic('لآ', 'لآ', 'E'): continue
        if check_for_arabic('لآ', 'ﻵ', 'E'): continue
        if check_for_arabic('لآ', 'ﻶ', 'D'): continue
        if check_for_arabic('لأ', 'لأ', 'E'): continue
        if check_for_arabic('لأ', 'ﻷ', 'E'): continue
        if check_for_arabic('لأ', 'ﻸ', 'D'): continue
        if check_for_arabic('لإ', 'لإ', 'E'): continue
        if check_for_arabic('لإ', 'ﻹ', 'E'): continue
        if check_for_arabic('لإ', 'ﻺ', 'D'): continue
        if check_for_arabic('لا', 'لا', 'E'): continue
        if check_for_arabic('لا', 'ﻻ', 'E'): continue
        if check_for_arabic('لا', 'ﻼ', 'D'): continue
        else: put_in_table('E', BMFont_to_FIB.row, i)
    FIB_database.save(database_directory)
