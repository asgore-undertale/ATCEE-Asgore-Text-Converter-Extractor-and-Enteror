import openpyxl

def Take_From_Table(database_directory):
    chars_table = {}
    wd = openpyxl.load_workbook(database_directory)
    database = wd.get_sheet_by_name("Database")
    
    for row in range(2, len(database['A'])+1):
        if database['A'+str(row)].value == 'آ':
            chars_table['ﺁ'] = database['E'+str(row)].value
            chars_table['ﺂ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'أ':
            chars_table['ﺃ'] = database['E'+str(row)].value
            chars_table['ﺄ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'ؤ':
            chars_table['ﺅ'] = database['E'+str(row)].value
            chars_table['ﺆ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'إ':
            chars_table['ﺇ'] = database['E'+str(row)].value
            chars_table['ﺈ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'ئ':
            chars_table['ﺉ'] = database['E'+str(row)].value
            chars_table['ﺊ'] = database['D'+str(row)].value
            chars_table['ﺌ'] = database['C'+str(row)].value
            chars_table['ﺋ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ا':
            chars_table['ﺍ'] = database['E'+str(row)].value
            chars_table['ﺎ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'ب':
            chars_table['ﺏ'] = database['E'+str(row)].value
            chars_table['ﺐ'] = database['D'+str(row)].value
            chars_table['ﺒ'] = database['C'+str(row)].value
            chars_table['ﺑ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ة':
            chars_table['ﺓ'] = database['E'+str(row)].value
            chars_table['ﺔ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'ت':
            chars_table['ﺕ'] = database['E'+str(row)].value
            chars_table['ﺖ'] = database['D'+str(row)].value
            chars_table['ﺘ'] = database['C'+str(row)].value
            chars_table['ﺗ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ث':
            chars_table['ﺙ'] = database['E'+str(row)].value
            chars_table['ﺚ'] = database['D'+str(row)].value
            chars_table['ﺜ'] = database['C'+str(row)].value
            chars_table['ﺛ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ج':
            chars_table['ﺝ'] = database['E'+str(row)].value
            chars_table['ﺞ'] = database['D'+str(row)].value
            chars_table['ﺠ'] = database['C'+str(row)].value
            chars_table['ﺟ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ح':
            chars_table['ﺡ'] = database['E'+str(row)].value
            chars_table['ﺢ'] = database['D'+str(row)].value
            chars_table['ﺤ'] = database['C'+str(row)].value
            chars_table['ﺣ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'خ':
            chars_table['ﺥ'] = database['E'+str(row)].value
            chars_table['ﺦ'] = database['D'+str(row)].value
            chars_table['ﺨ'] = database['C'+str(row)].value
            chars_table['ﺧ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'د':
            chars_table['ﺩ'] = database['E'+str(row)].value
            chars_table['ﺪ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'ذ':
            chars_table['ﺫ'] = database['E'+str(row)].value
            chars_table['ﺬ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'ر':
            chars_table['ﺭ'] = database['E'+str(row)].value
            chars_table['ﺮ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'ز':
            chars_table['ﺯ'] = database['E'+str(row)].value
            chars_table['ﺰ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'س':
            chars_table['ﺱ'] = database['E'+str(row)].value
            chars_table['ﺲ'] = database['D'+str(row)].value
            chars_table['ﺴ'] = database['C'+str(row)].value
            chars_table['ﺳ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ش':
            chars_table['ﺵ'] = database['E'+str(row)].value
            chars_table['ﺶ'] = database['D'+str(row)].value
            chars_table['ﺸ'] = database['C'+str(row)].value
            chars_table['ﺷ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ص':
            chars_table['ﺹ'] = database['E'+str(row)].value
            chars_table['ﺺ'] = database['D'+str(row)].value
            chars_table['ﺼ'] = database['C'+str(row)].value
            chars_table['ﺻ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ض':
            chars_table['ﺽ'] = database['E'+str(row)].value
            chars_table['ﺾ'] = database['D'+str(row)].value
            chars_table['ﻀ'] = database['C'+str(row)].value
            chars_table['ﺿ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ط':
            chars_table['ﻁ'] = database['E'+str(row)].value
            chars_table['ﻂ'] = database['D'+str(row)].value
            chars_table['ﻄ'] = database['C'+str(row)].value
            chars_table['ﻃ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ظ':
            chars_table['ﻅ'] = database['E'+str(row)].value
            chars_table['ﻅ'] = database['D'+str(row)].value
            chars_table['ﻈ'] = database['C'+str(row)].value
            chars_table['ﻇ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ع':
            chars_table['ﻉ'] = database['E'+str(row)].value
            chars_table['ﻊ'] = database['D'+str(row)].value
            chars_table['ﻋ'] = database['C'+str(row)].value
            chars_table['ﻌ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'غ':
            chars_table['ﻍ'] = database['E'+str(row)].value
            chars_table['ﻎ'] = database['D'+str(row)].value
            chars_table['ﻐ'] = database['C'+str(row)].value
            chars_table['ﻏ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ف':
            chars_table['ﻑ'] = database['E'+str(row)].value
            chars_table['ﻒ'] = database['D'+str(row)].value
            chars_table['ﻔ'] = database['C'+str(row)].value
            chars_table['ﻓ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ق':
            chars_table['ﻕ'] = database['E'+str(row)].value
            chars_table['ﻖ'] = database['D'+str(row)].value
            chars_table['ﻘ'] = database['C'+str(row)].value
            chars_table['ﻗ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ك':
            chars_table['ﻙ'] = database['E'+str(row)].value
            chars_table['ﻚ'] = database['D'+str(row)].value
            chars_table['ﻜ'] = database['C'+str(row)].value
            chars_table['ﻛ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ل':
            chars_table['ﻝ'] = database['E'+str(row)].value
            chars_table['ﻞ'] = database['D'+str(row)].value
            chars_table['ﻠ'] = database['C'+str(row)].value
            chars_table['ﻟ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'م':
            chars_table['ﻡ'] = database['E'+str(row)].value
            chars_table['ﻢ'] = database['D'+str(row)].value
            chars_table['ﻤ'] = database['C'+str(row)].value
            chars_table['ﻣ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ن':
            chars_table['ﻥ'] = database['E'+str(row)].value
            chars_table['ﻦ'] = database['D'+str(row)].value
            chars_table['ﻨ'] = database['C'+str(row)].value
            chars_table['ﻧ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'ه':
            chars_table['ﻩ'] = database['E'+str(row)].value
            chars_table['ﻪ'] = database['D'+str(row)].value
            chars_table['ﻬ'] = database['C'+str(row)].value
            chars_table['ﻫ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'و':
            chars_table['ﻭ'] = database['E'+str(row)].value
            chars_table['ﻮ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'ى':
            chars_table['ﻯ'] = database['E'+str(row)].value
            chars_table['ﻰ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'ي':
            chars_table['ﻱ'] = database['E'+str(row)].value
            chars_table['ﻲ'] = database['D'+str(row)].value
            chars_table['ﻴ'] = database['C'+str(row)].value
            chars_table['ﻳ'] = database['B'+str(row)].value
        elif database['A'+str(row)].value == 'لآ':
            chars_table['ﻵ'] = database['E'+str(row)].value
            chars_table['ﻶ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'لأ':
            chars_table['ﻷ'] = database['E'+str(row)].value
            chars_table['ﻸ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'لإ':
            chars_table['ﻹ'] = database['E'+str(row)].value
            chars_table['ﻺ'] = database['D'+str(row)].value
        elif database['A'+str(row)].value == 'لا':
            chars_table['ﻻ'] = database['E'+str(row)].value
            chars_table['ﻼ'] = database['D'+str(row)].value
        else:
            chars_table[database['A'+str(row)].value] = database['E'+str(row)].value
    
    return chars_table