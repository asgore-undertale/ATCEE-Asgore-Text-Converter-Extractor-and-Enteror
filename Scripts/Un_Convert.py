import openpyxl
import re
convert_dic = ''

def import_from_converting_database(converting_database_directory):
    global convert_dic
    wd = openpyxl.load_workbook(converting_database_directory)
    database = wd.get_sheet_by_name("Database")
    convert_dic={'ً' : database['E2'].value,
            'ٌ' : database['E3'].value,
            'ٍ' : database['E4'].value,
            'َ' : database['E5'].value,
            'ُ' : database['E6'].value,
            'ِ' : database['E7'].value,
            'ّ' : database['E8'].value,
            'ْ' : database['E9'].value,
            'ﺀ' : database['E10'].value,
            'ﺁ' : database['E11'].value,
            'ﺂ' : database['D11'].value,
            'ﺃ' : database['E12'].value,
            'ﺄ' : database['D12'].value,
            "ﺅ" : database['E13'].value,
            "ﺆ" : database['D13'].value,
            "ﺇ" : database['E14'].value,
            "ﺈ" : database['D14'].value,
            "ﺉ" : database['E15'].value,
            "ﺊ" : database['D15'].value,
            "ﺋ" : database['B15'].value,
            "ﺌ" : database['C15'].value,
            "ﺍ" : database['E16'].value,
            "ﺎ" : database['D16'].value,
            "ﺏ" : database['E17'].value,
            "ﺐ" : database['D17'].value,
            "ﺑ" : database['B17'].value,
            "ﺒ" : database['C17'].value,
            "ﺓ" : database['E18'].value,
            "ﺔ" : database['D18'].value,
            "ﺕ" : database['E19'].value,
            "ﺖ" : database['D19'].value,
            "ﺗ" : database['B19'].value,
            "ﺘ" : database['C19'].value,
            "ﺙ" : database['E20'].value,
            "ﺚ" : database['D20'].value,
            "ﺛ" : database['B20'].value,
            "ﺜ" : database['C20'].value,
            "ﺝ" : database['E21'].value,
            "ﺞ" : database['D21'].value,
            "ﺟ" : database['B21'].value,
            "ﺠ" : database['C21'].value,
            "ﺡ" : database['E22'].value,
            "ﺢ" : database['D22'].value,
            "ﺣ" : database['B22'].value,
            "ﺤ" : database['C22'].value,
            "ﺥ" : database['E23'].value,
            "ﺦ" : database['D23'].value,
            "ﺧ" : database['B23'].value,
            "ﺨ" : database['C23'].value,
            "ﺩ" : database['E24'].value,
            "ﺪ" : database['D24'].value,
            "ﺫ" : database['E25'].value,
            "ﺬ" : database['D25'].value,
            "ﺭ" : database['E26'].value,
            "ﺮ" : database['D26'].value,
            "ﺯ" : database['E27'].value,
            "ﺰ" : database['D27'].value,
            "ﺱ" : database['E28'].value,
            "ﺲ" : database['D28'].value,
            "ﺳ" : database['B28'].value,
            "ﺴ" : database['C28'].value,
            "ﺵ" : database['E29'].value,
            "ﺶ" : database['D29'].value,
            "ﺷ" : database['B29'].value,
            "ﺸ" : database['C29'].value,
            "ﺹ" : database['E30'].value,
            "ﺺ" : database['D30'].value,
            "ﺻ" : database['B30'].value,
            "ﺼ" : database['C30'].value,
            "ﺽ" : database['E31'].value,
            "ﺾ" : database['D31'].value,
            "ﺿ" : database['B31'].value,
            "ﻀ" : database['C31'].value,
            "ﻁ" : database['E32'].value,
            "ﻂ" : database['D32'].value,
            "ﻃ" : database['B32'].value,
            "ﻄ" : database['C32'].value,
            "ﻅ" : database['E33'].value,
            "ﻆ" : database['D33'].value,
            "ﻇ" : database['B33'].value,
            "ﻈ" : database['C33'].value,
            "ﻉ" : database['E34'].value,
            "ﻊ" : database['D34'].value,
            "ﻋ" : database['B34'].value,
            "ﻌ" : database['C34'].value,
            "ﻍ" : database['E35'].value,
            "ﻎ" : database['D35'].value,
            "ﻏ" : database['B35'].value,
            "ﻐ" : database['C35'].value,
            "ﻑ" : database['E36'].value,
            "ﻒ" : database['D36'].value,
            "ﻓ" : database['B36'].value,
            "ﻔ" : database['C36'].value,
            "ﻕ" : database['E37'].value,
            "ﻖ" : database['D37'].value,
            "ﻗ" : database['B37'].value,
            "ﻘ" : database['C37'].value,
            "ﻙ" : database['E38'].value,
            "ﻚ" : database['D38'].value,
            "ﻛ" : database['B38'].value,
            "ﻜ" : database['C38'].value,
            "ﻝ" : database['E39'].value,
            "ﻞ" : database['D39'].value,
            "ﻟ" : database['B39'].value,
            "ﻠ" : database['C39'].value,
            "ﻡ" : database['E40'].value,
            "ﻢ" : database['D40'].value,
            "ﻣ" : database['B40'].value,
            "ﻤ" : database['C40'].value,
            "ﻥ" : database['E41'].value,
            "ﻦ" : database['D41'].value,
            "ﻧ" : database['B41'].value,
            "ﻨ" : database['C41'].value,
            "ﻩ" : database['E42'].value,
            "ﻪ" : database['D42'].value,
            "ﻫ" : database['B42'].value,
            "ﻬ" : database['C42'].value,
            "ﻭ" : database['E43'].value,
            "ﻮ" : database['D43'].value,
            "ﻯ" : database['E44'].value,
            "ﻰ" : database['D44'].value,
            "ﻱ" : database['E45'].value,
            "ﻲ" : database['D45'].value,
            "ﻳ" : database['B45'].value,
            "ﻴ" : database['C45'].value,
            "ﻵ" : database['E46'].value,
            "ﻶ" : database['D46'].value,
            "ﻷ" : database['E47'].value,
            "ﻸ" : database['D47'].value,
            "ﻹ" : database['E48'].value,
            "ﻺ" : database['D48'].value,
            "ﻻ" : database['E49'].value,
            "ﻼ" : database['D49'].value,
            "؟" : database['E50'].value,
            "،" : database['E51'].value,
            "؛" : database['E52'].value,
            }

def Convert(text, case='convert', start_command = '', end_command = ''):
    global convert_dic
    
    def Convert_text(text):
        if case == 'convert':
            for char in text:
                if char in convert_dic and convert_dic[char] != None and convert_dic[char] != '':
                    text = text.replace(char, convert_dic[char])
        else:
            for k, v in convert_dic.items():
                if v != None and v != '':
                    text = text.replace(v, k)
        return text
    
    if start_command != '' and end_command != '':
        commands_chars = '.[]{}*+?()^'
        re_start_command = start_command
        re_end_command = end_command
        for char in commands_chars:
            re_start_command = re_start_command.replace(char, '\\'+char)
            re_end_command = re_end_command.replace(char, '\\'+char)
        pattern = re_start_command + "(.*?)" + re_end_command
        text_list = re.split(pattern, text)
        
        for _ in range(len(text_list)):
            if _%2 == 1:
                text_list[_] = start_command + text_list[_] + end_command
            else:
                if text_list[_] != '':
                    text_list[_] = Convert_text(text_list[_])
        text = ''.join(text_list)
    else:
        text = Convert_text(text)

    return text