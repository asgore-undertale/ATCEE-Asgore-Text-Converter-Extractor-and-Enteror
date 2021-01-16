def script(text, case='convert', database_directory=''):
    import openpyxl
    wd = openpyxl.load_workbook(database_directory)
    database = wd.get_sheet_by_name("Database")
    
    un_convert={u"\ufe70" : database['E2'].value , #ً
                u"\ufe72" : database['E3'].value , #ٌ
                u"\ufe74" : database['E4'].value , #ٍ
                u"\ufe76" : database['E5'].value , #َ
                u"\ufe78" : database['E6'].value , #ُ
                u"\ufe7a" : database['E7'].value , #ِ
                u"\ufe7c" : database['E8'].value , #ّ
                u"\ufe7e" : database['E9'].value , #ْ
                u"\ufe80" : database['E10'].value, #ﺀ
                u"\ufe81" : database['E11'].value, #ﺁ
                u"\ufe82" : database['D10'].value, #ﺂ
                u"\ufe83" : database['E12'].value, #ﺃ
                u"\ufe84" : database['D12'].value, #ﺄ
                u"\ufe85" : database['E13'].value, #ﺅ
                u"\ufe86" : database['D13'].value, #ﺆ
                u"\ufe87" : database['E14'].value, #ﺇ
                u"\ufe88" : database['D14'].value, #ﺈ
                u"\ufe89" : database['E15'].value, #ﺉ
                u"\ufe8a" : database['D15'].value, #ﺊ
                u"\ufe8b" : database['B15'].value, #ﺋ
                u"\ufe8c" : database['C15'].value, #ﺌ
                u"\ufe8d" : database['E16'].value, #ﺍ
                u"\ufe8e" : database['D16'].value, #ﺎ
                u"\ufe8f" : database['E17'].value, #ﺏ
                u"\ufe90" : database['D17'].value, #ﺐ
                u"\ufe91" : database['B17'].value, #ﺑ
                u"\ufe92" : database['C17'].value, #ﺒ
                u"\ufe93" : database['E18'].value, #ﺓ
                u"\ufe94" : database['D18'].value, #ﺔ
                u"\ufe95" : database['E19'].value, #ﺕ
                u"\ufe96" : database['D19'].value, #ﺖ
                u"\ufe97" : database['B19'].value, #ﺗ
                u"\ufe98" : database['C19'].value, #ﺘ
                u"\ufe99" : database['E20'].value, #ﺙ
                u"\ufe9a" : database['D20'].value, #ﺚ
                u"\ufe9b" : database['B20'].value, #ﺛ
                u"\ufe9c" : database['C20'].value, #ﺜ
                u"\ufe9d" : database['E21'].value, #ﺝ
                u"\ufe9e" : database['D21'].value, #ﺞ
                u"\ufe9f" : database['B21'].value, #ﺟ
                u"\ufea0" : database['C21'].value, #ﺠ
                u"\ufea1" : database['E22'].value, #ﺡ
                u"\ufea2" : database['D22'].value, #ﺢ
                u"\ufea3" : database['B22'].value, #ﺣ
                u"\ufea4" : database['C22'].value, #ﺤ
                u"\ufea5" : database['E23'].value, #ﺥ
                u"\ufea6" : database['D23'].value, #ﺦ
                u"\ufea7" : database['B23'].value, #ﺧ
                u"\ufea8" : database['C23'].value, #ﺨ
                u"\ufea9" : database['E24'].value, #ﺩ
                u"\ufeaa" : database['D24'].value, #ﺪ
                u"\ufeab" : database['E25'].value, #ﺫ
                u"\ufeac" : database['D25'].value, #ﺬ
                u"\ufead" : database['E26'].value, #ﺭ
                u"\ufeae" : database['D26'].value, #ﺮ
                u"\ufeaf" : database['E27'].value, #ﺯ
                u"\ufeb0" : database['D27'].value, #ﺰ
                u"\ufeb1" : database['E28'].value, #ﺱ
                u"\ufeb2" : database['D28'].value, #ﺲ
                u"\ufeb3" : database['B28'].value, #ﺳ
                u"\ufeb4" : database['C28'].value, #ﺴ
                u"\ufeb5" : database['E29'].value, #ﺵ
                u"\ufeb6" : database['D29'].value, #ﺶ
                u"\ufeb7" : database['B29'].value, #ﺷ
                u"\ufeb8" : database['C29'].value, #ﺸ
                u"\ufeb9" : database['E30'].value, #ﺹ
                u"\ufeba" : database['D30'].value, #ﺺ
                u"\ufebb" : database['B30'].value, #ﺻ
                u"\ufebc" : database['C30'].value, #ﺼ
                u"\ufebd" : database['E31'].value, #ﺽ
                u"\ufebe" : database['D31'].value, #ﺾ
                u"\ufebf" : database['B31'].value, #ﺿ
                u"\ufec0" : database['C31'].value, #ﻀ
                u"\ufec1" : database['E32'].value, #ﻁ
                u"\ufec2" : database['D32'].value, #ﻂ
                u"\ufec3" : database['B32'].value, #ﻃ
                u"\ufec4" : database['C32'].value, #ﻄ
                u"\ufec5" : database['E33'].value, #ﻅ
                u"\ufec6" : database['D33'].value, #ﻆ
                u"\ufec7" : database['B33'].value, #ﻇ
                u"\ufec8" : database['C33'].value, #ﻈ
                u"\ufec9" : database['E34'].value, #ﻉ
                u"\ufeca" : database['D34'].value, #ﻊ
                u"\ufecb" : database['B34'].value, #ﻋ
                u"\ufecc" : database['C34'].value, #ﻌ
                u"\ufecd" : database['E35'].value, #ﻍ
                u"\ufece" : database['D35'].value, #ﻎ
                u"\ufecf" : database['B35'].value, #ﻏ
                u"\ufed0" : database['C35'].value, #ﻐ
                u"\ufed1" : database['E36'].value, #ﻑ
                u"\ufed2" : database['D36'].value, #ﻒ
                u"\ufed3" : database['B36'].value, #ﻓ
                u"\ufed4" : database['C36'].value, #ﻔ
                u"\ufed5" : database['E37'].value, #ﻕ
                u"\ufed6" : database['D37'].value, #ﻖ
                u"\ufed7" : database['B37'].value, #ﻗ
                u"\ufed8" : database['C37'].value, #ﻘ
                u"\ufed9" : database['E38'].value, #ﻙ
                u"\ufeda" : database['D38'].value, #ﻚ
                u"\ufedb" : database['B38'].value, #ﻛ
                u"\ufedc" : database['C38'].value, #ﻜ
                u"\ufedd" : database['E39'].value, #ﻝ
                u"\ufede" : database['D39'].value, #ﻞ
                u"\ufedf" : database['B39'].value, #ﻟ
                u"\ufee0" : database['C39'].value, #ﻠ
                u"\ufee1" : database['E40'].value, #ﻡ
                u"\ufee2" : database['D40'].value, #ﻢ
                u"\ufee3" : database['B40'].value, #ﻣ
                u"\ufee4" : database['C40'].value, #ﻤ
                u"\ufee5" : database['E41'].value, #ﻥ
                u"\ufee6" : database['D41'].value, #ﻦ
                u"\ufee7" : database['B41'].value, #ﻧ
                u"\ufee8" : database['C41'].value, #ﻨ
                u"\ufee9" : database['E42'].value, #ﻩ
                u"\ufeea" : database['D42'].value, #ﻪ
                u"\ufeeb" : database['B42'].value, #ﻫ
                u"\ufeec" : database['C42'].value, #ﻬ
                u"\ufeed" : database['E43'].value, #ﻭ
                u"\ufeee" : database['D43'].value, #ﻮ
                u"\ufeef" : database['E44'].value, #ﻯ
                u"\ufef0" : database['D44'].value, #ﻰ
                u"\ufef1" : database['E45'].value, #ﻱ
                u"\ufef2" : database['D45'].value, #ﻲ
                u"\ufef3" : database['B45'].value, #ﻳ
                u"\ufef4" : database['C45'].value, #ﻴ
                u"\ufef5" : database['E46'].value, #ﻵ
                u"\ufef6" : database['D46'].value, #ﻶ
                u"\ufef7" : database['E47'].value, #ﻷ
                u"\ufef8" : database['D47'].value, #ﻸ
                u"\ufef9" : database['E48'].value, #ﻹ
                u"\ufefa" : database['D48'].value, #ﻺ
                u"\ufefb" : database['E49'].value, #ﻻ
                u"\ufefc" : database['D49'].value, #ﻼ
                u"\u061f" : database['E50'].value, #؟
                u"\u060c" : database['E51'].value, #،
                u"\u061b" : database['E52'].value, #؛
    }
    
    if case == 'convert':
        for char in text:
            if char in un_convert:
                if un_convert[char] != None and un_convert[char] != '':
                    text = text.replace(char, un_convert[char])
    else:
        for k, v in un_convert.items():
            if v != None and v != '':
                text = text.replace(v, k)

    return text