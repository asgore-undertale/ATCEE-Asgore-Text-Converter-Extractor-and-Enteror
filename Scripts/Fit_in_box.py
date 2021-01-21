def script(text, textzone_width, lines_num, database_directory = '', new_line_com = '\n', new_page_com = '\n\n'):
    import openpyxl
    wd = openpyxl.load_workbook(database_directory)
    database = wd.get_sheet_by_name("Database")
    
    chars_dic = {'ً' : database['E2'].value,
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
            "." : database['E53'].value,
            " " : database['E54'].value,
        }
    
    new_text = ""
    x, y = 0, 1
    
    if new_line_com != '': text = text.replace(new_line_com, u'\uffff')#so please do not use u'\uffff' in your text
    if new_page_com != '': text = text.replace(new_page_com, u'\ufffe')#u'\ufffe' too
    
    if new_line_com != '\n' and new_page_com != '\n': text = text.replace('\n', ' ')
    text_list = text.split(" ")

    for i in range(len(text_list)):
        if i < len(text_list) -1:
            text_list[i] += " "
        item_width = 0
        
        for char in text_list[i]:
            if char in chars_dic:
                item_width += chars_dic[char]
            else:
                if char != u'\uffff' and char != u'\ufffe':
                    print('"'+char+'" in not in dictionary.')
        
        if item_width > textzone_width:
            for char in text_list[i]:
                if char == u'\uffff':
                    if y < lines_num:
                        y += 1
                        new_text += u'\uffff'
                    else:
                        y = 1
                        new_text += u'\ufffe'
                elif char == u'\ufffe':
                    y = 1
                    new_text += u'\ufffe'
                else:
                    if char in chars_dic:
                        char_width = chars_dic[char]
                        if char_width > textzone_width:
                            print(char + 'is wider than text zone')
                            continue
                    else:
                        char_width = 0
                    if x + char_width > textzone_width:
                        if y < lines_num:
                            new_text += new_line_com + char
                            y += 1
                        else:
                            new_text += new_page_com + char
                            y = 1
                        x = char_width
                    else:
                        new_text += char
                        x += char_width
        else:
            if x + item_width > textzone_width:
                if y < lines_num:
                    new_text += new_line_com# + text_list[i]
                    y += 1
                else:
                    new_text += new_page_com# + text_list[i]
                    y = 1
                x = 0#item_width
                for char in text_list[i]:#
                    if char == u'\uffff':
                        if y < lines_num:
                            y += 1
                            new_text += u'\uffff'
                        else:
                            y = 1
                            new_text += u'\ufffe'
                    elif char == u'\ufffe':
                        y = 1
                        new_text += u'\ufffe'
                    else:
                        if char in chars_dic:
                            char_width = chars_dic[char]
                            if char_width > textzone_width:
                                print(char + 'is wider than text zone')
                                continue
                        else:
                            char_width = 0
                        new_text += char
                        x += char_width
            else:
              #  new_text += text_list[i]
               # x += item_width
                for char in text_list[i]:#
                    if char == u'\uffff':
                        if y < lines_num:
                            y += 1
                            new_text += u'\uffff'
                        else:
                            y = 1
                            new_text += u'\ufffe'
                    elif char == u'\ufffe':
                        y = 1
                        new_text += u'\ufffe'
                    else:
                        if char in chars_dic:
                            char_width = chars_dic[char]
                            if char_width > textzone_width:
                                print(char + 'is wider than text zone')
                                continue
                        else:
                            char_width = 0
                        new_text += char
                        x += char_width

    if new_page_com != '': 
        new_text = new_text.replace(u'\ufffe', new_page_com)
        new_text = new_text.replace(" "+new_page_com, new_page_com).replace(new_page_com+" ", new_page_com)
    else:
        if new_line_com != '': 
            new_text = new_text.replace(u'\ufffe', new_line_com)
        
    if new_line_com != '': 
        new_text = new_text.replace(u'\uffff', new_line_com)
        new_text = new_text.replace(" "+new_line_com, new_line_com).replace(new_line_com+" ", new_line_com)
    return new_text