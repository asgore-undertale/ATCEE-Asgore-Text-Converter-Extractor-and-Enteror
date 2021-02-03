import re

def fit_in_box(text, chars_width_dic, textzone_width, lines_num, new_line_com = '\n', new_page_com = '\n\n', start_command = '[', end_command = ']'):
    new_text = ""
    x, y = 0, 1
    
    def fit(text, x, y, new_text):
        item_width = 0
        for char in text:
            if char in chars_width_dic:
                item_width += chars_width_dic[char]
            else:
                if char != u'\uffff' and char != u'\ufffe':
                    print('"'+char+'" in not in dictionary.')
        
        if item_width > textzone_width:
            for char in text:
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
                    if char in chars_width_dic:
                        char_width = chars_width_dic[char]
                        if char_width > textzone_width:
                            print(char + ' is wider than text zone')
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
                    new_text += new_line_com
                    y += 1
                else:
                    new_text += new_page_com
                    y = 1
                x = 0
                for char in text:
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
                        if char in chars_width_dic:
                            char_width = chars_width_dic[char]
                            if char_width > textzone_width:
                                print(char + 'is wider than text zone')
                                continue
                        else:
                            char_width = 0
                        new_text += char
                        x += char_width
            else:
                for char in text:
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
                        if char in chars_width_dic:
                            char_width = chars_width_dic[char]
                            if char_width > textzone_width:
                                print(char + 'is wider than text zone')
                                continue
                        else:
                            char_width = 0
                        new_text += char
                        x += char_width
        return x, y, new_text
    
    if new_line_com != '': text = text.replace(new_line_com, u'\uffff')#so please do not use u'\uffff' in your text
    if new_page_com != '': text = text.replace(new_page_com, u'\ufffe')#u'\ufffe' too
    
    if new_line_com != '\n' and new_page_com != '\n' and start_command != '\n' and end_command != '\n':
        text = text.replace('\n', ' ')
    
    if start_command != '' and end_command != '':
        commands_chars = '.[]{}*+?()^'
        re_start_command = start_command
        re_end_command = end_command
        for char in commands_chars:
            re_start_command = re_start_command.replace(char, '\\'+char)
            re_end_command = re_end_command.replace(char, '\\'+char)
        pattern = re_start_command + "(.*?)" + re_end_command
        text_list = re.split(pattern, text)
    else:
        text_list = [text]
    
    for _ in range(len(text_list)):
        if _ < len(text_list): new_text += ' '
        if _%2 == 0:
            min_text_list = text_list[_].split(" ")
            
            for i in range(len(min_text_list)):
                x, y, new_text = fit(min_text_list[i], x, y, new_text)
        else:
            new_text +=  start_command + text_list[_] + end_command

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