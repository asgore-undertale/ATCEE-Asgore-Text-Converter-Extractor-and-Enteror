def script(text, case='convert', convert_dic='', start_command = '', end_command = ''):
    def Convert(text):
        if case == 'convert':
            for char in text:
                if char in convert_dic:
                    if convert_dic[char] != None and convert_dic[char] != '':
                        text = text.replace(char, convert_dic[char])
        else:
            for k, v in convert_dic.items():
                if v != None and v != '':
                    text = text.replace(v, k)
        return text
        
    if start_command != '' and end_command != '':
        text = text.replace(start_command, end_command)
        text_list = text.split(end_command)
        for _ in range(len(text_list)):
            if _%2 == 1:
                text_list[_] = start_command + text_list[_] + end_command
            else:
                if text_list[_] != '':
                    text_list[_] = Convert(text_list[_])
        text = ''.join(text_list)
    else:
        text = Convert(text)

    return text