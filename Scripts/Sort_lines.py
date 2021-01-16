def script(text, case = 'short to long'):
    lines_list = text.split('\n')
    lines_list.sort(key=len)
    
    if case == 'long to short': lines_list = lines_list[::-1]
    text = '\n'.join(lines_list)
    
    return text