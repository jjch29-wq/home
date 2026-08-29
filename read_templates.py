with open('c:\\Users\\jjch2\\Desktop\\PMI\\templates_list.txt', 'r', encoding='utf-16le') as f:
    text = f.read()
    with open('c:\\Users\\jjch2\\Desktop\\PMI\\templates_list_utf8.txt', 'w', encoding='utf-8') as f2:
        f2.write(text.replace('\ufeff', ''))
