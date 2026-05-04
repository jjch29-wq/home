import chardet

with open(r'Na-aba\Material-Master-Manager-V14.py', 'rb') as f:
    rawdata = f.read()
    result = chardet.detect(rawdata)
    print(result)
