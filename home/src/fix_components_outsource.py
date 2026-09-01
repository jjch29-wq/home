path = r'c:\Users\-\PMI\home\src\views\components.py'
with open(path, encoding='utf-8') as f:
    content = f.read()

# Find the line with broken encoding and replace it
lines = content.splitlines(keepends=True)
new_lines = []
replaced = False
for line in lines:
    # Find the else branch outsource_defaults single-line (broken or original)
    if 'outsource_defaults = [' in line and '\uCF00\uC774\uC5D4\uB514\uC774' in line and '\uACE0\uB824\uAC80\uC0AC' in line:
        # already has 고려검사 - rebuild properly
        indent = '            '
        new_lines.append(f'{indent}outsource_defaults = [\n')
        new_lines.append(f'{indent}    ("\ucf00\uc774\uc5d4\ub514\uc774",     "\ubc29\uc0ac\uc120\ud22c\uacfc\uac80\uc0ac", 0, 15000),\n')
        new_lines.append(f'{indent}    ("\uace0\ub824\uac80\uc0ac",       "\ubc29\uc0ac\uc120\ud22c\uacfc\uac80\uc0ac", 0, 13000),\n')
        new_lines.append(f'{indent}    ("\ud55c\uad6d\uae30\uacc4\uac80\uc0ac\uc18c", "\ubc29\uc0ac\uc120\ud22c\uacfc\uac80\uc0ac", 0, 15000),\n')
        new_lines.append(f'{indent}]\n')
        replaced = True
        print(f"Replaced broken line: {line[:80]}")
        continue
    elif 'outsource_defaults = [("케이엔디이"' in line:
        # Original single line - replace with multi-line
        indent = '            '
        new_lines.append(f'{indent}outsource_defaults = [\n')
        new_lines.append(f'{indent}    ("\ucf00\uc774\uc5d4\ub514\uc774",     "\ubc29\uc0ac\uc120\ud22c\uacfc\uac80\uc0ac", 0, 15000),\n')
        new_lines.append(f'{indent}    ("\uace0\ub824\uac80\uc0ac",       "\ubc29\uc0ac\uc120\ud22c\uacfc\uac80\uc0ac", 0, 13000),\n')
        new_lines.append(f'{indent}    ("\ud55c\uad6d\uae30\uacc4\uac80\uc0ac\uc18c", "\ubc29\uc0ac\uc120\ud22c\uacfc\uac80\uc0ac", 0, 15000),\n')
        new_lines.append(f'{indent}]\n')
        replaced = True
        print(f"Replaced original line: {line[:80]}")
        continue
    new_lines.append(line)

with open(path, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

if replaced:
    print("OK: components.py updated successfully")
else:
    # debug
    for i, line in enumerate(lines, 1):
        if 'outsource_defaults' in line:
            print(f"Line {i}: {repr(line[:120])}")
    print("NOT found - check above")
