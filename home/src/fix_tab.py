import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    content = f.read()

parts = content.split('if __name__ == "__main__":')
if len(parts) == 2:
    class_code = parts[0]
    main_code = 'if __name__ == "__main__":' + parts[1]
    
    main_parts = main_code.split('    def load_history(self):')
    
    if len(main_parts) == 2:
        clean_main = main_parts[0]
        methods_code = '    def load_history(self):' + main_parts[1]
        
        new_content = class_code + methods_code + '\n' + clean_main
        
        with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
            f.write(new_content)
        print('Fixed successfully.')
    else:
        print('Could not find load_history in main block')
else:
    print('Could not find __main__ block')
