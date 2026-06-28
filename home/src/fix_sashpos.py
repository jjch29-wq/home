import re

with open('Material-Master-Manager-V14_20260627.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Replace assignments: sash_pos = self.daily_usage_paned.sashpos(0) -> sash_pos = 500
code = re.sub(r'(\w+)\s*=\s*self\.daily_usage_paned\.sashpos\(0\)', r'\1 = 500', code)

# Replace method calls: self.daily_usage_paned.sashpos(0, target_pos) -> pass
# Since it might be part of a try-except or just a statement, let's wrap it in an if
# We can just replace 'self.daily_usage_paned.sashpos' with 'getattr(self.daily_usage_paned, "sashpos", lambda *a, **k: 500)'
# Actually, the safest way is to regex replace 'self.daily_usage_paned.sashpos' with 'getattr(self.daily_usage_paned, "sashpos", lambda *a: 500)'
code = code.replace('self.daily_usage_paned.sashpos', 'getattr(self.daily_usage_paned, "sashpos", lambda *args: 500)')

with open('Material-Master-Manager-V14_20260627.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("Sashpos fixed!")
