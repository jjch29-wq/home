import re

file_path = 'src/Material-Master-Manager-V13.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Update toggle_sash_lock to include sticky_bottom_panel height
content = content.replace(
    'content_h = bbox[1] + bbox[3] + 40 # Add some padding',
    'content_h = (bbox[1] + bbox[3]) + (self.sticky_bottom_panel.winfo_height() if hasattr(self, "sticky_bottom_panel") else 0) + 40'
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print('toggle_sash_lock updated for sticky bottom panel.')
