import os

for fname in ['flask_app.py', 'app.py']:
    path = os.path.join(os.path.dirname(__file__), fname)
    if not os.path.exists(path):
        continue
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Minified versions (no space after comma)
    content = content.replace("'IBM Plex Sans',sans-serif", "Arial,sans-serif")
    content = content.replace("'Playfair Display',serif", "Arial,sans-serif")
    content = content.replace("'IBM Plex Mono',monospace", "Arial,sans-serif")
    
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f'{fname}: done')

# Remove google fonts links
for fname in ['flask_app.py', 'app.py']:
    path = os.path.join(os.path.dirname(__file__), fname)
    if not os.path.exists(path):
        continue
    with open(path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    new_lines = [l for l in lines if 'fonts.googleapis.com' not in l]
    removed = len(lines) - len(new_lines)
    
    with open(path, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
    
    print(f'{fname}: removed {removed} Google Fonts links')

print('All done!')
