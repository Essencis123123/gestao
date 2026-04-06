import os

for fname in ['flask_app.py', 'app.py']:
    path = os.path.join(os.path.dirname(__file__), fname)
    if not os.path.exists(path):
        print(f'{fname} not found')
        continue
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    old_count = content.count('Arial')
    content = content.replace("'IBM Plex Sans', sans-serif", "Arial, sans-serif")
    content = content.replace("'Playfair Display', serif", "Arial, sans-serif")
    content = content.replace("'IBM Plex Mono', monospace", "Arial, sans-serif")
    new_count = content.count('Arial')
    
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f'{fname}: {new_count - old_count} replacements made')

print('Done!')
