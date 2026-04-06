import re

with open('temp_check.js', 'r', encoding='utf-8') as f:
    js = f.read()

print('JS length:', len(js))

# NaN check
nans = [m for m in re.finditer(r'NaN', js)]
print('NaN count:', len(nans))
for m in nans[:3]:
    start = max(0, m.start()-40)
    print('  Context:', repr(js[start:m.end()+20]))

# Infinity check
infs = [m for m in re.finditer(r'Infinity', js)]
print('Infinity count:', len(infs))
for m in infs[:3]:
    start = max(0, m.start()-40)
    print('  Context:', repr(js[start:m.end()+20]))

# except check
exc = re.search(r'\bexcept\b', js)
print('except keyword found:', bool(exc))

# Key functions
for fn in ['openPanelEstoque', 'renderEstInline', 'switchEstTab', 'filterEstoque', 'EST_DATA']:
    print(f'{fn} found:', fn in js)
