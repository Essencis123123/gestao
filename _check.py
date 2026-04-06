content = open('main.py', encoding='utf-8').read()

# Find generate_pagamentos_html or load_pagamentos end behavior
# Look for 'setHtml' usages around load_pagamentos
idx = content.index('def load_pagamentos')
pagamentos_section = content[idx:idx+3000]
# Find setHtml calls in this section
import re
matches = [(m.start(), content[idx+m.start()-200:idx+m.start()+50]) for m in re.finditer(r'setHtml', pagamentos_section)]
for pos, ctx in matches[:3]:
    print(f"--- setHtml at offset {pos} ---")
    print(ctx)
    print()

# Also find general html = get_base_html() + append pattern used in pagamentos
# Find generate_new_pagamento_html or the payment HTML function
fns = [m.start() for m in re.finditer(r'def generate_\w+html', content)]
print("HTML generator functions:")
for f in fns:
    print(f"  line {content[:f].count(chr(10))+1}: {content[f:f+80]}")
