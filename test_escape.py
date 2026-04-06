#!/usr/bin/env python3
"""Quick test for escaping in the f-string template"""

# Simulate what the f-string does
test = f"""rows += '<button onclick="removerFornecedor(\\'' + id + '\\')" >';"""

with open(r'C:\Users\2700024\Desktop\test_escape_result.txt', 'w', encoding='utf-8') as fout:
    fout.write("Generated JS:\n")
    fout.write(test + "\n\n")
    if "\\'" in test:
        fout.write("OK: \\' is present in output - JS will parse correctly\n")
    else:
        fout.write("BUG: \\' is NOT in output\n")
    if "removerFornecedor(''" in test:
        fout.write("BUG: Double empty quotes found - will cause SyntaxError!\n")
    fout.write("repr: " + repr(test) + "\n")
