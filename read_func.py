with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()
in_func = False
for i, l in enumerate(lines):
    if l.startswith('def tabla_mensual_conceptos'):
        in_func = True
    elif in_func and l.startswith('def '):
        break
    if in_func:
        print(l, end="")
