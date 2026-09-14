with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()
with open('current_ec.py', 'w', encoding='utf-8') as out:
    in_ec = False
    for l in lines:
        if 'elif menu=="Estado de Cuenta":' in l:
            in_ec = True
        elif in_ec and l.strip().startswith('elif menu=="'):
            break
        if in_ec:
            out.write(l)
