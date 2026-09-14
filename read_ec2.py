with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()
in_ec = False
count = 0
for i, l in enumerate(lines):
    if 'elif menu=="Estado de Cuenta":' in l:
        in_ec = True
    if in_ec:
        if count >= 200 and count < 400:
            print(f"{i}: {l}", end="")
        count += 1
