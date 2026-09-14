with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

old = 'defs = [h for h in opciones if "JULIO 2026" in h.upper() or "AGOSTO 2026" in h.upper()]'
new = 'defs = [h for h in opciones if "JULIO 2026" in h.upper() or "AGOSTO 2026" in h.upper() or h.upper().strip() == "CARTERA"]'
content = content.replace(old, new)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
