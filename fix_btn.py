with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()
content = content.replace('kind="primary"', 'type="primary"')
with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
