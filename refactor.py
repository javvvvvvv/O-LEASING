import re
with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Strip out the old CSS
new_content = re.sub(
    r'st\.markdown\("""\n<style>\n.*?</style>\n""", unsafe_allow_html=True\)', 
    'from ui.theme import inyectar_css\ninyectar_css()', 
    content, 
    flags=re.DOTALL
)

# Delete precargar_ejemplo
new_content = re.sub(
    r'def precargar_ejemplo\(\):.*?guardar\(c\);\s*st\.success\("Contrato de ejemplo precargado\."\)', 
    '', 
    new_content, 
    flags=re.DOTALL
)
new_content = new_content.replace('precargar_ejemplo()\n', '')

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(new_content)
