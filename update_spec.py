with open('qlvt.spec', 'r', encoding='utf-8') as f:
    content = f.read()

old_imports = """hiddenimports=[
        'win32com.client',
        'pythoncom',
        'PyQt5.sip'
    ],"""

new_imports = """hiddenimports=[
        'win32com.client',
        'pythoncom',
        'PyQt5.sip',
        'pandas',
        'excel_processor',
        'openpyxl',
        'numpy'
    ],"""

content = content.replace(old_imports, new_imports)

with open('qlvt.spec', 'w', encoding='utf-8') as f:
    f.write(content)

print("Updated qlvt.spec successfully!")
