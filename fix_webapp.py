import os
import glob

def fix_file(filepath):
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    
    original = content
    # Common broken utf-8 in win-1251
    replacements = {
        'вќЊ': '❌',
        'вЌЊ': '❌',
        'вњ…': '✅',
        'вљ': '⚠️',
        'в”Ђ': '─',
        '': ''
    }
    for bad, good in replacements.items():
        content = content.replace(bad, good)
        
    if content != original:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Fixed {filepath}")

for ext in ['*.html', '*.js', '*.css', 'js/*.js', 'css/*.css']:
    for filepath in glob.glob(ext):
        fix_file(filepath)
