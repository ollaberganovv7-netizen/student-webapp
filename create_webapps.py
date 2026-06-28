import codecs

# Read mustaqil.html
with codecs.open('mustaqil.html', 'r', 'utf-8') as f:
    content = f.read()

# Create diplom_settings.html
diplom_content = content.replace('Mustaqil Ish Tahrirlash', 'Diplom ishi Sozlamalari')
diplom_content = diplom_content.replace('Mavzu nomi', 'Diplom ishi')
with codecs.open('diplom_settings.html', 'w', 'utf-8') as f:
    f.write(diplom_content)

# Create quiz_settings.html
quiz_content = content.replace('Mustaqil Ish Tahrirlash', 'Quiz Sozlamalari')
quiz_content = quiz_content.replace('Mavzu nomi', 'Avtomatik quiz tuzish')
with codecs.open('quiz_settings.html', 'w', 'utf-8') as f:
    f.write(quiz_content)

print("Created diplom_settings.html and quiz_settings.html")
