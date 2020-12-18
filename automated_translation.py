from time import sleep
from pywinauto.application import Application

import os

program_path = r'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.exe'
lectures_dir = os.getcwd() + r'\\lectures\\'

for file in os.listdir(lectures_dir):
    file_name = file
    file_path = lectures_dir + file_name
    complete_path = r'{} "{}"'.format(program_path, file_path)

    language = 'hindi'

    app = Application(backend='uia').start(complete_path)

    main = app.window(title_re=f'{file_name} - Word')
    main.Review_tab.select()
    main.Translate.click_input()
    main.Translate_Document.click_input()

    while(not app.Word.exists()):
        sleep(3)

    app.Word.Save.click()
    app.Word['Save this fileDialog'].Edit.click_input()
    app.Word['Save this fileDialog'].Edit.set_text(f'{file_name} {language}')
    app.Word['Save this fileDialog'].Save.click_input()

    secondary = app.window(title_re=f'{file_name} {language}.docx - Word')
    secondary.Close.click()

    main.Close.click()
