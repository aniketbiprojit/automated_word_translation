from time import sleep
from pywinauto.application import Application

import os

program_path = r'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.exe'
lectures_dir = os.getcwd() + r'\\lectures\\'

for file in os.listdir(lectures_dir):
    language = 'bangla'
    if f'{file} {language}.docx' in os.listdir(os.getcwd() + r'\\translations\\') or file == 'lec4.doc':
        print('continue ', file)
        continue
    try:
        # file='lec10.doc'
        file_name = file
        file_path = lectures_dir + file_name
        complete_path = r'{} "{}"'.format(program_path, file_path)
        
        app = Application(backend='uia').start(complete_path)
        sleep(1)
        file_no_ext = file_name.split('.')[0]
        app.connect(title_re=rf'{file_no_ext}.*')
        sleep(1)
        print(app.windows())
        main = app.window(title_re=rf'{file_no_ext}.*')
        main.Review_tab.select()
        sleep(3)
        main.Translate.click_input()
        sleep(3)
        main.Translate_Document.click_input()

        while(not (app['Compatibility Mode - Word'].exists() or app['Word'])):
            sleep(3)

        sleep(10)
        if app['Compatibility Mode - Word'].exists():
            main.Close.click()
            app['Compatibility Mode - Word'].Close.click()
            sleep(3)
            app['Compatibility Mode - Word']['Save this fileDialog'].Edit.click_input()
            app['Compatibility Mode - Word']['Save this fileDialog'].Edit.set_text(
                f'{file_name} {language}')
            app['Compatibility Mode - Word']['Save this fileDialog'].Save.click_input()
        
        else:
            main.Close.click()
            app.Word.Save.click()
            app.Word['Save this fileDialog'].Edit.click_input()
            app.Word['Save this fileDialog'].Edit.set_text(f'{file_name} {language}')
            app.Word['Save this fileDialog'].Save.click_input()
            secondary = app.window(title_re=f'{file_name} {language}.docx - Word')
            secondary.Close.click()
        sleep(5)
    except Exception as e:
        print(e)
        with open('errs.txt','a') as f:
            try:
                f.write(str(e))
            except:
                pass
        print(file)
    

