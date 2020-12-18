from bs4 import BeautifulSoup
import os

from os.path import basename, splitext

folder_path = os.path.join(os.getcwd(),'converted')
for filename in os.listdir(folder_path):
    if(filename.endswith('.htm')):
        filepath = os.path.join(folder_path,filename)

        with open(filepath) as f:
            html_file = f.read()

        soup = BeautifulSoup(html_file, features='html.parser')

        complete_url = 'https://translation.aicte-india.org/v1/cdn/'

        for img in soup.findAll('img'):
            img['src'] = complete_url + splitext(basename(img['src']))[0]



        with open(filepath,'w') as f:
            f.write(str(soup))
