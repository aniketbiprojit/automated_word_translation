import os

folder_path = os.path.join(os.getcwd(),'translations')

for file in os.listdir(folder_path):
    complete_path = os.path.join(folder_path,file)
    print(complete_path)
    os.system(f"soffice --headless --convert-to htm:HTML --outdir converted \"{complete_path}\"")
