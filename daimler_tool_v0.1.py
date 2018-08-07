from docx import Document
import shutil
import os

files = os.listdir(os.curdir)

for f in files:
    if f.endswith('.docx'):

        file_name = f

        document = Document(file_name)
        tables = document.tables

        unique_key = tables[0].rows[31].cells[2].text
        unique_key_file_name = unique_key + '.s19'

        daimler_part_number = file_name.split("_")[1]
        file_name_sec = "-".join(unique_key.split("-")[3:9])

        folder_name = file_name_sec + '-' + daimler_part_number

        if not os.path.exists(folder_name):
            os.makedirs(folder_name)

        shutil.move(file_name, folder_name)
        shutil.move(unique_key_file_name, folder_name)










