from docx import Document
import shutil
import os
import datetime

files = os.listdir(os.curdir)
error_count = 0
key_position = 0


for f in files:

    if f.endswith('.docx'):
        file_name = f

        document = Document(file_name)
        tables = document.tables

        if not tables[0].rows[31].cells[2].text:
            error_count+=1

            if not tables[0].rows[32].cells[2].text:

                if not tables[0].rows[33].cells[2].text:

                    if not tables[0].rows[34].cells[2].text:
                        key_position = 0
                        with open("Error-Logs.txt", "a") as text_file:
                            text_file.write(str(error_count) + "-  " + file_name + " file has no code.\n")
                            text_file.write(
                                "------------------------------------------------------------------------------------------------ \n")
                    else:
                        unique_key=tables[0].rows[34].cells[2].text
                        key_position = 4
                else:
                    unique_key = tables[0].rows[33].cells[2].text
                    key_position = 3
            else:
                unique_key = tables[0].rows[32].cells[2].text
                key_position = 2

            if(key_position!=0):
                with open("Error-Logs.txt", "a") as text_file:
                    text_file.write(
                        str(error_count) + "- " + file_name + " 's code is in " + str(key_position) + ". line.\n")
                    text_file.write(
                        "------------------------------------------------------------------------------------------------ \n")


        else:
            unique_key = tables[0].rows[31].cells[2].text
            key_position=1

        if(key_position!=0):

            unique_key_file_name = unique_key + '.s19'

            daimler_part_number = file_name.split("_")[1]
            file_name_sec = "-".join(unique_key.split("-")[3:9])

            folder_name = file_name_sec + '-' + daimler_part_number

            if os.path.exists(unique_key_file_name):
                if not os.path.exists(folder_name):
                    os.makedirs(folder_name)

                shutil.move(unique_key_file_name, folder_name)
                shutil.move(file_name, folder_name)

            else:
                error_count += 1
                if not os.path.exists("Error-Logs.txt"):
                    with open("Error-Logs.txt", "w") as text_file:
                        text_file.write(str(error_count) + "-  " + file_name + " 's code dont match with\n    " + unique_key_file_name + "\n")
                        text_file.write("------------------------------------------------------------------------------------------------ \n")
                else:
                    with open("Error-Logs.txt", "a") as text_file:
                        text_file.write(str(error_count) + "-  " + file_name + " 's code dont match with\n    "  + unique_key_file_name + "\n")
                        text_file.write( "------------------------------------------------------------------------------------------------ \n")


with open("Error-Logs.txt", "a") as text_file:
    text_file.write("Total Error:"+str(error_count)+"   Time:" + str(datetime.datetime.now().time()) + "\n")
    text_file.write("------------------------------------------------------------------------------------------------ \n")






