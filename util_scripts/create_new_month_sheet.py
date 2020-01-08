"""
This script creates a new .xlsx file for monthly data and a shortcut for it
"""

import os
import openpyxl
from shutil import copyfile
import winshell


# 1. Get folder path
folder_path_file = open("../path.txt", "r")
root_path = folder_path_file.read()
month_folder_path = root_path + "/monthly data/"
folder_path_file.close()

# 2. Get names of the last file and the template
last_file_name = os.listdir(month_folder_path)[-1]
last_file_path = month_folder_path + last_file_name
template_file_path = month_folder_path + "/!Template.xlsx"

# 3. Create name of the new file
year_num = last_file_name[2:4]
month_num = last_file_name[5:7]

new_month_num = ((str(int(month_num)+1)
                  if int(month_num)+1 > 9 else "0"+str(int(month_num)+1))
                 if month_num != "12" else "01")
new_year_num = (str(int(year_num)+1) if month_num == "12" else year_num)

new_file_name = "/20" + new_year_num + "." + new_month_num + ".xlsx"

# 4. Copy a file, change the sheet name and create a shortcut
copyfile(template_file_path, month_folder_path + new_file_name)

wb = openpyxl.load_workbook(month_folder_path + new_file_name)
ws = wb.active
ws.title = new_month_num
wb.save(month_folder_path + new_file_name)

shortcut = winshell.shortcut(month_folder_path + new_file_name)
shortcut.write(root_path + new_file_name[:-5] + " — skrót.lnk")
shortcut.dump()
