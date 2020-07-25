"""
This script compares state of subsequent bank accounts to the difference
between earnings and spendings so far to check the ballance
"""

import os
import json

import sys
sys.path.append("..")
from classes import MyWorkbook


# 1. Get state of bank accounts from json file
with open("balance_check.json", encoding='utf-8') as file:
    bank_accounts_state = json.load(file)

current_cash = sum(list(bank_accounts_state.values()))

# 2. Get data from current month
folder_path_file = open("../path.txt", "r")
folder_path = folder_path_file.read()
folder_path_file.close()

month_folder_path = folder_path + "/monthly data"
month_file_name = os.listdir(month_folder_path)[-1]

month_path = os.path.join(month_folder_path, month_file_name)

month_workbook = MyWorkbook(month_path)
month_worksheet = month_workbook.sheets_list[0]
month_spendings = month_worksheet.sum_total
month_incomes = month_worksheet.incomes

# 3. Get data from total workbook
total_path = folder_path + "/total data/total.xlsx"
total_workbook = MyWorkbook(total_path)

total_spendings = 0
total_incomes = 0

for sheet in total_workbook.sheets_list:
    total_spendings += sheet.sum_total
    total_incomes += sheet.incomes

# 4. Collation of incomes, spendings and cash
balance = current_cash + total_spendings + month_spendings - total_incomes \
          - month_incomes

balance_dict = {"balance": round(balance, 2),
                "current_money": round(current_cash, 2),
                "all_incomes": round(total_incomes + month_incomes, 2),
                "all_spendings": round(total_spendings + month_spendings, 2)}

with open(folder_path+'/!Reports/balance check/bilans {}.json'
          .format(month_file_name[:-5]), 'w') as f:
    json.dump(balance_dict, f)
