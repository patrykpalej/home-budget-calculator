import pandas as pd


config_txt = pd.read_csv("config/spends_to_find.txt", header=None)
config_list = list(config_txt[0])

print(config_list)
