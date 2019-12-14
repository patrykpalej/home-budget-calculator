import os

args_list = ['10 18', '11 18', '12 18', '01 19', '02 19', '03 19', '04 19',
             '05 19', '06 19', '07 19', '08 19', '09 19', '10 19', '11 19']

for arg in args_list:
    os.system('python monthAnalysis.py ' + arg)
