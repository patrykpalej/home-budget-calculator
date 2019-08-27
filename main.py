import os, sys, openpyxl
from classes import *

args_list = ['01 99', '02 99', '03 99']

for arg in args_list:
    os.system('python monthAnalysis.py ' + arg)
