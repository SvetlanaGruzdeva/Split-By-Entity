# SplitByEntity - splits one consolidated file in separate files, one file per entity

import os, logging, win32com.client
from tkinter.filedialog import askopenfilename

logging.basicConfig(filename='logs_win32.txt', level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')
# logging.disable(logging.CRITICAL)

columnList = []
entitiesList = []
# consolFile = askopenfilename() # Selected by user from browser
consolFile = 'Copy_KZ_G&A_planning_template_FCST2_2020.xlsm'
# region = input('Type in region: ')
region = 'KZ'

if region not in consolFile:
    print('Wrong region name, programm has been stopped.')
    exit()

# TODO: Get list of entity codes
xl = win32com.client.Dispatch('Excel.Application')
# TODO: select file from browser
wb = xl.Workbooks.Open("C:\\Users\\gruzd\\Documents\\Python_Scripts\\Excel\\Copy_KZ_G&A_planning_template_FCST2_2020.xlsm")
ws = wb.Worksheets('Summary')


# TODO: Remove all unnecessary lines

# TODO: Protect sheet/all sheets

# TODO: Open tab with instructions

# TODO: save file with new name