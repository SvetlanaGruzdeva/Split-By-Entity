# SplitByEntity - splits one consolidated file in separate files, one file per entity

import os, logging, pandas as pd
from tkinter.filedialog import askopenfilename

logging.basicConfig(filename='C:\\Users\\gruzd\\Documents\\Python_Scripts\\Excel\\logs_opt2.txt', level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')
# logging.disable(logging.CRITICAL)

entitiesList = []
# consolFile = askopenfilename() # Selected by user from browser
consolFile = 'C:\\Users\\gruzd\\Documents\\Python_Scripts\\Excel\\KZ_G&A_planning_template_FCST2_2020.xlsm'
# region = input('Type in region: ')
region = 'KZ'

if region not in consolFile:
    print('Wrong region name, programm has been stopped.')
    exit()

# TODO: Get list of entity codes
ws = pd.read_excel(consolFile, sheet_name='Summary', header=3) # can also index sheet by name or fetch all sheets
entitiesList = ws['Company Code'].unique().tolist()[1:]
logging.debug(entitiesList)

# TODO: Remove all unnecessary lines

# TODO: Protect sheet/all sheets

# TODO: Open tab with instructions

# TODO: save file with new name