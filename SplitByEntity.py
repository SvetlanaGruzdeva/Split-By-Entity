# SplitByEntity - splits one consolidated file in separate files, one file per entity

import openpyxl, os, logging
from pprint import pprint

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')
# logging.disable(logging.CRITICAL)
os.chdir('C:\\Users\\gruzd\\Documents\\Python_Scripts\\Excel')

columnList = []
entitiesList = []
# TODO: Workbook should be selected by user
consolFile = 'KZ_G&A_planning_template_FCST2_2020.xlsm'
region = input('Type in region: ')
# region = 'KZ' # Temporary

wb = openpyxl.load_workbook(consolFile, data_only=True, keep_vba=True) # data_only - to show cells value, not formula
ws = wb.get_sheet_by_name('Summary')
column = ws['D']
columnList = [column[x].value for x in range(6, len(column))]
logging.debug(columnList)
for value in columnList:
    if value not in entitiesList:
        entitiesList.append(value)
logging.debug(entitiesList)

entitiesList = ['KZ101', 'KZ102', 'KZ104', 'KZ104A', 'KZ104B', 'KZ104C']

logging.info('Started entity loop.')
for entity in entitiesList:
    logging.debug(entity)
    logging.info('Deleting lines...')
    for rowNum in range (ws.max_row, 5, -1):
        if ws.cell(row=rowNum, column=4).value != entity and ws.cell(row=rowNum, column=4).value != 'FORMULA':
            ws.delete_rows(rowNum)
    logging.info('Deleting is finished.')
    logging.info('Entity loop finished.')    
    # ws.protection.password = '' # Set password if necessary
    # ws.protection.sheet = True
    
    newFileName = consolFile.replace(region, entity)
    wb.save(newFileName)
