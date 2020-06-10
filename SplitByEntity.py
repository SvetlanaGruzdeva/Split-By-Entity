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

wb = openpyxl.load_workbook(consolFile, data_only=True, keep_vba=True) # data_only - to show cells value, not formula
ws = wb.get_sheet_by_name('Summary')
column = ws['D']
columnList = [column[x].value for x in range(6, len(column))]
logging.debug(columnList)
for value in columnList:
    if value not in entitiesList:
        entitiesList.append(value)
logging.debug(entitiesList)

logging.info('Deleting lines...')
for rowNum in range (ws.max_row, 5, -1):
    if ws.cell(row=rowNum, column=4).value != 'KZ101' and ws.cell(row=rowNum, column=4).value != 'FORMULA':
        ws.delete_rows(rowNum)

logging.info('Deleting is finished')
wb.save('G&A_planning_template_FCST2_2020.xlsm')
