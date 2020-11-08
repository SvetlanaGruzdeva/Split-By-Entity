# SplitByEntity - splits one consolidated file in separate files, one file per entity

import os, logging, openpyxl, pandas as pd, win32com.client as win32
from tkinter.filedialog import askopenfilename

logging.basicConfig(filename='C:\\Users\\gruzd\\Documents\\Python_Scripts\\Excel\\logs_opt2.txt', level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')
logging.disable(logging.CRITICAL)

entitiesList = []
consolFile = os.path.abspath(askopenfilename()) # Selected by user from browser
region = input('Type in region: ')

if region not in consolFile:
    print('Wrong region name, programm has been stopped.')
    exit()

# Get list of entity codes
ws = pd.read_excel(consolFile, sheet_name='Summary', header=3)
entitiesList = ws['Company Code'].unique().tolist()[1:-1]
logging.debug('List of unique entities as per consolidated file: ')
logging.debug(entitiesList)

# TODO: Remove all unnecessary lines
for entity in entitiesList:
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    wb = xl.Workbooks.Open(consolFile)
    ws = wb.Worksheets('Summary')

    logging.info('Deleting lines...')
    for rowNum in range ((ws.UsedRange.Rows.Count-1), 5, -1):
        logging.debug('Entity code when DELETE loop started: ' + entity)
        if ws.Cells(rowNum, 4).Value != entity and ws.Cells(rowNum, 4).Value != 'FORMULA':
            logging.debug('Row number is ' + str(rowNum))
            logging.debug('Row value is ' + str(ws.Cells(rowNum, 4).Value) + ' and entity value is ' + entity)
            ws.Rows(rowNum).EntireRow.Delete()

    logging.info('Deleting is finished.')
    logging.info('Entity loop finished.')    

    newFileName = consolFile.replace(region, entity)
    wb.SaveAs(newFileName)
    wb.Close()
    xl.Quit()

    # Protect sheet/all sheets
    wb = openpyxl.load_workbook(newFileName, keep_vba=True)
    for sheet in wb.sheetnames:
        ws = wb.get_sheet_by_name(sheet)
        ws.protection.password = '1234'
        ws.protection.sheet = True
        ws.protection.autoFilter = False

    # Save on a tab with instructions
    if region == 'KZ':
        insctuctions = 'Инструкции'
    else:
        insctuctions = 'Instructions'
    wb.active = wb.sheetnames.index(insctuctions)
    wb['Summary'].views.sheetView[0].tabSelected = False
    wb.save(newFileName)