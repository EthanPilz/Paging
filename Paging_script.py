from openpyxl import Workbook
from openpyxl.styles.borders import Border, Side, BORDER_THIN
from openpyxl import load_workbook
import pandas as pd
from datetime import datetime
import easygui
import os
import copy
from Configuration import writeConfigFile, readConfigFile, getBordersOption, getRequestnotesOption

def getInputFile():
    inputPath = easygui.fileopenbox()

    ## Make sure file is excel
    #if not inputFile.lower().endswith('.xlsx'):
        #raise MemoryError("Only Excel (.xlsx) file types are accepted.")
        #raise SystemExit

    return inputPath

def getOutputPath():
    outputPath = easygui.diropenbox(default = "~/Downloads/")
    return outputPath + '/PagingList_' + datetime.now().strftime("%Y-%m-%d %Ih%Mm %p") + '.xlsx'

def formatTable(wb, titleWidth, requestWidth, locationWidth):

    config = readConfigFile()
    print("Printing the value gives:", (config.get('Cell Borders', 'borders'))) #Shows up as false (what I want, since I modified the config.ini)
    bordersOption = getBordersOption
    print("getBordersOption returns:", bordersOption) #Shows up as true REGARDLESS of config.ini modification
    requestnotesOption = getRequestnotesOption()

    ws = wb.active

    ## Delete unneeded info from table
    ws.delete_cols(2, 14)
    ws.delete_cols(5, 4)

    ## Shift all cells to the right by one
    ws.move_range("A1:D100", rows=0, cols=1)

    ## Move call number column all the way to the left
    ws.move_range("E1:E100", rows=0, cols=-4)

    ## Save the output file
    wb.save(outputFile)

    ## Read by default 1st sheet of an excel file
    df1 = pd.read_excel(outputFile)

    ## Sort by Location then Call Number
    df2 = df1.sort_values(['Location', 'Call Number'])
    df2.to_excel(outputFile)

    ## Load back in the excel file from pandas output
    wb = load_workbook(filename = outputFile)
    ws = wb.active

    ws.delete_cols(1, 1)

    ## Wrap text
    for row in ws.iter_rows():
        for cell in row:      
            alignment = copy.copy(cell.alignment)
            alignment.wrapText=True
            alignment.vertical = "top"
            cell.alignment = alignment

    thin_border = Border(
    left=Side(border_style=BORDER_THIN, color='00000000'),
    right=Side(border_style=BORDER_THIN, color='00000000'),
    top=Side(border_style=BORDER_THIN, color='00000000'),
    bottom=Side(border_style=BORDER_THIN, color='00000000')
    )

    ## Make columns the correct width for visual clarity
    dims = {}
    for row in ws.rows:
        for cell in row:
            if(bordersOption):
                cell.border = thin_border
            if cell.value:
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
    for col, value in dims.items():
        ws.column_dimensions[col].width = value

    ## Set column B (title) width, since it is stubborn
    ws.column_dimensions['B'].width = titleWidth

    ## Set column C (Request Notes) width if it's too large, since it makes the file unreadable
    if (ws.column_dimensions['C'].width > 50):
        ws.column_dimensions['C'].width = requestWidth

    ## Set column D (Location) width, since it wraps poorly
    ws.column_dimensions['D'].width = locationWidth

    ## Set print area to only nonzero cells
    printArea = 'A1:D'
    printArea = printArea + str(df2.shape[0] + 1)
    ws.print_area = printArea

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = True

    ## Save the output file
    wb.save(outputFile)
    return ws

writeConfigFile()
inputFile = getInputFile()
outputFile = getOutputPath()

## Load file as xlsx
wb = load_workbook(filename = inputFile)



## Run formatTable method. 
## Params: formatTable(workbook, titleWidth, requestWidth, locationWidth)
ws = formatTable(wb, 75, 20, 15)

## Print the file
#os.startfile(outputFile, "print")