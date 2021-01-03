import numpy as np
import OpenExcel
import Helpers
from datetime import datetime

# Get params from command line
strRulesExcelFileName ="DM Devis eBooks Reflow - variantes et questions.xlsx"

nRows=0
nCells=0
wb,nRows,nCells = OpenExcel.LoadExcelWb( strRulesExcelFileName)
if not(wb):
    print("Failure to read " + strRulesExcelFileName )
    exit

# arrRules = np.chararray((nRows+1,nCells+1),unicode=True)
# arrRules = np.array((nRows+1,nCells+1,dtype=np.str_))
arrRules = np.chararray((nRows+1,nCells+1))  #  Works but shouts on first non ascii char (>128)


ret = OpenExcel.ReadExcelWbToArray(wb,arrRules)
if not(ret):
    print("Failure to read " + strRulesExcelFileName )
else: 
    print("=================== ")
    print("Rules read OK")
    print(datetime.now())
for row in arrRules:
    for cell in row:
        print(cell.value)
#         print(cell.tostring())

Helpers.PrintArray_1(arrRules)

# Helpers.PrintArray_2(arrRules)