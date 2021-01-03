import numpy as np
import OpenExcel
import Helpers

# Get params from command line
strRulesExcelFileName ="DM Devis eBooks Reflow - variantes et questions.xlsx"

arrRules = np.chararray((1,1))
ret = OpenExcel.ReadExcelFileToArray(strRulesExcelFileName,arrRules)
if not(ret):
    print("Failure to read " + strRulesExcelFileName )
else: print("Rules read OK")
for row in arrRules.rows:
    for cell in row:
        print(cell.value)

Helpers.PrintArray(arrRules)