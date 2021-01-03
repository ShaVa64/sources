def PrintArray(arrStr):
    nRow=0
    for row in arrStr.rows:
        nRow += 1
        nCell=0
        for cell in row:
            nCell += 1
            print(arrRules[nRow,nCell])
    return True