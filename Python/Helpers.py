def minutes2hour(iMinutes):
    strTime = ' ' + str(iMinutes//60).zfill(2)  +':' +str(iMinutes%60).zfill(2) +':00' 
    return strTime
    
def PrintArray_2(arrStr):
    nRow=0
    for row in arrStr:
        nRow += 1
        nCell=0
        for cell in row:
            nCell += 1
            print(str(cell))
            print(arrStr[nCell,nRow])
    return True

def PrintArray_1(arrStr):
    nRow=0
    for row in arrStr:
        nRow += 1
        nCell=0
        for cell in row:
            nCell += 1
            print(f'R={nRow:5d};C{nCell:5d} ==> {cell.value:150}')
    return True
