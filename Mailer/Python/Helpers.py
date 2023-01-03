
def format_HTML_body(strSalut,isSingle,isVous,strAjout,strSignature,strSender,strType):                     
    strSalut = strSalut.strip()
    if strSalut=='':
        strHTMLBody ='err: no Salutation'
        return strHTMLBody
    strSalut += ',' 
    strSalut = strSalut.replace(',,',',')
    strHTMLBody = strSalut

    strHTMLBody += '<br/>'
    strHTMLBody += '<br/>'
    if isVous:
        strHTMLBody += 'Belle et heureuse année à vous et à vos proches !'
    else:
        strHTMLBody += 'Belle et heureuse année à toi et à tes proches !'
    #
    strHTMLBody += '<br/>'
    # one liner (2023)::
    strHTMLBody += "<b>Que celle-ci soit calme et qu'elle nous apporte à tous santé et douceur.</b>"
    # multi lines (2022) ::
    # strHTMLBody += "<br/><b>Que celle-ci soit Stimulante et Doux,</b>"
    # strHTMLBody += "<br/><b>&nbsp; Qu'elle nous apporte à tous</b>"
    # strHTMLBody += "<br/><b>&nbsp;&nbsp; Santé et Sérénité, et</b>"
    # strHTMLBody += "<br/><b>&nbsp;&nbsp;&nbsp; Que nous restions Calmes dans les tempêtes.</b>" 

    # strAjout 
    strAjout = strAjout.strip()
    if strAjout == '_NONE':
        strAjout= ' '
    else :    
        strHTMLBody += '<br/>'
        strHTMLBody += '<br/>'
        if strAjout == '_COMMENT':
            if isVous:
                strAjout= 'Comment allez-vous ?'
            else:
                strAjout = 'Comment vas-tu ?'
        elif strAjout == '_JESPERE':
            if isVous:
                strAjout = "J'espère que vous allez bien."
            else:
                strAjout = "J'espère que tu vas bien."
        else:
            strAjout += ' ' # the explicit ajout is kept !
        # Only if it is NOT _NONE    
        strHTMLBody += strAjout 
    # ............................

    # strSignature
    strSignature = strSignature.strip()
    if strSignature=='':
        strHTMLBody ='err: no Signature'
        return strHTMLBody

    if strSignature == '_NONE':
        strSignature = ' '
    else:    
        strHTMLBody += '<br/>'
        strHTMLBody += '<br/>'
        if strSignature == '_BIEN':
            if isVous:
                strSignature = 'Bien à vous,'
            else:
                strSignature = 'Bien à toi,'
        else:
            # If explicit signature and it doesn't have a comma at the end :
            if not strSignature.endswith(',') :
                strSignature += ',' 
        # Only if it is NOT _NONE    
        strHTMLBody += strSignature
   # ............................

    # strSignature
    strHTMLBody += '<br/>'
    strSender = strSender.strip()
    if strSender=='':
        strSender ='Shalev'
    else:    
        strSender += ' '
    strHTMLBody += strSender
   # ............................

    return strHTMLBody


def minutes2hour(iMinutes,iSeconds):
    # the %60  around seconds is a mere protection
    strTime = ' ' + str(iMinutes//60).zfill(2)  +':' +str(iMinutes%60).zfill(2) +':' + str(iSeconds%60).zfill(2)
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
