Attribute VB_Name = "ConvertBinHex"
'
'Hex•¶š—ñ‚©‚çBin•¶š—ñ‚ğ•Ô‚·
'‹ó•¶šw’è‚Ìê‡‚Í#VALUE!
'Hex•¶š—ñ‚Å‚Í‚È‚¢ê‡‚Í#NUM!‚ğ•Ô‚·
'
Public Function convHexIntToBinInt(ByVal hex As String) As Variant
    
    Dim stringBuilder() As String
    
    On Error GoTo ERR
    
    lenOfHex = Len(hex)
    
    If (lenOfHex = 0) Then
        convHexIntToBinInt = CVErr(xlErrValue) '#VALUE!‚ğ•Ô‚·
        Exit Function
        
    End If
    
    For cnt = 1 To lenOfHex
        
        ReDim Preserve stringBuilder(cnt - 1) '—ÌˆæŠg’£
        stringBuilder(cnt - 1) = WorksheetFunction.Hex2Bin(Mid(hex, cnt, 1), 4) '•¶š—ñ’Ç‰Á
        
    Next cnt
    
    convHexIntToBinInt = Join(stringBuilder, vbNullString) '•¶š—ñ˜AŒ‹
    Exit Function
    
ERR:
    convHexIntToBinInt = CVErr(xlErrNum) '#NUM!‚ğ•Ô‹p
    Exit Function
    
End Function

'
'Bin•¶š—ñ‚©‚çHex•¶š—ñ‚ğ•Ô‚·
'‹ó•¶šw’è‚Ìê‡‚Í#VALUE!
'Bin•¶š—ñ‚Å‚Í‚È‚¢ê‡‚Í#NUM!‚ğ•Ô‚·
'
Public Function convBinIntToHexInt(ByVal bin As String) As Variant
    
    Dim toConvertBin As String
    Dim stringBuilder() As String
    
    lenOfbin = Len(bin)
    
    If (lenOfbin = 0) Then
        convBinIntToHexInt = CVErr(xlErrValue) '#VALUE!‚ğ•Ô‚·
        Exit Function
        
    End If
    
    modNum = lenOfbin Mod 4
    
    toConvertBin = IIf(modNum = 0, "", String(4 - modNum, "0")) & bin '4•¶š‚¸‚Âˆ—‚Å‚«‚é—l‚É0–„‚ß
    
    On Error GoTo ERR
    
    cntMax = Len(toConvertBin)
    For cnt = 1 To cntMax Step 4
        ReDim Preserve stringBuilder(cnt - 1) '—ÌˆæŠg’£
        stringBuilder(cnt - 1) = WorksheetFunction.Bin2Hex(Mid(toConvertBin, cnt, 4), 1) '•¶š—ñ’Ç‰Á
        
    Next cnt
    
    convBinIntToHexInt = Join(stringBuilder, vbNullString)
    Exit Function
    
ERR:
    convBinIntToHexInt = CVErr(xlErrNum) '#NUM!‚ğ•Ô‹p
    Exit Function
    
End Function
