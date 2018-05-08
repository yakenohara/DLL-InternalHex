Attribute VB_Name = "ConvertBinHex"
'
'Hex�����񂩂�Bin�������Ԃ�
'�󕶎��w��̏ꍇ��#VALUE!
'Hex������ł͂Ȃ��ꍇ��#NUM!��Ԃ�
'
Public Function convHexIntToBinInt(ByVal hex As String) As Variant
    
    Dim stringBuilder() As String
    
    On Error GoTo ERR
    
    lenOfHex = Len(hex)
    
    If (lenOfHex = 0) Then
        convHexIntToBinInt = CVErr(xlErrValue) '#VALUE!��Ԃ�
        Exit Function
        
    End If
    
    For cnt = 1 To lenOfHex
        
        ReDim Preserve stringBuilder(cnt - 1) '�̈�g��
        stringBuilder(cnt - 1) = WorksheetFunction.Hex2Bin(Mid(hex, cnt, 1), 4) '������ǉ�
        
    Next cnt
    
    convHexIntToBinInt = Join(stringBuilder, vbNullString) '������A��
    Exit Function
    
ERR:
    convHexIntToBinInt = CVErr(xlErrNum) '#NUM!��ԋp
    Exit Function
    
End Function

'
'Bin�����񂩂�Hex�������Ԃ�
'�󕶎��w��̏ꍇ��#VALUE!
'Bin������ł͂Ȃ��ꍇ��#NUM!��Ԃ�
'
Public Function convBinIntToHexInt(ByVal bin As String) As Variant
    
    Dim toConvertBin As String
    Dim stringBuilder() As String
    
    lenOfbin = Len(bin)
    
    If (lenOfbin = 0) Then
        convBinIntToHexInt = CVErr(xlErrValue) '#VALUE!��Ԃ�
        Exit Function
        
    End If
    
    modNum = lenOfbin Mod 4
    
    toConvertBin = IIf(modNum = 0, "", String(4 - modNum, "0")) & bin '4�����������ł���l��0����
    
    On Error GoTo ERR
    
    cntMax = Len(toConvertBin)
    For cnt = 1 To cntMax Step 4
        ReDim Preserve stringBuilder(cnt - 1) '�̈�g��
        stringBuilder(cnt - 1) = WorksheetFunction.Bin2Hex(Mid(toConvertBin, cnt, 4), 1) '������ǉ�
        
    Next cnt
    
    convBinIntToHexInt = Join(stringBuilder, vbNullString)
    Exit Function
    
ERR:
    convBinIntToHexInt = CVErr(xlErrNum) '#NUM!��ԋp
    Exit Function
    
End Function
