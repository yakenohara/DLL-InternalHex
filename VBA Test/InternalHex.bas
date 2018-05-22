Attribute VB_Name = "InternalHex"
'dllインポート

#If Win64 Then

Private Declare PtrSafe Function convDecStrToOperandAndGetInternalHex _
                        Lib "InternalHex_x64.dll" _
                            (ByVal toWriteStr As String, _
                             ByVal lenOfToWriteStr As Long, _
                             ByVal toConvStr As String, _
                             ByVal lenOfToConvStr As Long, _
                             ByVal typ As Long) As Long

Private Declare PtrSafe Function operateArithmeticByInternalHex _
                        Lib "InternalHex_x64.dll" _
                            (ByVal val1 As String, _
                             ByVal val2 As String, _
                             ByVal sum As String, _
                             ByVal lenOfSum As Long, _
                             ByVal operandType As Long, _
                             ByVal operateType As Long) As Long

Private Declare PtrSafe Function getSizeOfOperandExp _
                        Lib "InternalHex_x64.dll" _
                            (ByVal typ As Long) As Long

#Else

Private Declare Function convDecStrToOperandAndGetInternalHex _
                        Lib "InternalHex_x86.dll" _
                            (ByVal toWriteStr As String, _
                             ByVal lenOfToWriteStr As Long, _
                             ByVal toConvStr As String, _
                             ByVal lenOfToConvStr As Long, _
                             ByVal typ As Long) As Long

Private Declare Function operateArithmeticByInternalHex _
                        Lib "InternalHex_x86.dll" _
                            (ByVal val1 As String, _
                             ByVal val2 As String, _
                             ByVal sum As String, _
                             ByVal lenOfSum As Long, _
                             ByVal operandType As Long, _
                             ByVal operateType As Long) As Long

Private Declare Function getSizeOfOperandExp _
                        Lib "InternalHex_x86.dll" _
                            (ByVal typ As Long) As Long
                            
#End If

'定数
Const BUF_SIZE As Long = 255 'バッファ文字列長
Const DLL_NAME As String = "getInternalHexFromDecStr.dll" 'dll名

Dim wsh As Object 'chdir用

'
'文字列を10進float値として、
'floatにキャストした時の内部hex表現を返す
'文字列が数値に変換できない場合は#NUM!を返す
'dllが存在しない場合は#VALUE!を返す
'
Public Function convDecStrToCDoubleAndGetInternalHex(ByVal str As String) As Variant
    
    convDecStrToCDoubleAndGetInternalHex = callConvDecStrToOperandAndGetInternalHex(str, 0)
    
End Function

'
'文字列を10進float値として、
'floatにキャストした時の内部hex表現を返す
'文字列が数値に変換できない場合は#NUM!を返す
'dllが存在しない場合は#VALUE!を返す
'
Public Function convDecStrToCFloatAndGetInternalHex(ByVal str As String) As Variant
    
    convDecStrToCFloatAndGetInternalHex = callConvDecStrToOperandAndGetInternalHex(str, 1)
    
End Function

'
'文字列を10進long値として、
'floatにキャストした時の内部hex表現を返す
'文字列が数値に変換できない場合は#NUM!を返す
'dllが存在しない場合は#VALUE!を返す
'
Public Function convDecStrToCLongAndGetInternalHex(ByVal str As String) As Variant
    
    convDecStrToCLongAndGetInternalHex = callConvDecStrToOperandAndGetInternalHex(str, 2)
    
End Function

'
'共通関数
'
Private Function callConvDecStrToOperandAndGetInternalHex(ByVal str As String, ByVal typ As Long) As Variant
    
    '変数宣言
    Dim nowDir As String
    Dim bufStr As String * BUF_SIZE
    Dim wroteLen As Long
    
    Dim ret As Variant
    Dim wasError As Boolean
    
    nowDir = CurDir 'カレントディレクトリ保存
    
    If (wsh Is Nothing) Then
        Set wsh = CreateObject("WScript.Shell")
    End If
    
    wsh.CurrentDirectory = ThisWorkbook.Path 'カレントディレクトリ変更
    
    On Error GoTo ERR 'dllが存在しない場合は、ERR:にジャンプ
    
    wasError = False
    wroteLen = convDecStrToOperandAndGetInternalHex(bufStr, BUF_SIZE, str, Len(str), typ) 'dllコール
    
    wsh.CurrentDirectory = nowDir  'カレントディレクトリに戻す
    
    '返却値チェック
    If wasError Then 'dllが存在しない場合
        ret = CVErr(xlErrValue) '#VALUE!を返す
    
    ElseIf (wroteLen < 0) Then 'dllが異常値を返却
        
        If (wroteLen = -3) Then '数値ではなかった
            ret = CVErr(xlErrNum) '#NUM!を返す
            
        Else '上記以外のエラー(メモリ不足等)
            ret = CVErr(xlErrValue) '#VALUE!を返す
            
        End If
        
        
    Else 'dllは正常終了
        ret = Left(bufStr, wroteLen)
        
    End If
    
    callConvDecStrToOperandAndGetInternalHex = ret
    Exit Function
    
ERR:
    wasError = True
    Resume Next
    
End Function

'
'doubleの内部表現(hex)で加算する
'
Public Function addtionCDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    addtionCDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 0)
End Function

'
'doubleの内部表現(hex)で減算する
'
Public Function substractionCDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    substractionCDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 1)
End Function

'
'doubleの内部表現(hex)で乗算する
'
Public Function multiplicationCDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    multiplicationCDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 2)
End Function

'
'doubleの内部表現(hex)で除算する
'
Public Function divisionCDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    divisionCDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 3)
End Function

'
'Floatの内部表現(hex)で加算する
'
Public Function addtionCFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    addtionCFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 0)
End Function

'
'Floatの内部表現(hex)で減算する
'
Public Function substractionCFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    substractionCFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 1)
End Function

'
'Floatの内部表現(hex)で乗算する
'
Public Function multiplicationCFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    multiplicationCFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 2)
End Function

'
'Floatの内部表現(hex)で除算する
'
Public Function divisionCFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    divisionCFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 3)
End Function

'
'Longの内部表現(hex)で加算する
'
Public Function addtionCLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    addtionCLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 0)
End Function

'
'Longの内部表現(hex)で減算する
'
Public Function substractionCLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    substractionCLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 1)
End Function

'
'Longの内部表現(hex)で乗算する
'
Public Function multiplicationCLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    multiplicationCLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 2)
End Function

'
'Longの内部表現(hex)で除算する
'
Public Function divisionCLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    divisionCLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 3)
End Function

'
'共通関数
'
Private Function callOperateArithmeticByInternalHex(ByVal firstValue As String, ByVal secondValue As String, ByVal operandType As Long, ByVal operateType As Long) As Variant
    
    Dim wasError As Boolean
    Dim wroteLen As Long
    
    Dim ans As String * BUF_SIZE '計算結果
    
    Dim ret As Variant
    
    nowDir = CurDir 'カレントディレクトリ保存
    
    If (wsh Is Nothing) Then
        Set wsh = CreateObject("WScript.Shell")
    End If
    
    wsh.CurrentDirectory = ThisWorkbook.Path 'カレントディレクトリ変更
    
    wasError = False
    On Error GoTo ERR 'dllが存在しない場合は、ERR:にジャンプ
    
    wroteLen = operateArithmeticByInternalHex(firstValue, secondValue, ans, BUF_SIZE, operandType, operateType)
    
    wsh.CurrentDirectory = nowDir
    
    If (wasError) Then 'dllが見つからない場合
        ret = CVErr(xlErrValue) '#VALUE!を返却
        
        
    ElseIf (wroteLen <= 0) Then 'dllに対する引き数異常の場合
    
        If (wroteLen = -1) Or (wroteLen = -2) Then 'val1かval2がオペランドに入れられない文字列
            ret = CVErr(xlErrNum) '#NUM!を返却
            
        ElseIf (wroteLen = -7) Then
            ret = CVErr(xlErrDiv0) '#DIV/0!を返却
            
        Else '上記以外のエラー(メモリ不足等)
            ret = CVErr(xlErrValue) '#VALUE!を返却
            
        End If
        
    Else
        ret = Left(ans, wroteLen)
        
    End If
    
    callOperateArithmeticByInternalHex = ret
    
    Exit Function

ERR:
    wasError = True
    Resume Next
    
End Function


'
'doubleのサイズを返す
'
Public Function getSizeOfCDouble() As Variant
        
    getSizeOfCDouble = getSizeOfCOperand(0)
    
End Function

'
'floatのサイズを返す
'
Public Function getSizeOfCFloat() As Variant
    
    getSizeOfCFloat = getSizeOfCOperand(1)
    
End Function

'
'longのサイズを返す
'
Public Function getSizeOfCLong() As Variant
    
    getSizeOfCLong = getSizeOfCOperand(2)
    
End Function

'
'共通関数
'
Private Function getSizeOfCOperand(ByVal typ As Integer) As Variant

    Dim wasError As Boolean
    
    nowDir = CurDir 'カレントディレクトリ保存
    
    If (wsh Is Nothing) Then
        Set wsh = CreateObject("WScript.Shell")
    End If
    
    wsh.CurrentDirectory = ThisWorkbook.Path 'カレントディレクトリ変更
    
    wasError = False
    On Error GoTo ERR 'dllが存在しない場合は、ERR:にジャンプ
    
    getSizeOfCOperand = getSizeOfOperandExp(typ)
    
    wsh.CurrentDirectory = nowDir
    
    If (wasError) Then
        getSizeOfCOperand = CVErr(xlErrValue) '#VALUE!を返却
    End If
    
    Exit Function

ERR:
    wasError = True
    Resume Next
    
End Function
