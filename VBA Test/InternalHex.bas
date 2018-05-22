Attribute VB_Name = "InternalHex"
'dll�C���|�[�g

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

'�萔
Const BUF_SIZE As Long = 255 '�o�b�t�@������
Const DLL_NAME As String = "getInternalHexFromDecStr.dll" 'dll��

Dim wsh As Object 'chdir�p

'
'�������10�ifloat�l�Ƃ��āA
'float�ɃL���X�g�������̓���hex�\����Ԃ�
'�����񂪐��l�ɕϊ��ł��Ȃ��ꍇ��#NUM!��Ԃ�
'dll�����݂��Ȃ��ꍇ��#VALUE!��Ԃ�
'
Public Function convDecStrToCDoubleAndGetInternalHex(ByVal str As String) As Variant
    
    convDecStrToCDoubleAndGetInternalHex = callConvDecStrToOperandAndGetInternalHex(str, 0)
    
End Function

'
'�������10�ifloat�l�Ƃ��āA
'float�ɃL���X�g�������̓���hex�\����Ԃ�
'�����񂪐��l�ɕϊ��ł��Ȃ��ꍇ��#NUM!��Ԃ�
'dll�����݂��Ȃ��ꍇ��#VALUE!��Ԃ�
'
Public Function convDecStrToCFloatAndGetInternalHex(ByVal str As String) As Variant
    
    convDecStrToCFloatAndGetInternalHex = callConvDecStrToOperandAndGetInternalHex(str, 1)
    
End Function

'
'�������10�ilong�l�Ƃ��āA
'float�ɃL���X�g�������̓���hex�\����Ԃ�
'�����񂪐��l�ɕϊ��ł��Ȃ��ꍇ��#NUM!��Ԃ�
'dll�����݂��Ȃ��ꍇ��#VALUE!��Ԃ�
'
Public Function convDecStrToCLongAndGetInternalHex(ByVal str As String) As Variant
    
    convDecStrToCLongAndGetInternalHex = callConvDecStrToOperandAndGetInternalHex(str, 2)
    
End Function

'
'���ʊ֐�
'
Private Function callConvDecStrToOperandAndGetInternalHex(ByVal str As String, ByVal typ As Long) As Variant
    
    '�ϐ��錾
    Dim nowDir As String
    Dim bufStr As String * BUF_SIZE
    Dim wroteLen As Long
    
    Dim ret As Variant
    Dim wasError As Boolean
    
    nowDir = CurDir '�J�����g�f�B���N�g���ۑ�
    
    If (wsh Is Nothing) Then
        Set wsh = CreateObject("WScript.Shell")
    End If
    
    wsh.CurrentDirectory = ThisWorkbook.Path '�J�����g�f�B���N�g���ύX
    
    On Error GoTo ERR 'dll�����݂��Ȃ��ꍇ�́AERR:�ɃW�����v
    
    wasError = False
    wroteLen = convDecStrToOperandAndGetInternalHex(bufStr, BUF_SIZE, str, Len(str), typ) 'dll�R�[��
    
    wsh.CurrentDirectory = nowDir  '�J�����g�f�B���N�g���ɖ߂�
    
    '�ԋp�l�`�F�b�N
    If wasError Then 'dll�����݂��Ȃ��ꍇ
        ret = CVErr(xlErrValue) '#VALUE!��Ԃ�
    
    ElseIf (wroteLen < 0) Then 'dll���ُ�l��ԋp
        
        If (wroteLen = -3) Then '���l�ł͂Ȃ�����
            ret = CVErr(xlErrNum) '#NUM!��Ԃ�
            
        Else '��L�ȊO�̃G���[(�������s����)
            ret = CVErr(xlErrValue) '#VALUE!��Ԃ�
            
        End If
        
        
    Else 'dll�͐���I��
        ret = Left(bufStr, wroteLen)
        
    End If
    
    callConvDecStrToOperandAndGetInternalHex = ret
    Exit Function
    
ERR:
    wasError = True
    Resume Next
    
End Function

'
'double�̓����\��(hex)�ŉ��Z����
'
Public Function addtionCDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    addtionCDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 0)
End Function

'
'double�̓����\��(hex)�Ō��Z����
'
Public Function substractionCDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    substractionCDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 1)
End Function

'
'double�̓����\��(hex)�ŏ�Z����
'
Public Function multiplicationCDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    multiplicationCDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 2)
End Function

'
'double�̓����\��(hex)�ŏ��Z����
'
Public Function divisionCDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    divisionCDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 3)
End Function

'
'Float�̓����\��(hex)�ŉ��Z����
'
Public Function addtionCFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    addtionCFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 0)
End Function

'
'Float�̓����\��(hex)�Ō��Z����
'
Public Function substractionCFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    substractionCFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 1)
End Function

'
'Float�̓����\��(hex)�ŏ�Z����
'
Public Function multiplicationCFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    multiplicationCFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 2)
End Function

'
'Float�̓����\��(hex)�ŏ��Z����
'
Public Function divisionCFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    divisionCFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 3)
End Function

'
'Long�̓����\��(hex)�ŉ��Z����
'
Public Function addtionCLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    addtionCLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 0)
End Function

'
'Long�̓����\��(hex)�Ō��Z����
'
Public Function substractionCLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    substractionCLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 1)
End Function

'
'Long�̓����\��(hex)�ŏ�Z����
'
Public Function multiplicationCLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    multiplicationCLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 2)
End Function

'
'Long�̓����\��(hex)�ŏ��Z����
'
Public Function divisionCLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    divisionCLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 3)
End Function

'
'���ʊ֐�
'
Private Function callOperateArithmeticByInternalHex(ByVal firstValue As String, ByVal secondValue As String, ByVal operandType As Long, ByVal operateType As Long) As Variant
    
    Dim wasError As Boolean
    Dim wroteLen As Long
    
    Dim ans As String * BUF_SIZE '�v�Z����
    
    Dim ret As Variant
    
    nowDir = CurDir '�J�����g�f�B���N�g���ۑ�
    
    If (wsh Is Nothing) Then
        Set wsh = CreateObject("WScript.Shell")
    End If
    
    wsh.CurrentDirectory = ThisWorkbook.Path '�J�����g�f�B���N�g���ύX
    
    wasError = False
    On Error GoTo ERR 'dll�����݂��Ȃ��ꍇ�́AERR:�ɃW�����v
    
    wroteLen = operateArithmeticByInternalHex(firstValue, secondValue, ans, BUF_SIZE, operandType, operateType)
    
    wsh.CurrentDirectory = nowDir
    
    If (wasError) Then 'dll��������Ȃ��ꍇ
        ret = CVErr(xlErrValue) '#VALUE!��ԋp
        
        
    ElseIf (wroteLen <= 0) Then 'dll�ɑ΂���������ُ�̏ꍇ
    
        If (wroteLen = -1) Or (wroteLen = -2) Then 'val1��val2���I�y�����h�ɓ�����Ȃ�������
            ret = CVErr(xlErrNum) '#NUM!��ԋp
            
        ElseIf (wroteLen = -7) Then
            ret = CVErr(xlErrDiv0) '#DIV/0!��ԋp
            
        Else '��L�ȊO�̃G���[(�������s����)
            ret = CVErr(xlErrValue) '#VALUE!��ԋp
            
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
'double�̃T�C�Y��Ԃ�
'
Public Function getSizeOfCDouble() As Variant
        
    getSizeOfCDouble = getSizeOfCOperand(0)
    
End Function

'
'float�̃T�C�Y��Ԃ�
'
Public Function getSizeOfCFloat() As Variant
    
    getSizeOfCFloat = getSizeOfCOperand(1)
    
End Function

'
'long�̃T�C�Y��Ԃ�
'
Public Function getSizeOfCLong() As Variant
    
    getSizeOfCLong = getSizeOfCOperand(2)
    
End Function

'
'���ʊ֐�
'
Private Function getSizeOfCOperand(ByVal typ As Integer) As Variant

    Dim wasError As Boolean
    
    nowDir = CurDir '�J�����g�f�B���N�g���ۑ�
    
    If (wsh Is Nothing) Then
        Set wsh = CreateObject("WScript.Shell")
    End If
    
    wsh.CurrentDirectory = ThisWorkbook.Path '�J�����g�f�B���N�g���ύX
    
    wasError = False
    On Error GoTo ERR 'dll�����݂��Ȃ��ꍇ�́AERR:�ɃW�����v
    
    getSizeOfCOperand = getSizeOfOperandExp(typ)
    
    wsh.CurrentDirectory = nowDir
    
    If (wasError) Then
        getSizeOfCOperand = CVErr(xlErrValue) '#VALUE!��ԋp
    End If
    
    Exit Function

ERR:
    wasError = True
    Resume Next
    
End Function
