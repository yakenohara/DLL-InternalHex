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
Public Function convDecPntStrToCDoubleAndGetInternalHex(ByVal str As String) As Variant
    
    convDecPntStrToCDoubleAndGetInternalHex = callConvDecStrToOperandAndGetInternalHex(str, 0)
    
End Function

'
'�������10�ifloat�l�Ƃ��āA
'float�ɃL���X�g�������̓���hex�\����Ԃ�
'�����񂪐��l�ɕϊ��ł��Ȃ��ꍇ��#NUM!��Ԃ�
'dll�����݂��Ȃ��ꍇ��#VALUE!��Ԃ�
'
Public Function convDecPntStrToCFloatAndGetInternalHex(ByVal str As String) As Variant
    
    convDecPntStrToCFloatAndGetInternalHex = callConvDecStrToOperandAndGetInternalHex(str, 1)
    
End Function

'
'�������10�ilong�l�Ƃ��āA
'float�ɃL���X�g�������̓���hex�\����Ԃ�
'�����񂪐��l�ɕϊ��ł��Ȃ��ꍇ��#NUM!��Ԃ�
'dll�����݂��Ȃ��ꍇ��#VALUE!��Ԃ�
'
Public Function convDecIntStrToCLongAndGetInternalHex(ByVal str As String) As Variant
    
    convDecIntStrToCLongAndGetInternalHex = callConvDecStrToOperandAndGetInternalHex(str, 2)
    
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
        ret = CVErr(xlErrNum) '#NUM!��Ԃ�
        
        
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
Public Function addtionDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    addtionDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 0)
End Function

'
'double�̓����\��(hex)�Ō��Z����
'
Public Function substractionDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    substractionDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 1)
End Function

'
'double�̓����\��(hex)�ŏ�Z����
'
Public Function multiplicationDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    multiplicationDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 2)
End Function

'
'double�̓����\��(hex)�ŏ��Z����
'
Public Function divisionDoubleByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    divisionDoubleByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 0, 3)
End Function

'
'Float�̓����\��(hex)�ŉ��Z����
'
Public Function addtionFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    addtionFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 0)
End Function

'
'Float�̓����\��(hex)�Ō��Z����
'
Public Function substractionFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    substractionFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 1)
End Function

'
'Float�̓����\��(hex)�ŏ�Z����
'
Public Function multiplicationFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    multiplicationFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 2)
End Function

'
'Float�̓����\��(hex)�ŏ��Z����
'
Public Function divisionFloatByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    divisionFloatByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 1, 3)
End Function

'
'Long�̓����\��(hex)�ŉ��Z����
'
Public Function addtionLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    addtionLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 0)
End Function

'
'Long�̓����\��(hex)�Ō��Z����
'
Public Function substractionLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    substractionLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 1)
End Function

'
'Long�̓����\��(hex)�ŏ�Z����
'
Public Function multiplicationLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    multiplicationLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 2)
End Function

'
'Long�̓����\��(hex)�ŏ��Z����
'
Public Function divisionLongByInternalHex(ByVal firstValue As String, ByVal secondValue As String) As Variant
    divisionLongByInternalHex = callOperateArithmeticByInternalHex(firstValue, secondValue, 2, 3)
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
        ret = CVErr(xlErrValue) '#VALUE!��ԋp
        
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

