Attribute VB_Name = "convNPntNPnt"
'<PrivateFunction�p�e�X�g�֐�>---------------------------------------------------------------------------------------------------------------------
'

Public Function TESTaddBinInt(ByVal val1 As String, ByVal val2 As String) As Variant
    
    Dim stsOfSub As Variant
    TESTaddBinInt = addBinInt(val1, val2, stsOfSub)
    
End Function

Public Function TESTaddBinIntByRef1(ByVal val1 As String, ByVal val2 As String) As Variant
    
    Dim stsOfSub As Variant
    x = addBinInt(val1, val2, stsOfSub)
    TESTaddBinIntByRef1 = stsOfSub
    
End Function

Public Function TESTmultipleBinInt(ByVal multiplicand As String, ByVal multiplier As String) As Variant
    
    Dim stsOfSub As Variant
    TESTmultipleBinInt = multipleBinInt(multiplicand, multiplier, stsOfSub)
    
End Function

Public Function TESTmultipleBinIntByRef1(ByVal multiplicand As String, ByVal multiplier As String) As Variant
    
    Dim stsOfSub As Variant
    x = multipleBinInt(multiplicand, multiplier, stsOfSub)
    TESTmultipleBinIntByRef1 = stsOfSub
    
End Function

Public Function TESTdivide(ByVal dividend As String, ByVal divisor As String, ByVal radix As Byte, ByVal numOfFrcDigits As Long) As Variant
    
    Dim remainder As String
    Dim stsOfSub As Variant
    TESTdivide = divide(dividend, divisor, radix, numOfFrcDigits, remainder, stsOfSub)
    
End Function

Public Function TESTdivideByRef1(ByVal dividend As String, ByVal divisor As String, ByVal radix As Byte, ByVal numOfFrcDigits As Long) As Variant
    
    Dim remainder As String
    Dim stsOfSub As Variant
    x = divide(dividend, divisor, radix, numOfFrcDigits, remainder, stsOfSub)
    TESTdivideByRef1 = remainder
    
End Function

Public Function TESTdivideByRef2(ByVal dividend As String, ByVal divisor As String, ByVal radix As Byte, ByVal numOfFrcDigits As Long) As Variant
    
    Dim remainder As String
    Dim stsOfSub As Variant
    x = divide(dividend, divisor, radix, numOfFrcDigits, remainder, stsOfSub)
    TESTdivideByRef2 = stsOfSub
    
End Function

Public Function TESTconvIntPrtOfNPntToIntPrtOfNPnt(ByVal intStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte) As Variant
    
    Dim stsOfSub As Variant
    TESTconvIntPrtOfNPntToIntPrtOfNPnt = convIntPrtOfNPntToIntPrtOfNPnt(intStr, fromRadix, toRadix, stsOfSub)
    
End Function

Public Function TESTconvIntPrtOfNPntToIntPrtOfNPntByRef1(ByVal intStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte) As Variant
    
    Dim stsOfSub As Variant
    ans = convIntPrtOfNPntToIntPrtOfNPnt(intStr, fromRadix, toRadix, stsOfSub)
    TESTconvIntPrtOfNPntToIntPrtOfNPntByRef1 = stsOfSub

End Function

Public Function TESTconvFrcPrtOfNPntToFrcPrtOfNPnt(ByVal frcStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByVal numOfDigits As Long) As Variant
    
    Dim stsOfSub As Variant
    TESTconvFrcPrtOfNPntToFrcPrtOfNPnt = convFrcPrtOfNPntToFrcPrtOfNPnt(frcStr, fromRadix, toRadix, numOfDigits, stsOfSub)
    
End Function

Public Function TESTconvFrcPrtOfNPntToFrcPrtOfNPntByRef1(ByVal frcStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByVal numOfDigits As Long) As Variant
    
    Dim stsOfSub As Variant
    x = convFrcPrtOfNPntToFrcPrtOfNPnt(frcStr, fromRadix, toRadix, numOfDigits, stsOfSub)
    TESTconvFrcPrtOfNPntToFrcPrtOfNPntByRef1 = stsOfSub
    
End Function
'
'--------------------------------------------------------------------------------------------------------------------</PrivateFunction�p�e�X�g�֐�>

'
'2�i���̐������m��a�Z����
'
'!CAUTION!
'    val1, val2 ���L����2�i�l�ł��邩�̓`�F�b�N���Ȃ�
'
Private Function addBinInt(ByVal val1 As String, ByVal val2 As String, ByRef endStatus As Variant) As String

    Dim lenOfVal1 As Long
    Dim lenOfVal2 As Long
    Dim idxOfval As Long
    Dim stringBuilder() As String
    Dim sumH As Boolean
    Dim carH As Boolean
    Dim sumF As Boolean
    Dim carF As Boolean
    
    '���l�񒷃`�F�b�N
    lenOfVal1 = Len(val1)
    lenOfVal2 = Len(val2)
    
    '�a�Z���[�v�O������
    If (lenOfVal1 < lenOfVal2) Then 'val2�̐��l�񒷂̕�������
        val1 = String(lenOfVal2 - lenOfVal1, "0") & val1 '�������킹
        idxOfval = lenOfVal2 '�X�^�[�g�C���f�b�N�X
        
    Else
        val2 = String(lenOfVal1 - lenOfVal2, "0") & val2 '�������킹
        idxOfval = lenOfVal1
        
    End If
    
    ReDim stringBuilder(idxOfval) '���������� + 1 ���̔z���`
    carF = False
    
    '�a�Z���[�v
    Do While (idxOfval > 0)
        
        digitOfVal1 = (Mid(val1, idxOfval, 1) = "1")
        digitOfVal2 = (Mid(val2, idxOfval, 1) = "1")
        
        sumH = (digitOfVal1) Xor (digitOfVal2)
        carH = (digitOfVal1) And (digitOfVal2)
        
        sumF = (sumH) Xor (carF)
        carF = ((sumH) And (carF)) Or (carH)
        
        stringBuilder(idxOfval) = IIf(sumF, "1", "0")
        
        idxOfval = idxOfval - 1 'decrement
        
    Loop
    
    stringBuilder(idxOfval) = IIf(carF, "1", "")
    
    addBinInt = Join(stringBuilder, vbNullString)
    
    
End Function

'
'2�i���̐������m����Z����
'
'!CAUTION!
'    multiplicand, multiplier ���L����2�i�l�ł��邩�̓`�F�b�N���Ȃ�
'
Private Function multipleBinInt(ByVal multiplicand As String, ByVal multiplier As String, ByRef endStatus As Variant) As String
    
    Dim stsOfSub As Variant
    Dim tmpAns As String
    Dim numOfShift As Long
    Dim idxOfMultiplier As Long
    
    tmpAns = "0"
    numOfShift = 0
    idxOfMultiplier = Len(multiplier)
    
    '��Z���[�v
    Do While (idxOfMultiplier > 0)
        
        If (Mid(multiplier, idxOfMultiplier, 1) <> "0") Then
            
            '����(�������� - 1)�V�t�g�������l�Řa�Z����
            tmpAns = addBinInt(tmpAns, (multiplicand & String(numOfShift, "0")), stsOfSub)
            
        End If
        
        idxOfMultiplier = idxOfMultiplier - 1 'decrement
        numOfShift = numOfShift + 1 'increment
        
    Loop
    
    multipleBinInt = tmpAns
    
End Function


'
'�������m����Z����
'
'!CAUTION!
'    multiplicand, multiplier ���L����n�i�l�ł��邩�̓`�F�b�N���Ȃ�
'
Private Function multiple(ByVal multiplicand As String, ByVal multiplier As String, ByVal radix As Byte, ByRef endStatus As Variant) As String
    
    Dim binMultiplicand As String
    Dim binMultiplier As String
    Dim binAns As String
    Dim ans As String
    Dim stsOfSub As Variant
    
    '2�i���ϊ�
    binMultiplicand = convIntPrtOfNPntToIntPrtOfNPnt(multiplicand, radix, 2, statusOfSub)
    binMultiplier = convIntPrtOfNPntToIntPrtOfNPnt(multiplier, radix, 2, statusOfSub)
    
    '2�i���ŏ�Z
    binAns = multipleBinInt(binMultiplicand, binMultiplier, statusOfSub)
    
    '��Z���ʂ��w��n�i�l�ɖ߂�
    ans = convIntPrtOfNPntToIntPrtOfNPnt(binAns, 2, radix, statusOfSub)
    
    multiple = ans
    
End Function

'
'���Z������
'
'�ȉ��̏ꍇ�͋󕶎���ԋp���A
'errCode�ɃG���[�R�[�h���i�[����
'    �E0���̏ꍇ�B(�G���[�R�[�h��#DIV/0!)
'    �Edividend / divisor ��long�^�Ŏ�舵���Ȃ��傫�Ȑ��l������ꍇ�B(�G���[�R�[�h��#NUM!)
'
'numOfFrcDigits:
'    ���߂鏬���_�ȉ��̌���
'    �w�茅���ŏ��Z��ł��؂�
'    (-)�l��ݒ肵���ꍇ�́A�����_�ȉ��͋��߂Ȃ�
'
'remainder
'    ��]
'    (numOfFrcDigits > 0)�̏ꍇ�́A
'    numOfFrcDigits�ł̏�]���i�[����
'    ex:)
'    �y�O��z10 / 8 = 1.2 �]�� 0.4
'    �y���s���@�zx = divide("10", "8", 10, 1, rm, code)
'    �y���ʁz x:012
'            rm:4
'
'!CAUTION!
'    dividend, divisor ���L����n�i�l�ł��邩�̓`�F�b�N���Ȃ�
'    radix��2~16�͈͓̔��ł��鎖�̓`�F�b�N���Ȃ�
'
Private Function divide(ByVal dividend As String, ByVal divisor As String, ByVal radix As Byte, ByVal numOfFrcDigits As Long, ByRef remainder As String, ByRef errCode As Variant) As String

    '�ϐ��錾
    Dim quot As Long '��
    Dim rmnd As Long '�]��
    Dim repTimes As Long 'IndivisibleNumber�ɑ΂���divide��
    Dim digitOfDividend As Long '�ꎞ�폜��
    Dim stringBuilder() As String '���i�[�p
    Dim digitIdxOfDividend As Long 'Division���ʕ�����
    Dim divisorDec As Long
    Dim stsOfSub As Variant
    
    'divisor�̕s�v��0����菜��
    divisor = removeLeft0(divisor, stsOfSub)
    
    '<divisor��10�i�ϊ�>------------------------------------------------------------
    
    tmp = convIntPrtOfNPntToIntPrtOfNPnt(divisor, radix, 10, stsOfSub)
    
    'divisor��Long�^�ϊ�
    On Error GoTo OVERFLOW
    divisorDec = CLng(tmp)
    
    If divisorDec = 0 Then '0���`�F�b�N
        divide = ""
        errCode = CVErr(xlErrDiv0) '#DIV0!��Ԃ�
        Exit Function
    
    ElseIf divisorDec = 1 Then
        divide = dividend '1���̏ꍇ�͂��̂܂ܕԂ�
        Exit Function
        
    End If
    
    '-----------------------------------------------------------</divisor��10�i�ϊ�>
    
    '������
    rmnd = 0
    digitIdxOfDividend = 1
    repTimes = 0
    
    '���s���[�v
    Do
        digitOfDividend = rmnd * radix + convNCharToByte(Mid(dividend, digitIdxOfDividend, 1), stsOfSub) '��ʌ��̗]�� & �Y����
        
        quot = digitOfDividend \ divisorDec '��
        rmnd = digitOfDividend Mod divisorDec '�]��
        
        ReDim Preserve stringBuilder(digitIdxOfDividend - 1) '�̈�g��
        
        '����ǋL
        '���錅�ɑ΂��鏜�Z�̏��́A�����������蕿�Ȃ�
        stringBuilder(digitIdxOfDividend - 1) = convByteToNChar(quot, stsOfSub)
        
        digitIdxOfDividend = digitIdxOfDividend + 1
        
        If (rmnd > 0) And (Len(dividend) < digitIdxOfDividend) Then '�]�肪���邯��ǁA���̌�������
        
            If (numOfFrcDigits > -1) And (repTimes < numOfFrcDigits) Then '�ċA�v�Z�񐔂��w��񐔈ȉ�
                dividend = dividend & "0" '"0"��t��
                
                repTimes = repTimes + 1
                
            End If
            
        End If
        
    Loop While digitIdxOfDividend <= Len(dividend) '�ŏI�����ɓ��B���Ȃ���
    
    If (rmnd = 0) Then '�]�肪0�̂Ƃ�
        remainder = "0"
        
    Else '�]�肪���݂��鎞
        
        remainder = convIntPrtOfNPntToIntPrtOfNPnt(rmnd, 10, radix, statusOfSub)
    
    End If
    
    divide = Join(stringBuilder, vbNullString) '������A��
    
    Exit Function
    
OVERFLOW: '�I�[�o�[�t���[�̏ꍇ
    divide = ""
    errCode = CVErr(xlErrNum) '#NUM!��Ԃ�
    Exit Function
    
End Function

'
'n�i��������n�i�������ɕϊ�����
'
'!CAUTION!
'    intStr���L����(fromRadix)�i�l�ł��邩�̓`�F�b�N���Ȃ�
'    fromRadix,toRadix��2~16�͈͓̔��ł��鎖�̓`�F�b�N���Ȃ�
'
Private Function convIntPrtOfNPntToIntPrtOfNPnt(ByVal intStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByRef endStatus As Variant) As String
    
    Dim stsOfSub As Variant
    Dim retOfTryConvRadix As String
    Dim strLenOfToRadix As Long
    Dim stringBuilder() As String '�ϊ��㕶���񐶐��p
    Dim sizeOfStringBuilder As Long
    Dim rm As String
    
    intStr = removeLeft0(intStr, stsOfSub)
    
    '�ϊ������������ꍇ
    If (fromRadix = toRadix) Then
        convIntPrtOfNPntToIntPrtOfNPnt = intStr '"0"����菜���������̒l��Ԃ�
        Exit Function
        
    End If
    
    'convRadix�ŉ����\���ǂ����`�F�b�N
    retOfTryConvRadix = tryConvRadix(intStr, fromRadix, toRadix, stsOfSub)
    
    If (retOfTryConvRadix <> "") Then '��ϊ��p�e�[�u���ɉ���������
        convIntPrtOfNPntToIntPrtOfNPnt = retOfTryConvRadix
        Exit Function
        
    End If
    
    '�������[�v�O������
    sizeOfStringBuilder = 0
    chOfToRadix = convRadix(fromRadix, toRadix, stsOfSub)
    strLenOfToRadix = Len(chOfToRadix)
    
    '�������[�v - toRadix�ɂ�鏜�Z�ɂ���ĉ������߂� -
    Do While True
        
        If (Len(intStr) <= strLenOfToRadix) Then
            
            retOfTryConvRadix = tryConvRadix(intStr, fromRadix, 10, stsOfSub)
            
            If (retOfTryConvRadix <> "") Then
                
                If (CByte(retOfTryConvRadix) < toRadix) Then '��Ŋ���鐔���Ȃ��Ȃ��� ���K�� retOfTryConvRadix > 0 �Ƃ͂Ȃ� ��
                    Exit Do
                    
                End If
                
            End If
            
        End If
        
        intStr = divide(intStr, chOfToRadix, fromRadix, 0, rm, stsOfSub) '16(10�i�l)�ȉ��ɂ�鏜�Z�Ȃ̂ŁA�I�[�o�[�t���[�͔��������Ȃ�
        
        intStr = removeLeft0(intStr, stsOfSub) '�����̕s�v��"0"����菜��
        
        ReDim Preserve stringBuilder(sizeOfStringBuilder) '�̈�g��
        
        '��]��(toRadix)�i�l�ɕϊ��������ʂ��Z�oDigit
        stringBuilder(sizeOfStringBuilder) = convIntPrtOfNPntToIntPrtOfNPnt(rm, fromRadix, toRadix, stsOfSub)
        
        sizeOfStringBuilder = sizeOfStringBuilder + 1
        
    Loop
    
    ReDim Preserve stringBuilder(sizeOfStringBuilder) '�̈�g��
    
    '�ŏ��Bit�t��
    '��]��(toRadix)�i�l�ɕϊ��������ʂ��Z�oDigit
    stringBuilder(sizeOfStringBuilder) = convIntPrtOfNPntToIntPrtOfNPnt(intStr, fromRadix, toRadix, stsOfSub)
    
    convIntPrtOfNPntToIntPrtOfNPnt = Join(invertStringArray(stringBuilder, stsOfSub), vbNullString) '������A��
    
End Function

'
'n�i��������n�i�������ɕϊ�����
'
'numOfDigits:
'    ���߂錅��
'    0�ȉ����w�肵���ꍇ�́A�󕶎���ԋp����
'
'!CAUTION!
'    frcStr���L����(fromRadix)�i�l�ł��邩�̓`�F�b�N���Ȃ�
'    fromRadix,toRadix��2~16�͈͓̔��ł��鎖�̓`�F�b�N���Ȃ�
'
Private Function convFrcPrtOfNPntToFrcPrtOfNPnt(ByVal frcStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByVal numOfDigits As Long, ByRef endStatus As Variant) As String
    
    
    Dim stringBuilder() As String '�ϊ��㕶���񐶐��p
    Dim sizeOfStringBuilder As Long
    Dim retOfMultiple As String
    
    '�������[�v�O������
    chOfToRadix = convRadix(fromRadix, toRadix, stsOfSub)
    sizeOfStringBuilder = 0
    lenOfFrcStrB = Len(frcStr)
    
    '�������[�v - toRadix�ɂ���Z�ɂ���ĉ������߂� -
    Do While (sizeOfStringBuilder < numOfDigits)
        
        '�����̐ς�0�ɂȂ�����I��
        If (frcStr = "0") Then
            Exit Do
            
        End If
        
        frcStr = multiple(frcStr, chOfToRadix, fromRadix, statusOfSub)
        
        ReDim Preserve stringBuilder(sizeOfStringBuilder) '�̈�g��
        
        '�����������E��
        lenOfFrcStrA = Len(frcStr)
        
        If (lenOfFrcStrA > lenOfFrcStrB) Then
            tmp = Left(frcStr, lenOfFrcStrA - lenOfFrcStrB)
            frcStr = Right(frcStr, lenOfFrcStrB)
            increasedDigits = convIntPrtOfNPntToIntPrtOfNPnt(tmp, fromRadix, toRadix, stsOfSub)
            
        Else
            increasedDigits = "0"
        
        End If
        
        stringBuilder(sizeOfStringBuilder) = increasedDigits '����ǋL
        
        frcStr = removeRight0(frcStr, stsOfSub) ' �E���̕s�v��0����菜��
        
        lenOfFrcStrB = Len(frcStr)
        sizeOfStringBuilder = sizeOfStringBuilder + 1
        
    Loop
    
    convFrcPrtOfNPntToFrcPrtOfNPnt = Join(stringBuilder, vbNullString)

End Function

'
'���l������10�i�l�ł�������Ԃ�
'
'!CAUTION!
'    ch��2~F�͈͓̔��ł��鎖�̓`�F�b�N���Ȃ�
'
Private Function convNCharToByte(ByVal ch As String, ByRef endStatus As Variant) As Byte
    
    Dim toRetByte As Byte
    Dim ascOfA As Integer
    
    ascOfA = Asc("A")
    ascOfCh = Asc(ch)
    
    If (ascOfA <= ascOfCh) Then 'A~F�̏ꍇ
        toRetByte = 10 + (ascOfCh - ascOfA)
    
    Else '0~9�̏ꍇ
        toRetByte = CByte(ch)
    
    End If
    
    convNCharToByte = toRetByte
    
End Function

'
'10�i�l���琔�l������Ԃ�
'
'!CAUTION!
'    byt��0~16�͈͓̔��ł��鎖�̓`�F�b�N���Ȃ�
'
Private Function convByteToNChar(ByVal byt As Byte, ByRef endStatus As Variant) As String
    
    Dim toRetStr As String
    
    If (byt > 9) Then 'A~F�̏ꍇ
        toRetStr = Chr((byt - 10) + Asc("A"))
    
    Else '0~9�̏ꍇ
        toRetStr = Chr(byt + Asc("0"))
    
    End If
    
    convByteToNChar = toRetStr
    
End Function

'
'��ϊ��ŕK�v�ȕ�����𓾂�
'
'!CAUTION!
'    fromRadix,toRadix��2~16�͈͓̔��ł��鎖�̓`�F�b�N���Ȃ�
'
Private Function convRadix(ByVal fromRadix As Byte, ByVal toRadix As Byte, ByRef endStatus As Variant) As String
    
    Dim radixTable As Variant
    
    '��ϊ��p�e�[�u��
    radixTable = Array( _
        Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""), _
        Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""), _
        Array("0", "1", "10", "11", "100", "101", "110", "111", "1000", "1001", "1010", "1011", "1100", "1101", "1110", "1111", "10000"), _
        Array("0", "1", "2", "10", "11", "12", "20", "21", "22", "100", "101", "102", "110", "111", "112", "120", "121"), _
        Array("0", "1", "2", "3", "10", "11", "12", "13", "20", "21", "22", "23", "30", "31", "32", "33", "100"), _
        Array("0", "1", "2", "3", "4", "10", "11", "12", "13", "14", "20", "21", "22", "23", "24", "30", "31"), _
        Array("0", "1", "2", "3", "4", "5", "10", "11", "12", "13", "14", "15", "20", "21", "22", "23", "24"), _
        Array("0", "1", "2", "3", "4", "5", "6", "10", "11", "12", "13", "14", "15", "16", "20", "21", "22"), _
        Array("0", "1", "2", "3", "4", "5", "6", "7", "10", "11", "12", "13", "14", "15", "16", "17", "20"), _
        Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "10", "11", "12", "13", "14", "15", "16", "17"), _
        Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"), _
        Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "10", "11", "12", "13", "14", "15"), _
        Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "10", "11", "12", "13", "14"), _
        Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "10", "11", "12", "13"), _
        Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "10", "11", "12"), _
        Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "10", "11"), _
        Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "10") _
    )
    
    convRadix = radixTable(fromRadix)(toRadix)
    
End Function

'
'convRadix���g����N�i��N�i�ϊ����g���C����
'�ϊ������̏ꍇ�́A�ϊ����N�i�l��Ԃ�
'���s�̏ꍇ�́AendStatus��#N/A!���i�[���A�󕶎���Ԃ�
'
'!CAUTION!
'    fromRadix,toRadix��2~16�͈͓̔��ł��鎖�̓`�F�b�N���Ȃ�
'
Private Function tryConvRadix(ByVal intStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByRef endStatus As Variant) As String
    
    Dim idxOfRTable As Byte
    Dim toRetStr As String
    
    'convRadix�ŉ����\���ǂ����`�F�b�N
    For idxOfRTable = 0 To 16
        If (intStr = convRadix(fromRadix, idxOfRTable, stsOfSub)) Then '��ϊ��p�e�[�u���ɉ���������
            toRetStr = convRadix(toRadix, idxOfRTable, stsOfSub) '��ϊ��e�[�u���������Ԃ�
            Exit For
            
        End If
        
    Next idxOfRTable
    
    If (idxOfRTable > 16) Then '������Ȃ������ꍇ
        endStatus = CVErr(xlErrNA)
        toRetStr = ""
    
    End If
    
    tryConvRadix = toRetStr
    
End Function

'
'�����̕s�v��"0"����菜��
'
'�ȉ��̏ꍇ�́A"0"��Ԃ�
'    �E�󕶎����w�肵���ꍇ
'    �E���ׂ�"0"(���K�\���ŕ\��"0+")�ȕ�����̏ꍇ
'
'!CAUTION!
'    intStr���L���Ȑ��l�����񂩂ǂ����̓`�F�b�N���Ȃ�
'
Private Function removeLeft0(ByVal intStr As String, ByRef endStatus As Variant) As String
    
    Dim lpIdx As Long
    Dim lpMx As Long
    Dim toRetStr As String
    
    lpMx = Len(intStr)
    lpIdx = 1
    
    '������{�����[�v
    Do While (lpIdx <= lpMx)
        
        If (Mid(intStr, lpIdx, 1) <> "0") Then '�{���Ώە�����"0"�łȂ�
            Exit Do
            
        End If
        
        lpIdx = lpIdx + 1 'increment
        
    Loop
    
    If (lpIdx > lpMx) Then '�󕶎� or ���ׂ�"0"�ȕ�����
        toRetStr = "0"
        
    Else
        toRetStr = Right(intStr, lpMx - lpIdx + 1)
        
    End If
    
    removeLeft0 = toRetStr
    
End Function

'
'�E���̕s�v��"0"����菜��
'
'�ȉ��̏ꍇ�́A"0"��Ԃ�
'    �E�󕶎����w�肵���ꍇ
'    �E���ׂ�"0"(���K�\���ŕ\��"0+")�ȕ�����̏ꍇ
'
'!CAUTION!
'    intStr���L���Ȑ��l�����񂩂ǂ����̓`�F�b�N���Ȃ�
'
Private Function removeRight0(ByVal intStr As String, ByRef endStatus As Variant) As String
    
    Dim lpIdx As Long
    Dim toRetStr As String
    
    lpIdx = Len(intStr)
    
    '������{�����[�v
    Do While (lpIdx > 0)
        
        If (Mid(intStr, lpIdx, 1) <> "0") Then '�{���Ώە�����"0"�łȂ�
            Exit Do
            
        End If
        
        lpIdx = lpIdx - 1 'decrement
        
    Loop
    
    If (lpIdx = 0) Then  '�󕶎� or ���ׂ�"0"�ȕ�����
        toRetStr = "0"
        
    Else
        toRetStr = Left(intStr, lpIdx)
        
    End If
    
    removeRight0 = toRetStr
    
End Function

'
'String�z��̏��Ԃ���ւ���
'
Private Function invertStringArray(ByRef srcArr() As String, ByRef endStatus As Variant) As String()
    
    Dim cnt As Long
    Dim cntMx As Long
    Dim idx As Long
    
    cntMx = UBound(srcArr)
    
    ReDim retArr(cntMx) As String
    
    idx = cntMx
    For cnt = 0 To cntMx
        retArr(cnt) = srcArr(idx)
        idx = idx - 1
    Next cnt
    
    invertStringArray = retArr
    
End Function


