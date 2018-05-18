Attribute VB_Name = "convNPntNPnt"
'<�萔>------------------------------------------------------------------------------------------

Private Const DOT As String = "." '�����_�\�L
'
'-----------------------------------------------------------------------------------------</�萔>

'<PrivateFunction�p�e�X�g�֐�>---------------------------------------------------------------------------------------------------------------------
'
Public Function TESTseparateToIntAndFrc(ByVal pntStr As String, ByVal radix As Byte, ByVal remove0 As Boolean) As Variant
    Dim intPrt As String
    Dim frcPrt As String
    Dim isMinus As Boolean
    TESTseparateToIntAndFrc = separateToIntAndFrc(pntStr, radix, remove0, intPrt, frcPrt, isMinus)
End Function

Public Function TESTseparateToIntAndFrcByRef1(ByVal pntStr As String, ByVal radix As Byte, ByVal remove0 As Boolean) As Variant
    Dim intPrt As String
    Dim frcPrt As String
    Dim isMinus As Boolean
    x = separateToIntAndFrc(pntStr, radix, remove0, intPrt, frcPrt, isMinus)
    TESTseparateToIntAndFrcByRef1 = intPrt
End Function

Public Function TESTseparateToIntAndFrcByRef2(ByVal pntStr As String, ByVal radix As Byte, ByVal remove0 As Boolean) As Variant
    Dim intPrt As String
    Dim frcPrt As String
    Dim isMinus As Boolean
    x = separateToIntAndFrc(pntStr, radix, remove0, intPrt, frcPrt, isMinus)
    TESTseparateToIntAndFrcByRef2 = frcPrt
End Function

Public Function TESTseparateToIntAndFrcByRef3(ByVal pntStr As String, ByVal radix As Byte, ByVal remove0 As Boolean) As Variant
    Dim intPrt As String
    Dim frcPrt As String
    Dim isMinus As Boolean
    x = separateToIntAndFrc(pntStr, radix, remove0, intPrt, frcPrt, isMinus)
    TESTseparateToIntAndFrcByRef3 = isMinus
End Function

Public Function TESTcheckNPntStr(ByVal pntStr As String, ByVal radix As Byte) As Variant
    Dim isMinus As Boolean
    Dim idxOfDot As Long
    Dim stsOfSub As Variant
    TESTcheckNPntStr = checkNPntStr(pntStr, radix, isMinus, idxOfDot, stsOfSub)
End Function

Public Function TESTcheckNPntStrByRef1(ByVal pntStr As String, ByVal radix As Byte) As Variant
    Dim isMinus As Boolean
    Dim idxOfDot As Long
    Dim stsOfSub As Variant
    x = checkNPntStr(pntStr, radix, isMinus, idxOfDot, stsOfSub)
    TESTcheckNPntStrByRef1 = isMinus
End Function

Public Function TESTcheckNPntStrByRef2(ByVal pntStr As String, ByVal radix As Byte) As Variant
    Dim isMinus As Boolean
    Dim idxOfDot As Long
    Dim stsOfSub As Variant
    x = checkNPntStr(pntStr, radix, isMinus, idxOfDot, stsOfSub)
    TESTcheckNPntStrByRef2 = idxOfDot
End Function

Public Function TESTcheckNPntStrByRef3(ByVal pntStr As String, ByVal radix As Byte) As Variant
    Dim isMinus As Boolean
    Dim idxOfDot As Long
    Dim stsOfSub As Variant
    x = checkNPntStr(pntStr, radix, isMinus, idxOfDot, stsOfSub)
    TESTcheckNPntStrByRef3 = stsOfSub
End Function

Public Function TESTadd(ByVal val1 As String, ByVal val2 As String, ByVal radix As Byte) As Variant
    TESTadd = add(val1, val2, radix)
End Function

Public Function TESTsubtract(ByVal val1 As String, ByVal val2 As String, ByVal radix As Byte) As Variant
    Dim stsOfSub As Boolean
    TESTsubtract = subtract(val1, val2, radix, stsOfSub)
End Function

Public Function TESTsubtractByRef1(ByVal val1 As String, ByVal val2 As String, ByVal radix As Byte) As Variant
    Dim stsOfSub As Boolean
    x = subtract(val1, val2, radix, stsOfSub)
    TESTsubtractByRef1 = stsOfSub
End Function

Public Function TESTmultiple(ByVal multiplicand As String, ByVal multiplier As String, ByVal radix As Byte) As Variant
    TESTmultiple = multiple(multiplicand, multiplier, radix)
End Function

Public Function TESTmultipleByOneDigit(ByVal multiplicand As String, ByVal multiplierCh As String, ByVal radix As Byte) As Variant
    TESTmultipleByOneDigit = multipleByOneDigit(multiplicand, multiplierCh, radix)
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
    TESTconvIntPrtOfNPntToIntPrtOfNPnt = convIntPrtOfNPntToIntPrtOfNPnt(intStr, fromRadix, toRadix)
End Function

Public Function TESTconvFrcPrtOfNPntToFrcPrtOfNPnt(ByVal frcStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByVal numOfDigits As Long) As Variant
    TESTconvFrcPrtOfNPntToFrcPrtOfNPnt = convFrcPrtOfNPntToFrcPrtOfNPnt(frcStr, fromRadix, toRadix, numOfDigits)
End Function
'
'--------------------------------------------------------------------------------------------------------------------</PrivateFunction�p�e�X�g�֐�>

'
'2�������Z����
'
'�������s���̏ꍇ�́A�ȉ��ɉ�����CvErr��ԋp����
'    �Eradix��2~16�ȊO���A���l���n�i�l�Ƃ��ĕs���̏ꍇ(�G���[�R�[�h��#NUM!)
'    �E���l�񂪋󕶎���Null�̏ꍇ(�G���[�R�[�h��#NULL!)
'
Public Function addNPntNPnt(ByVal val1 As String, ByVal val2 As String, Optional ByVal radix As Byte = 10) As Variant
    
    Dim intPrtOfVal1 As String
    Dim frcPrtOfVal1 As String
    Dim isMinusOfVal1 As Boolean
    Dim lenOfVal1FrcPrt As Long
    
    Dim intPrtOfVal2 As String
    Dim frcPrtOfVal2 As String
    Dim isMinusOfVal2 As Boolean
    Dim lenOfVal2FrcPrt As Long
    
    Dim stsOfSub As Variant
    Dim subtractionWasMinus As Boolean
    
    Dim tmpVal1 As Variant
    Dim tmpVal2 As Variant
    Dim tmpAns As Variant
    
    Dim signOfAns As String
    Dim intPrtOfAns As Variant
    Dim frcPrtOfAns As Variant
    
    'val1�̕�����`�F�b�N&�����A��������
    stsOfSub = separateToIntAndFrc(val1, radix, True, intPrtOfVal1, frcPrtOfVal1, isMinusOfVal1)
    If IsError(stsOfSub) Then 'val1��n�i�l�Ƃ��ĕs��
        addNPntNPnt = stsOfSub 'checkNPntStr�̃G���[�R�[�h��Ԃ�
        Exit Function
        
    End If
    
    'valw�̕�����`�F�b�N&�����A��������
    stsOfSub = separateToIntAndFrc(val2, radix, True, intPrtOfVal2, frcPrtOfVal2, isMinusOfVal2)
    If IsError(stsOfSub) Then 'val2��n�i�l�Ƃ��ĕs��
        addNPntNPnt = stsOfSub 'checkNPntStr�̃G���[�R�[�h��Ԃ�
        Exit Function
        
    End If
    
    '�������̌������킹
    lenOfVal1FrcPrt = Len(frcPrtOfVal1)
    lenOfVal2FrcPrt = Len(frcPrtOfVal2)
    If (lenOfVal1FrcPrt > lenOfVal2FrcPrt) Then 'val1�̌������傫��
        frcPrtOfVal2 = frcPrtOfVal2 & String(lenOfVal1FrcPrt - lenOfVal2FrcPrt, "0") 'val2�̉E����0����
        lenOfVal2FrcPrt = Len(frcPrtOfVal2)
        
    Else 'val2�̌������傫��
        frcPrtOfVal1 = frcPrtOfVal1 & String(lenOfVal2FrcPrt - lenOfVal1FrcPrt, "0") 'val1�̉E����0����
        lenOfVal1FrcPrt = Len(frcPrtOfVal1)
        
    End If
    
    tmpVal1 = intPrtOfVal1 & frcPrtOfVal1
    tmpVal2 = intPrtOfVal2 & frcPrtOfVal2
    
    '���Zor���Z
    If (isMinusOfVal1) Then 'val1�̓}�C�i�X�l
        If (isMinusOfVal2) Then 'val2�̓}�C�i�X�l
            tmpAns = add(tmpVal1, tmpVal2, radix)
            signOfAns = "-"
            
        Else 'val2�̓v���X�l
            tmpAns = subtract(tmpVal1, tmpVal2, radix, subtractionWasMinus)
            If (subtractionWasMinus) Then
                signOfAns = ""
            Else
                signOfAns = "-"
            End If
        
        End If
        
    Else 'val1�̓v���X�l
        If (isMinusOfVal2) Then 'val2�̓}�C�i�X�l
            tmpAns = subtract(tmpVal1, tmpVal2, radix, subtractionWasMinus)
            If (subtractionWasMinus) Then
                signOfAns = "-"
            Else
                signOfAns = ""
            End If
            
        Else 'val2�̓v���X�l
            tmpAns = add(tmpVal1, tmpVal2, radix)
            signOfAns = ""
        
        End If
    
    End If
    
    '�����_����
    intPrtOfAns = Left(tmpAns, Len(tmpAns) - lenOfVal1FrcPrt)
    frcPrtOfAns = Right(tmpAns, lenOfVal1FrcPrt)
    
    '�s�v��"0"���폜
    intPrtOfAns = removeLeft0(intPrtOfAns)
    If (frcPrtOfAns <> "") Then
        frcPrtOfAns = removeRight0(frcPrtOfAns)
        If (frcPrtOfAns = "0") Then
            frcPrtOfAns = ""
        End If
    End If
    
    '-0�m�F
    If ((intPrtOfAns & frcPrtOfAns) = "0") Then
        signOfAns = ""
    End If
    
    addNPntNPnt = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)

End Function

'
'���l��n�i���l�񂩂ǂ����`�F�b�N���āA
'�������Ə������ɕ�������
'�������̋L�ڂ��Ȃ��ꍇ�́A�������͋󕶎����i�[����
'
'�����̏ꍇ��0��ԋp����
'
'���s�̏ꍇ�͈ȉ��ɉ�����CvErr��ԋp����
'    �@�Eradix��2~16�ȊO���A���l���n�i�l�Ƃ��ĕs���̏ꍇ(�G���[�R�[�h��#NUM!)
'    �@�E���l�񂪋󕶎���Null�̏ꍇ(�G���[�R�[�h��#NULL!)
'
'radix
'    �(2~16�̂�)
'
'remove0
'    �s�v��0(�������͍�����0�A�������͉E����0)����菜�����ǂ���
'    TRUE���w�肵�ď��������S��0�̏ꍇ�A�������͋󕶎����i�[����
'
Private Function separateToIntAndFrc(ByVal pntStr As String, ByVal radix As Byte, ByVal remove0 As Boolean, ByRef intPrt As String, ByRef frcPrt, ByRef isMinus As Boolean) As Variant
    
    Dim retOfCheckNPntStr As Long
    Dim idxOfDot As Long
    Dim stsOfSub As Variant
    Dim lenOfPntStr As Long
    Dim toRetIsMinus As Boolean
    Dim toRetIntPrt As String
    Dim toRetFrcPrt As String
    
    lenOfPntStr = Len(pntStr)
    
    'n�i�l�Ƃ��Đ��������`�F�b�N&��������&�����_�ʒu�擾
    retOfCheckNPntStr = checkNPntStr(pntStr, radix, toRetIsMinus, idxOfDot, stsOfSub)
    If IsError(stsOfSub) Then 'n�i�l�Ƃ��ĕs��
        separateToIntAndFrc = stsOfSub 'checkNPntStr�̃G���[�R�[�h��Ԃ�
        Exit Function
        
    End If
    
    '���������o�J�n�ʒu�̔���
    If (toRetIsMinus) Then '(-)�l�̏ꍇ
        stIdxOfIntPrt = 2
    Else
        stIdxOfIntPrt = 1
    End If
    
    '���o
    If (idxOfDot = 0) Then '�������̋L�ڂ��Ȃ��ꍇ
        toRetIntPrt = Mid(pntStr, stIdxOfIntPrt, (lenOfPntStr - stIdxOfIntPrt) + 1)
        toRetFrcPrt = ""
        
    Else '����������
        toRetIntPrt = Mid(pntStr, stIdxOfIntPrt, idxOfDot - stIdxOfIntPrt)
        toRetFrcPrt = Right(pntStr, lenOfPntStr - idxOfDot)
        
    End If
    
    '0�폜
    If (remove0) Then
        toRetIntPrt = removeLeft0(toRetIntPrt)
        
        If (toRetFrcPrt <> "") Then
            
            toRetFrcPrt = removeRight0(toRetFrcPrt)
                
            If (toRetFrcPrt = "0") Then '���ׂ�0��������
                toRetFrcPrt = ""
            End If
        End If
        
    End If
    
    '�ԋp
    intPrt = toRetIntPrt
    frcPrt = toRetFrcPrt
    isMinus = toRetIsMinus
    separateToIntAndFrc = 0
    
End Function


'
'���l��n�i���l�񂩂ǂ����`�F�b�N����
'
'�ԋp�l
'    n�i�l�����񂾂����̏ꍇ��errCode��0���i�[���A������ + 1��ԋp����
'    �����łȂ��ꍇ�́AerrCode��#NUM!���i�[���A
'    �ŏ��Ɍ�������10�i�����ȊO�̕����ʒu��ԋp����
'
'    �ȉ��̏ꍇ�́AerrCode�ɃG���[�R�[�h���i�[���A0��ԋp����
'    �@�Eradix��2~16�ȊO�̏ꍇ(�G���[�R�[�h��#NUM!)
'    �@�E�������󕶎���Null�̏ꍇ(�G���[�R�[�h��#NULL!)
'
'radix
'    �(2~16�̂�)
'
'idxOfDot(ByRef)
'    �����_�����ʒu
'    �����_�����������ꍇ0
'
Private Function checkNPntStr(ByVal pntStr As String, ByVal radix As Byte, ByRef isMinus As Boolean, ByRef idxOfDot As Long, ByRef errCode As Variant) As Long
    
    Dim minOkChar1 As Integer
    Dim maxOkChar1 As Integer
    Dim minOkChar2 As Integer
    Dim maxOkChar2 As Integer
    Dim radixIsBiggerThan10 As Boolean
    Dim cnt As Long
    Dim lpMx As Long
    Dim stCnt As Long
    Dim foundIdxOfDot As Long '�����_�������ŏ��Ɍ������������ʒu
    Dim ngIdx As Long
    Dim numOfDigits As Long
    
    '�����`�F�b�N
    If (radix < 2) Or (16 < radix) Then
        errCode = CVErr(xlErrNum) '#NUM!���i�[����
        checkNPntStr = 0
        Exit Function
        
    End If
    
    lpMx = Len(pntStr)
    
    If (lpMx = 0) Then
        errCode = CVErr(xlErrNull) '#NULL!���i�[����
        checkNPntStr = 0
        Exit Function
        
    End If
    
    '�����OK�ȕ����R�[�h�͈͂����
    minOkChar1 = Asc("0")
    If (radix <= 10) Then
        maxOkChar1 = Asc(CStr(radix - 1))
        radixIsBiggerThan10 = False
        
    Else
        maxOkChar1 = Asc("9")
        minOkChar2 = Asc("A")
        maxOkChar2 = Asc("A") + (radix - 11)
        
        radixIsBiggerThan10 = True
    
    End If
    
    '�������݃`�F�b�N
    If (Left(pntStr, 1) = "-") Then '������(-)
        isMinus = True
        stCnt = 2
        
    Else '������(+)
        isMinus = False
        stCnt = 1
        
    End If
    
    '�����񌟍����[�v
    foundIdxOfDot = 0
    ngIdx = 0
    numOfDigits = 0
    For cnt = stCnt To lpMx
        
        ch = Mid(pntStr, cnt, 1)
        chCode = Asc(ch)
        
        If (chCode < minOkChar1) Or (maxOkChar1 < chCode) Then  '������0~9������ł��Ȃ�
            If IIf(radixIsBiggerThan10, (chCode < minOkChar2) Or (maxOkChar2 < chCode), True) Then '������A~F������ł��Ȃ�
                
                If (ch = DOT) Then '�����_�����̏ꍇ
                    If (foundIdxOfDot = 0) Then '�����_�����̏o����1���
                    
                        If (numOfDigits = 0) Then '�������̌�����0
                            ngIdx = cnt
                            Exit For
                            
                        End If
                        
                        foundIdxOfDot = cnt
                        numOfDigits = 0
                    
                    Else '�����_�����̏o����2���
                        ngIdx = cnt
                        Exit For
                        
                    End If
                
                Else '�����͐��l�����ł��Ȃ��A�����_�����ł��Ȃ�
                    ngIdx = cnt
                    Exit For
                    
                End If
                
            Else '������A~F
                numOfDigits = numOfDigits + 1 'increment
            End If
            
        Else '������0~9
            numOfDigits = numOfDigits + 1 'increment
        End If
        
    Next cnt
    
    If (numOfDigits = 0) And (ngIdx = 0) Then '���l��������Ȃ��ꍇ
        ngIdx = cnt - 1
        
    End If
    
    If (ngIdx > 0) Then 'NG���������݂���ꍇ
        errCode = CVErr(xlErrNum) '#NUM!���i�[����
        checkNPntStr = ngIdx 'NG�����ʒu��ԋp
        
    Else '���ׂ�OK�ȏꍇ
    
        idxOfDot = foundIdxOfDot
        errCode = 0
        checkNPntStr = cnt '������ + 1��ԋp
        
    End If
    
End Function

'
'2����a�Z����
'
'!CAUTION!
'    val1, val2 ���L����n�i�l�ł��邩�̓`�F�b�N���Ȃ�
'    radix��2~16�͈͓̔��ł��鎖�̓`�F�b�N���Ȃ�
'
Private Function add(ByVal val1 As String, ByVal val2 As String, ByVal radix As Byte) As String
    
    '�ϐ��錾
    Dim lenOfVal1 As Long
    Dim lenOfVal2 As Long
    Dim idxOfVal As Long
    Dim stringBuilder() As String
    Dim decDigitOfVal1 As Integer
    Dim decDigitOfVal2 As Integer
    Dim decCarrier As Integer
    Dim decDigitOfAns  As Integer
    Dim stsOfSub As Variant
    
    '���l�񒷎擾
    lenOfVal1 = Len(val1)
    lenOfVal2 = Len(val2)
    
    '0���ߊm�F
    If (lenOfVal1 > lenOfVal2) Then
        val2 = String(lenOfVal1 - lenOfVal2, "0") & val2
        idxOfVal = lenOfVal1
        
    Else
        val1 = String(lenOfVal2 - lenOfVal1, "0") & val1
        idxOfVal = lenOfVal2
        
    End If
    
    '���[�v�O������
    ReDim stringBuilder(idxOfVal) '�̈�m��
    decCarrier = 0
    
    '���̐������[�v
    Do While (idxOfVal > 0)
        
        '�Ώی��̘a�Z
        decDigitOfVal1 = convNCharToByte(Mid(val1, idxOfVal, 1))
        decDigitOfVal2 = convNCharToByte(Mid(val2, idxOfVal, 1))
        decDigitOfAns = decDigitOfVal1 + decDigitOfVal2 + decCarrier
        
        '�J��オ��`�F�b�N
        If (decDigitOfAns >= radix) Then '�J��オ�肠��
            decCarrier = 1
            decDigitOfAns = decDigitOfAns - radix
            
        Else '�J��オ��Ȃ�
            decCarrier = 0
            
        End If
        
        stringBuilder(idxOfVal) = convByteToNChar(decDigitOfAns) '�����i�[
        
        idxOfVal = idxOfVal - 1 'decrement
        
    Loop
    
    '�ŏ�ʌ��i�[
    stringBuilder(idxOfVal) = IIf(decCarrier > 0, "1", "")
    
    add = Join(stringBuilder, vbNullString)
    
End Function

'
'val1����val2�����Z����
'���Z���ʂ�(-)�l�̏ꍇ�́AresultIsMinus��TRUE���i�[����
'���Z���ʂ�(+)�l�̏ꍇ�́AresultIsMinus��FALSE���i�[����
'
'!CAUTION!
'    val1, val2 ���L����n�i�l�ł��邩�̓`�F�b�N���Ȃ�
'    radix��2~16�͈͓̔��ł��鎖�̓`�F�b�N���Ȃ�
'
Private Function subtract(ByVal val1 As String, ByVal val2 As String, ByVal radix As Byte, ByRef resultIsMinus As Boolean) As String
    
    '�ϐ��錾
    Dim idxMxOfVal As Long
    Dim diffIdx As Long
    Dim val1IsLarger As Integer '0:�s��, 1:yes, -1:no
    Dim lenOfVal1 As Long
    Dim lenOfVal2 As Long
    Dim idxOfVal As Long
    Dim stringBuilder() As String
    Dim decDigitOfVal1 As Integer
    Dim decDigitOfVal2 As Integer
    Dim decCarrier As Integer
    Dim decDigitOfAns  As Integer
    
    '���l�񒷎擾
    lenOfVal1 = Len(val1)
    lenOfVal2 = Len(val2)
    
    '0���ߊm�F
    If (lenOfVal1 > lenOfVal2) Then
        val2 = String(lenOfVal1 - lenOfVal2, "0") & val2
        idxOfVal = lenOfVal1
        
    Else
        val1 = String(lenOfVal2 - lenOfVal1, "0") & val1
        idxOfVal = lenOfVal2
        
    End If
    
    '<�召��r�`�F�b�N>--------------------------------------------------------------------
    
    diffIdx = 1
    val1IsLarger = 0
    Do While (diffIdx <= idxOfVal)
        
        decDigitOfVal1 = convNCharToByte(Mid(val1, diffIdx, 1))
        decDigitOfVal2 = convNCharToByte(Mid(val2, diffIdx, 1))
        
        '�ǂ��炩���傫�������� break
        If decDigitOfVal1 > decDigitOfVal2 Then
            val1IsLarger = 1
            Exit Do
        
        ElseIf decDigitOfVal1 < decDigitOfVal2 Then
            val1IsLarger = -1
            Exit Do
        
        End If
        
        diffIdx = diffIdx + 1
        
    Loop
    
    'val1��val2���͓������l�̏ꍇ
    If (val1IsLarger = 0) Then
        resultIsMinus = False
        subtract = String(idxOfVal, "0") '0��ԋp
        Exit Function
        
    End If
    
    
    If (val1IsLarger = 1) Then 'val1�̕����傫�����l�̏ꍇ
        resultIsMinus = False '(+)���i�[
        
    Else 'val2�̕����傫�����l�̏ꍇ
        
        '2�������ւ���
        buf = val1
        val1 = val2
        val2 = buf
        
        resultIsMinus = True '(-)���i�[
        
    End If
    
    '-------------------------------------------------------------------</�召��r�`�F�b�N>
    
    '���[�v�O������
    ReDim stringBuilder(idxOfVal) '�̈�m��
    decCarrier = 0
    
    '���̐������[�v
    Do While (idxOfVal > 0)
        
        '�Ώی��̌��Z
        decDigitOfVal1 = convNCharToByte(Mid(val1, idxOfVal, 1))
        decDigitOfVal2 = convNCharToByte(Mid(val2, idxOfVal, 1))
        
        '�J�艺����`�F�b�N
        If (decDigitOfVal1 = 0) And (decCarrier = -1) Then
            decCarrier = -1
            decDigitOfVal1 = radix - 1
            
        Else
            decDigitOfVal1 = decDigitOfVal1 + decCarrier
            decCarrier = 0
            
            If (decDigitOfVal1 < decDigitOfVal2) Then
                decDigitOfVal1 = radix + decDigitOfVal1
                decCarrier = -1
                
            End If
            
        End If
        
        decDigitOfAns = decDigitOfVal1 - decDigitOfVal2
        
        stringBuilder(idxOfVal) = convByteToNChar(decDigitOfAns) '�����i�[
        
        idxOfVal = idxOfVal - 1 'decrement
        
    Loop
    
    '�ŏ�ʌ��i�[
    stringBuilder(idxOfVal) = ""
    
    subtract = Join(stringBuilder, vbNullString)
    
    
End Function

'
'��Z������
'
'!CAUTION!
'    multiplicand, multiplier ���L����n�i�l�ł��邩�̓`�F�b�N���Ȃ�
'    radix��2~16�͈͓̔��ł��鎖�̓`�F�b�N���Ȃ�
'
Private Function multiple(ByVal multiplicand As String, ByVal multiplier As String, ByVal radix As Byte) As String

    Dim ansOfMultipleByOneDigit As String
    Dim numOfShift As Long
    Dim tmpAns As String
    Dim stsOfSub As Variant
    Dim idxOfMultiplier As Long
    
    'multiplier�̕s�v��0����菜��
    multiplier = removeLeft0(multiplier)
    
    numOfShift = 0
    tmpAns = String(Len(multiplicand), "0")
    
    '��Z���[�v
    For idxOfMultiplier = Len(multiplier) To 1 Step -1
        
        digitOfMultiplier = Mid(multiplier, idxOfMultiplier, 1)
        
        If (digitOfMultiplier <> "0") Then '1�ȏ�̐��l�̎������A���ɑ������킹��
            ansOfMultipleByOneDigit = multipleByOneDigit(multiplicand, digitOfMultiplier, radix)
            tmpAns = add(tmpAns, ansOfMultipleByOneDigit & String(numOfShift, "0"), radix)
            
        End If
        
        numOfShift = numOfShift + 1
        
    Next idxOfMultiplier
    
    multiple = tmpAns
    
End Function

'
'1�����l�ɂ���Z������
'
'!CAUTION!
'    multiplicand, multiplierCh ���L����n�i�l�ł��邩�̓`�F�b�N���Ȃ�
'    radix��2~16�͈͓̔��ł��鎖�̓`�F�b�N���Ȃ�
'
Private Function multipleByOneDigit(ByVal multiplicand As String, ByVal multiplierCh As String, ByVal radix As Byte) As String

    Dim decMultiplier As Byte
    Dim stsOfSub As Variant
    
    Dim decDigitOfMultiplicand As Byte
    Dim decCarrier As Byte
    Dim decDigitOfAns  As Byte
    
    Dim digitIdxOfMultiplicand As Long
    Dim stringBuilder() As String '����Z���ʊi�[�p
    
    '�搔��10�i�ϊ�
    decMultiplier = convNCharToByte(multiplierCh)
    
    '0�|��&1�|���`�F�b�N
    If (decMultiplier = 0) Then
        multipleByOneDigit = String(Len(multiplicand), "0") '0�|���̏ꍇ��0��Ԃ�
        Exit Function
    
    ElseIf (multiplierCh = "1") Then '1�|���̏ꍇ�͂��̂܂ܕԂ�
        multipleByOneDigit = multiplicand
        Exit Function
        
    End If
    
    '���[�v�O������
    digitIdxOfMultiplicand = Len(multiplicand)
    ReDim stringBuilder(digitIdxOfMultiplicand) '�̈�m��
    decCarrier = 0
    
    Do While (digitIdxOfMultiplicand > 0) '��搔���c���Ă����
        
        '�Ώی��̏�Z
        decDigitOfMultiplicand = convNCharToByte(Mid(multiplicand, digitIdxOfMultiplicand, 1))
        decDigitOfAns = decDigitOfMultiplicand * decMultiplier + decCarrier
        
        digitOfAns = convIntPrtOfNPntToIntPrtOfNPnt(decDigitOfAns, 10, radix) '10�i��n�i�ϊ�
        
        '�J��オ��&���i�[
        If (Len(digitOfAns) = 2) Then '�J��オ�肠��
            decCarrier = convNCharToByte(Left(digitOfAns, 1))
            digitOfAns = Right(digitOfAns, 1)
            
        Else '�J��オ��Ȃ�
            decCarrier = 0
            
        End If
        
        '���i�[
        stringBuilder(digitIdxOfMultiplicand) = digitOfAns
        
        digitIdxOfMultiplicand = digitIdxOfMultiplicand - 1 'decrement
        
    Loop
    
    '�ŏ�ʌ��i�[
    stringBuilder(digitIdxOfMultiplicand) = IIf(decCarrier > 0, convByteToNChar(decCarrier), "")
    
    multipleByOneDigit = Join(stringBuilder, vbNullString) '������A��
    
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
    divisor = removeLeft0(divisor)
    
    '<divisor��10�i�ϊ�>------------------------------------------------------------
    
    tmp = convIntPrtOfNPntToIntPrtOfNPnt(divisor, radix, 10)
    
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
        digitOfDividend = rmnd * radix + convNCharToByte(Mid(dividend, digitIdxOfDividend, 1)) '��ʌ��̗]�� & �Y����
        
        quot = digitOfDividend \ divisorDec '��
        rmnd = digitOfDividend Mod divisorDec '�]��
        
        ReDim Preserve stringBuilder(digitIdxOfDividend - 1) '�̈�g��
        
        '����ǋL
        '���錅�ɑ΂��鏜�Z�̏��́A�����������蕿�Ȃ�
        stringBuilder(digitIdxOfDividend - 1) = convByteToNChar(quot)
        
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
        
        remainder = convIntPrtOfNPntToIntPrtOfNPnt(rmnd, 10, radix)
    
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
Private Function convIntPrtOfNPntToIntPrtOfNPnt(ByVal intStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte) As String
    
    Dim stsOfSub As Variant
    Dim retOfTryConvRadix As String
    Dim strLenOfToRadix As Long
    Dim stringBuilder() As String '�ϊ��㕶���񐶐��p
    Dim sizeOfStringBuilder As Long
    Dim rm As String
    
    intStr = removeLeft0(intStr)
    
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
    chOfToRadix = convRadix(fromRadix, toRadix)
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
        
        intStr = removeLeft0(intStr) '�����̕s�v��"0"����菜��
        
        ReDim Preserve stringBuilder(sizeOfStringBuilder) '�̈�g��
        
        '��]��(toRadix)�i�l�ɕϊ��������ʂ��Z�oDigit
        stringBuilder(sizeOfStringBuilder) = convIntPrtOfNPntToIntPrtOfNPnt(rm, fromRadix, toRadix)
        
        sizeOfStringBuilder = sizeOfStringBuilder + 1
        
    Loop
    
    ReDim Preserve stringBuilder(sizeOfStringBuilder) '�̈�g��
    
    '�ŏ��Bit�t��
    '��]��(toRadix)�i�l�ɕϊ��������ʂ��Z�oDigit
    stringBuilder(sizeOfStringBuilder) = convIntPrtOfNPntToIntPrtOfNPnt(intStr, fromRadix, toRadix)
    
    convIntPrtOfNPntToIntPrtOfNPnt = Join(invertStringArray(stringBuilder), vbNullString) '������A��
    
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
Private Function convFrcPrtOfNPntToFrcPrtOfNPnt(ByVal frcStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByVal numOfDigits As Long) As String
    
    Dim stsOfSub As Variant
    Dim stringBuilder() As String '�ϊ��㕶���񐶐��p
    Dim sizeOfStringBuilder As Long
    Dim retOfMultiple As String
    
    frcStr = removeRight0(frcStr)
    
    '�ϊ������������ꍇ
    If (fromRadix = toRadix) Then
        
        If (numOfDigits > 0) Then
            convFrcPrtOfNPntToFrcPrtOfNPnt = frcStr '"0"����菜���������̒l��Ԃ�
        Else
            convFrcPrtOfNPntToFrcPrtOfNPnt = "" '�󕶎���Ԃ�
        End If
        
        Exit Function
    End If
    
    '"0"��ϊ�����ꍇ
    If (frcStr = "0") Then
        
        If (numOfDigits > 0) Then
            convFrcPrtOfNPntToFrcPrtOfNPnt = "0" '"0"��Ԃ�
        Else
            convFrcPrtOfNPntToFrcPrtOfNPnt = "" '�󕶎���Ԃ�
        End If
        
        Exit Function
    End If
    
    '�������[�v�O������
    strOfToRadix = convRadix(fromRadix, toRadix)
    sizeOfStringBuilder = 0
    lenOfFrcStrB = Len(frcStr)
    
    '�������[�v - toRadix�ɂ���Z�ɂ���ĉ������߂� -
    Do While (sizeOfStringBuilder < numOfDigits)
        
        '�����̐ς�0�ɂȂ�����I��
        If (frcStr = "0") Then
            Exit Do
            
        End If
        
        frcStr = multiple(frcStr, strOfToRadix, fromRadix)
        
        ReDim Preserve stringBuilder(sizeOfStringBuilder) '�̈�g��
        
        '�����������E��
        lenOfFrcStrA = Len(frcStr)
        
        If (lenOfFrcStrA > lenOfFrcStrB) Then
            tmp = Left(frcStr, lenOfFrcStrA - lenOfFrcStrB)
            frcStr = Right(frcStr, lenOfFrcStrB)
            increasedDigits = convIntPrtOfNPntToIntPrtOfNPnt(tmp, fromRadix, toRadix)
            
        Else
            increasedDigits = "0"
        
        End If
        
        stringBuilder(sizeOfStringBuilder) = increasedDigits '����ǋL
        
        frcStr = removeRight0(frcStr) ' �E���̕s�v��0����菜��
        
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
Private Function convNCharToByte(ByVal ch As String) As Byte
    
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
Private Function convByteToNChar(ByVal byt As Byte) As String
    
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
Private Function convRadix(ByVal fromRadix As Byte, ByVal toRadix As Byte) As String
    
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
        If (intStr = convRadix(fromRadix, idxOfRTable)) Then '��ϊ��p�e�[�u���ɉ���������
            toRetStr = convRadix(toRadix, idxOfRTable) '��ϊ��e�[�u���������Ԃ�
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
Private Function removeLeft0(ByVal intStr As String) As String
    
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
Private Function removeRight0(ByVal intStr As String) As String
    
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
Private Function invertStringArray(ByRef srcArr() As String) As String()
    
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

