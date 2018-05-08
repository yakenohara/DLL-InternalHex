Attribute VB_Name = "NumericString"
'wish
'n�i���v�Z���ł���悤�ɂ�����

Public Const DOT As String = "." '�����_�\�L

'����؂�Ȃ����l�ɑ΂��ĉ��񊄂�Z���邩
Const DEFAULT_DIV_TIMES_FOR_INDIVISIBLE As Long = 255

'2�i���������̎Z�o����
Const DEFAULT_FRC_DIGITS As Long = 255

'
'n�i��������w�茅�����V�t�g����
'
Public Function shiftNumeralPnt(ByVal str As String, ByVal shift As Long, Optional radix As Byte = 10) As Variant
    
    Dim idxOfDot As Long
    Dim ret As Long
    Dim toRet As Variant
    Dim sign As String
    
    Dim intPt As String
    Dim prcPt As String
    
    '�����`�F�b�N
    If str = "" Then '�󕶎��̏ꍇ
        shiftNumeralPnt = CVErr(xlErrValue) '#VALUE!��Ԃ�
        Exit Function
        
    End If
    
    '��������菜��
    If Left(str, 1) = "-" Then '(-)�l�̎�
        str = Right(str, Len(str) - 1)
        sign = "-"
        
    Else
        sign = ""
    
    End If
    
    '�`�F�b�N
    ret = checkNumeralPntStr(str, idxOfDot, radix)
    
    If (ret = (Len(str) + 1)) Then 'n�i�����񂾂����ꍇ
        
        If (idxOfDot <= Len(str)) Then '�����񒆂ɏ����_�������������ꍇ
            str = Replace(str, DOT, "")
            
        End If
        
        shift = idxOfDot + shift
        lenOfStr = Len(str)
        
        If (shift <= 1) Then '������0���߂�����
            intPt = "0"
            frcPt = String(1 - shift, "0") & str
            
        ElseIf (shift > (lenOfStr + 1)) Then '�E����0���߂���
            intPt = str & String(shift - lenOfStr - 1, "0")
            frcPt = ""
            
        ElseIf (shift = (lenOfStr + 1)) Then '�����_�ʒu���L�ڕs�v�̏ꍇ
            intPt = str
            frcPt = ""
            
        Else '�����񒆂ɏ����_��}������
            intPt = Left(str, shift - 1)
            frcPt = Right(str, lenOfStr - Len(intPt))
            
        End If
        
        '�������̕s�v��"0"����菜��
        If (intPt <> "0") Then
            Do While Left(intPt, 1) = "0"
                intPt = Right(intPt, Len(intPt) - 1)
                
            Loop
        End If
        
        '�������̕s�v��"0"����菜��
        Do While Right(frcPt, 1) = "0"
            frcPt = Left(frcPt, Len(frcPt) - 1)
            
        Loop
        
        toRet = sign & intPt & IIf(frcPt = "", "", DOT & frcPt)
    
    Else 'n�i������Ŗ��������ꍇ
        toRet = CVErr(xlErrNum) '#NUM!��Ԃ�
    
    End If
    
    shiftNumeralPnt = toRet
    
End Function

'
'������n�i�����񂩂ǂ�����Ԃ�
'
'radix
'    �(2~16�܂�)
'
Public Function isNumeralPnt(ByVal decStr As String, Optional radix As Byte = 10) As Boolean
    
    Dim idxOfDot As Long
    Dim ret As Long
    Dim toRet As Boolean
    
    '�`�F�b�N
    ret = checkNumeralPntStr(decStr, idxOfDot, radix)
    
    If (ret = (Len(decStr) + 1)) Then 'n�i�����񂾂����ꍇ
        toRet = True
        
    Else
        toRet = False
    
    End If
    
    isNumeralPnt = toRet
    
    
End Function

'
'�����t��10�i�l���珬���t��2�i�l�ɕϊ�����
'
'numOfDigits
'    �����Z�o���̌��E���Z��
'
Public Function convDecPntToBinPnt(ByVal decStr As String, Optional numOfDigits As Long = DEFAULT_FRC_DIGITS) As Variant
    
    convDecPntToBinPnt = convNPntToNPnt(decStr, 10, numOfDigits)
    
End Function

'
'�����t��2�i�l���珬���t��10�i�l�ɕϊ�����
'
Public Function convBinPntToDecPnt(ByVal binStr As String) As Variant
    
    convBinPntToDecPnt = convNPntToNPnt(binStr, 2, -1)
    
End Function

'
'10�i��2�i�ϊ��p���ʊ֐�
'
Private Function convNPntToNPnt(ByVal pntStr As String, ByVal radix As Byte, numOfDigits As Long) As Variant

    Dim intPtOfBefore As String '������
    Dim frcPtOfBefore As String '������
    
    Dim intPtOfAfter As String '������
    Dim frcPtOfAfter As String '������
    Dim sign As String '����
    
    Dim ret As Long
    Dim idxOfDot As Long
    
    '�����`�F�b�N
    If (Len(pntStr) = 0) Then
        convNPntToNPnt = CVErr(xlErrValue) '#VALUE!��ԋp
        Exit Function
        
    End If
    
    '�}�C�i�X�l�`�F�b�N
    If (Left(pntStr, 1) = "-") Then
        If (Len(pntStr) < 2) Then
            convNPntToNPnt = CVErr(xlErrNum) '#NUM!��ԋp
            Exit Function
        End If
        
        sign = "-"
        pntStr = Right(pntStr, Len(pntStr) - 1)
        
    Else
        sign = ""
        
    End If
    
    '���l�񂩂ǂ����`�F�b�N
    ret = checkNumeralPntStr(pntStr, idxOfDot, radix)
    
    If (ret <> (Len(pntStr) + 1)) Then '�������n�i�l�ł͂Ȃ�
        convNPntToNPnt = CVErr(xlErrNum) '#NUM!��ԋp
        Exit Function
        
    End If
    
    '������&����������
    intPtOfBefore = Left(pntStr, idxOfDot - 1)
    
    '��������2�i�ϊ�
    If (radix = 2) Then
        intPtOfAfter = convIntPrtOfBinToIntPrtOfDecPrt(intPtOfBefore) '2�i�����ϊ��Ώۂ̎�
        
    Else
        intPtOfAfter = convIntPrtOfDecPntToIntPrtOfBinPnt(intPtOfBefore) '10�i�����ϊ��Ώۂ̎�
        
    End If
    
    If (idxOfDot = ret) Then '�������͑��݂��Ȃ��ꍇ
        frcPtOfBefore = ""
        frcPtOfAfter = ""
        
    Else '�����������݂���ꍇ
        frcPtOfBefore = Right(pntStr, Len(pntStr) - idxOfDot)
        
        '��������2�i�ϊ�
        If (radix = 2) Then
            frcPtOfAfter = convFrcPrtOfBinPntToFrcPrtOfDecPnt(frcPtOfBefore) '2�i�����ϊ��Ώۂ̎�
        Else
            frcPtOfAfter = convFrcPrtOfDecPntToFrcPrtOfBinPnt(frcPtOfBefore, numOfDigits) '10�i�����ϊ��Ώۂ̎�
        End If
        
        If frcPtOfAfter <> "" Then '�����_�t�����K�v�ȏꍇ
            frcPtOfAfter = DOT & frcPtOfAfter
            
        End If
        
    End If
    
    '�����񌋍�
    convNPntToNPnt = sign & intPtOfAfter & frcPtOfAfter

End Function

'
'2�������Z����
'
Public Function addDecPntDecPnt(ByVal value1 As String, ByVal value2 As String) As Variant
    
    Dim intPrtOfVal1 As String
    Dim frcPrtOfVal1 As String
    Dim isMinusOfVal1 As Boolean
    
    Dim intPrtOfVal2 As String
    Dim frcPrtOfVal2 As String
    Dim isMinusOfVal2 As Boolean
    
    Dim retOfSeparatePnt As Variant
    
    Dim toRetSign As String
    
    Dim substitutionWasMinus As Boolean
    
    'val1�̕�����`�F�b�N&�����A��������
    retOfSeparatePnt = separateToIntAndFrc(value1, intPrtOfVal1, frcPrtOfVal1, isMinusOfVal1, 10)
    
    If (retOfSeparatePnt <> 0) Then 'val1��10�i�l�Ƃ��ĕs��
        addDecPntDecPnt = addDecPntDecPnt
        Exit Function
        
    End If
    
    'valw�̕�����`�F�b�N&�����A��������
    retOfSeparatePnt = separateToIntAndFrc(value2, intPrtOfVal2, frcPrtOfVal2, isMinusOfVal2, 10)
    
    If (retOfSeparatePnt <> 0) Then 'val2��10�i�l�Ƃ��ĕs��
        addDecPntDecPnt = addDecPntDecPnt
        Exit Function
        
    End If
    
    '�������̌������킹
    lenOfVal1FrcPrt = Len(frcPrtOfVal1)
    lenOfVal2FrcPrt = Len(frcPrtOfVal2)
    If (lenOfVal1FrcPrt > lenOfVal2FrcPrt) Then
        frcPrtOfVal2 = frcPrtOfVal2 & String(lenOfVal1FrcPrt - lenOfVal2FrcPrt, "0")
        
    Else
        frcPrtOfVal1 = frcPrtOfVal1 & String(lenOfVal2FrcPrt - lenOfVal1FrcPrt, "0")
        
    End If
    
    tmpVal1 = intPrtOfVal1 & frcPrtOfVal1
    tmpVal2 = intPrtOfVal2 & frcPrtOfVal2
    
    '���Zor���Z
    If (isMinusOfVal1) Then 'value1�̓}�C�i�X�l
        If (isMinusOfVal2) Then 'value2�̓}�C�i�X�l
            tmpVal = addition(tmpVal1, tmpVal2)
            toRetSign = "-"
            
        Else 'value2�̓v���X�l
            tmpVal = substitution(tmpVal1, tmpVal2, substitutionWasMinus)
            If (substitutionWasMinus) Then
                toRetSign = ""
            Else
                toRetSign = "-"
            End If
        
        End If
        
    Else 'value1�̓v���X�l
        If (isMinusOfVal2) Then 'value2�̓}�C�i�X�l
            tmpVal = substitution(tmpVal1, tmpVal2, substitutionWasMinus)
            If (substitutionWasMinus) Then
                toRetSign = "-"
            Else
                toRetSign = ""
            End If
            
        Else 'value2�̓v���X�l
            tmpVal = addition(tmpVal1, tmpVal2)
            toRetSign = ""
        
        End If
    
    End If
    
    '�����_����
    intPrt = Left(tmpVal, Len(tmpVal) - Len(frcPrtOfVal1))
    frcPrt = Right(tmpVal, Len(frcPrtOfVal1))
    
    '�������̕s�v"0"�폜
    Do While (Right(frcPrt, 1) = "0")
        frcPrt = Left(frcPrt, Len(frcPrt) - 1)
    Loop
    
    addDecPntDecPnt = toRetSign & intPrt & IIf(frcPrt = "", "", DOT & frcPrt)

End Function

'
'1st������2nd�����Ŋ|����
'2nd������1~9�̂݉�
'
Public Function multipleDecPntByOneDig(ByVal multiplicand As String, ByVal multiplier As Integer) As Variant
    
    Dim multiplicandIsMinus As Boolean
    Dim multiplierIsMinus As Boolean
    
    Dim intPrtOfMultiplicand As String
    Dim frcPrtOfMultiplicand As String
    
    Dim intPrtOfAns As String
    Dim frcPrtOfAns As String
    Dim signOfAns As String
    
    Dim retOfSeparatePnt As Variant
    Dim retOfMultiple As String
    
    '�����`�F�b�N
    If Not ((-9 <= multiplier) And (multiplier <= 9)) Then
        multipleDecPntByOneDig = CVErr(xlErrNum) '#NUM!��Ԃ�
        Exit Function
        
    End If
    
    '��搔�̕�����`�F�b�N&�����A��������
    retOfSeparatePnt = separateToIntAndFrc(multiplicand, intPrtOfMultiplicand, frcPrtOfMultiplicand, multiplicandIsMinus, 10)
    If (retOfSeparatePnt <> 0) Then '10�i�l�Ƃ��ĕs��
        multipleDecPntByOneDig = retOfSeparatePnt
        Exit Function
        
    End If
    
    '�搔�̕����`�F�b�N
    If (multiplier = 0) Then '0�|���̏ꍇ
        multipleDecPntByOneDig = "0" '0��Ԃ�
        Exit Function
        
    ElseIf (multiplier < 0) Then '�搔���}�C�i�X�l
        multiplierIsMinus = True
        multiplier = Abs(multiplier)
        
    Else '�搔�̓v���X�l
        multiplierIsMinus = False
        
    End If
    
    '��Z
    retOfMultiple = multiple(intPrtOfMultiplicand & frcPrtOfMultiplicand, CByte(multiplier))
    
    intPrtOfAns = Left(retOfMultiple, Len(retOfMultiple) - Len(frcPrtOfMultiplicand))
    frcPrtOfAns = Right(retOfMultiple, Len(frcPrtOfMultiplicand))
    
    
    '�������̕s�v��0����菜��
    Do While (Left(intPrtOfAns, 1) = "0")
        intPrtOfAns = Right(intPrtOfAns, Len(intPrtOfAns) - 1)
    Loop
    If (intPrtOfAns = "") Then '�S��0��������
        intPrtOfAns = "0"
    End If
    
    '�������̕s�v��0����菜��
    Do While (Right(frcPrtOfAns, 1) = "0")
        frcPrtOfAns = Left(frcPrtOfAns, Len(frcPrtOfAns) - 1)
    Loop
    
    '��������
    If (multiplicandIsMinus Xor multiplierIsMinus) Then
        signOfAns = "-"
    Else
        signOfAns = ""
    
    End If
    
    multipleDecPntByOneDig = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)
    
End Function

'
'1st������2nd�����Ŋ���
'����؂�Ȃ����́A3rd�����Ŏw�肳�ꂽ�񐔂������������ʂ�Ԃ�
'2nd������1~9�̂݉�
'3rd������(-)�l�̏ꍇ�͍ی��Ȃ����葱����
'
Public Function divideDecPntByOneDig(ByVal dividend As String, ByVal divisor As Integer, Optional ByVal limitOfRepTimes As Long = DEFAULT_DIV_TIMES_FOR_INDIVISIBLE) As Variant
    
    Dim dividendIsMinus As Boolean
    Dim divisorIsMinus As Boolean
    
    Dim intPrtOfDividend As String
    Dim frcPrtOfDividend As String
    
    Dim intPrtOfAns As String
    Dim frcPrtOfAns As String
    Dim signOfAns As String
    
    Dim retOfSeparatePnt As Variant
    Dim retOfDivide As String
    
    '�����`�F�b�N
    If Not ((-9 <= divisor) And (divisor <= 9)) Then
        divideDecPntByOneDig = CVErr(xlErrNum) '#NUM!��Ԃ�
        Exit Function
        
    End If
    
    If (divisor = 0) Then
        divideDecPntByOneDig = CVErr(xlErrDiv0) '#DIV0!��Ԃ�
        Exit Function
        
    End If
    
    
    '�폜���̕�����`�F�b�N&�����A��������
    retOfSeparatePnt = separateToIntAndFrc(dividend, intPrtOfDividend, frcPrtOfDividend, dividendIsMinus, 10)
    If (retOfSeparatePnt <> 0) Then '10�i�l�Ƃ��ĕs��
        divideDecPntByOneDig = retOfSeparatePnt
        Exit Function
        
    End If
    
    '�����̕����`�F�b�N
    If (divisor = 0) Then '0�|���̏ꍇ
        divideDecPntByOneDig = "0" '0��Ԃ�
        Exit Function
        
    ElseIf (divisor < 0) Then '�������}�C�i�X�l
        divisorIsMinus = True
        divisor = Abs(divisor)
        
    Else '�����̓v���X�l
        divisorIsMinus = False
        
    End If
    
    '���Z
    retOfDivide = divide(intPrtOfDividend & frcPrtOfDividend, CByte(divisor), limitOfRepTimes)
    
    intPrtOfAns = Left(retOfDivide, Len(intPrtOfDividend))
    frcPrtOfAns = Right(retOfDivide, Len(retOfDivide) - Len(intPrtOfDividend))
    
    
    '�������̕s�v��0����菜��
    Do While (Left(intPrtOfAns, 1) = "0")
        intPrtOfAns = Right(intPrtOfAns, Len(intPrtOfAns) - 1)
    Loop
    If (intPrtOfAns = "") Then '�S��0��������
        intPrtOfAns = "0"
    End If
    
    '�������̕s�v��0����菜��
    Do While (Right(frcPrtOfAns, 1) = "0")
        frcPrtOfAns = Left(frcPrtOfAns, Len(frcPrtOfAns) - 1)
    Loop
    
    '��������
    If (dividendIsMinus Xor divisorIsMinus) Then
        signOfAns = "-"
    Else
        signOfAns = ""
    
    End If
    
    divideDecPntByOneDig = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)
    
End Function

'
'�������Ɛ������ɕ�������
'
'�����̏ꍇ��0��ԋp����
'���s�̏ꍇ��CvErr��ԋp����
'
Private Function separateToIntAndFrc(ByVal pnt As String, ByRef intPrt As String, ByRef frcPrt, ByRef isMinus As Boolean, ByVal radix As Byte) As Variant
    
    Dim idxOfDot As Long
    
    Dim value1IsMinus
    
    Dim retToIsMinus As Boolean
    Dim retToIntPrt As String
    Dim retToFrcPrt As String
    
    '���񒷃`�F�b�N
    If (Len(pnt) < 1) Then '�����񒷂�0
        separateToIntAndFrc = CVErr(xlErrValue) '#NUM!��Ԃ�
        Exit Function
        
    End If
    
    '��������菜��
    If (Left(pnt, 1) = "-") Then
        retToIsMinus = True
        pnt = Right(pnt, Len(pnt) - 1)
        If (pnt = "") Then
            pnt = "0"
        End If
        
    Else
        retToIsMinus = False
        
    End If
    
    '10�i�l�Ƃ��Đ��������`�F�b�N
    ret = checkNumeralPntStr(pnt, idxOfDot, radix)
    
    If (ret <> (Len(pnt) + 1)) Then 'value1��10�i�l�Ƃ��ĕs��
        separateToIntAndFrc = CVErr(xlErrNum) '#NUM!��Ԃ�
        Exit Function
        
    End If
    
    '�������Ə������ɕ�����
    
    '�������𒊏o����
    retToIntPrt = Left(pnt, idxOfDot - 1)
    If (retToIntPrt = "") Then '�������̋L�ڂ��Ȃ������ꍇ
        retToIntPrt = "0"
    End If
    
    '�������𒊏o����
    If (idxOfDot < Len(pnt)) Then '�������̋L�ڂ�����
        retToFrcPrt = Right(pnt, Len(pnt) - idxOfDot)
        
    Else '�������̋L�ڂ��Ȃ�
        retToFrcPrt = "0"
    
    End If
    
    intPrt = retToIntPrt
    frcPrt = retToFrcPrt
    isMinus = retToIsMinus
    separateToIntAndFrc = 0
    
End Function

'
'������n�i�l�����񂩂ǂ����`�F�b�N����
'
'�ԋp�l
'    n�i�l�����񂾂����̏ꍇ�͕����� + 1
'    �����łȂ��ꍇ�́A�ŏ��Ɍ�������10�i�����ȊO�̕����ʒu
'    �󕶎����w�肳�ꂽ�ꍇ��0��Ԃ�
'
'idxOfDot
'    �����_�����ʒu
'    �����_�����������ꍇ�͍ŏI�����ʒu+1���i�[����
'
'radix
'    �(2~16�̂�)
'
Private Function checkNumeralPntStr(ByVal decStr As String, ByRef idxOfDot As Long, ByRef radix As Byte) As Long
    
    Dim foundIdxOfDot As Long '�����_�������ŏ��Ɍ������������ʒu
    Dim cnt As Long
    Dim lpMx As Long
    Dim stCnt As Long
    
    Dim minOkChar1 As Integer
    Dim maxOkChar1 As Integer
    Dim minOkChar2 As Integer
    Dim maxOkChar2 As Integer
    
    Dim radixIsBiggerThan10 As Boolean
    
    lpMx = Len(decStr)
    
    '�����`�F�b�N
    If (lpMx = 0) Or (radix < 2) Or (16 < radix) Then
        checkNumeralPntStr = 0
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
    
    foundIdxOfDot = 0
    
    If (Left(decStr, 1) = "-") Then '�ŏ���(-)�����͖�������
        stCnt = 2
        
    Else
        stCnt = 1
        
    End If
    
    For cnt = stCnt To lpMx
        
        ch = Mid(decStr, cnt, 1)
        chCode = Asc(ch)
        
        If (chCode < minOkChar1) Or (maxOkChar1 < chCode) Then  '������0~9������ł��Ȃ�
            If IIf(radixIsBiggerThan10, (chCode < minOkChar2) Or (maxOkChar2 < chCode), True) Then '������A~F������ł��Ȃ�
            
                If (ch = DOT) Then
                    If (foundIdxOfDot = 0) Then '�����_������1���
                        foundIdxOfDot = cnt
                    
                    Else '�����_������2���
                        Exit For
                        
                    End If
                
                Else '������0~9������ł��Ȃ��A�����_�����ł��Ȃ�
                    Exit For
                    
                End If
            End If
        End If
    Next cnt
    
    
    If foundIdxOfDot = 0 Then '�����_������������Ȃ������ꍇ
        idxOfDot = lpMx + 1
        
    Else '�����_���������������ꍇ
        idxOfDot = foundIdxOfDot
    
    End If
    
    checkNumeralPntStr = cnt
    
End Function

'
'2����a�Z����
'
Private Function addition(ByVal val1 As String, ByVal val2 As String, Optional ByVal fill0Left As Boolean = True) As String
    
    Dim lenOfVal1 As Long
    Dim lenOfVal2 As Long
    
    Dim idxOfVal As Long
    Dim stringBuilder() As String
    
    Dim tmpStr As String
    Dim carrier As Integer
    
    Dim ret As String
    
    '
    '�L����10�i�l�ł��邩�̓`�F�b�N���Ȃ�
    '
    
    '�����񒷊m�F
    lenOfVal1 = Len(val1)
    lenOfVal2 = Len(val2)
    
    
    '0���ߊm�F
    If (lenOfVal1 > lenOfVal2) Then
        
        If fill0Left Then '����0���߂���
            val2 = String(lenOfVal1 - lenOfVal2, "0") & val2
            
        Else
            val2 = val2 & String(lenOfVal1 - lenOfVal2, "0")
            
        End If
        
        lenOfVal2 = lenOfVal1
        
    Else
        If fill0Left Then '����0���߂���
            val1 = String(lenOfVal2 - lenOfVal1, "0") & val1
            
        Else
            val1 = val1 & String(lenOfVal2 - lenOfVal1, "0")
            
        End If
        
        lenOfVal1 = lenOfVal2
        
    End If
    
    ReDim stringBuilder(lenOfVal1 - 1) '�̈�g��
    
    carrier = 0
    
    'addition���[�v
    For idxOfVal = lenOfVal1 To 1 Step -1
        tmpStr = Format(CInt(Mid(val1, idxOfVal, 1)) + CInt(Mid(val2, idxOfVal, 1)) + carrier, "00")
        carrier = CInt(Left(tmpStr, 1))
        stringBuilder(idxOfVal - 1) = Right(tmpStr, 1)
        
    Next idxOfVal
    
    ret = Join(stringBuilder, vbNullString)
    
    If (carrier > 0) Then '����������
        ret = CInt(carrier) & ret
        
    End If
    
    addition = ret
    
End Function

'
'val1����val2�����Z����
'
Private Function substitution(ByVal val1 As String, ByVal val2 As String, ByRef minus As Boolean) As String
    
    '�ϐ��錾
    Dim lenOfVal1 As Long
    Dim lenOfVal2 As Long
    
    Dim val1IsLarger As Integer '0:�s��, 1:yes, -1:no
    
    Dim idxOfVal As Long
    Dim idxMxOfVal As Long
    
    Dim stringBuilder() As String
    
    Dim wasMinus As Boolean
    
    '
    '�L����10�i�l���ǂ����̓`�F�b�N���Ȃ�
    '
    
    
    '�����񒷊m�F
    lenOfVal1 = Len(val1)
    lenOfVal2 = Len(val2)
    
    '0���ߊm�F
    If (lenOfVal1 > lenOfVal2) Then
        val2 = String(lenOfVal1 - lenOfVal2, "0") & val2
        
    Else
        val1 = String(lenOfVal2 - lenOfVal1, "0") & val1
        
    End If
    
    
    '�召��r�`�F�b�N
    idxOfVal = 1
    val1IsLarger = 0
    idxMxOfVal = Len(val1)
    Do
        val1Digit = CInt(Mid(val1, idxOfVal, 1))
        val2Digit = CInt(Mid(val2, idxOfVal, 1))
        
        '�ǂ��炩���傫�������� break
        If val1Digit > val2Digit Then
            val1IsLarger = 1
            Exit Do
        
        ElseIf val1Digit < val2Digit Then
            val1IsLarger = -1
            Exit Do
        
        End If
        
        idxOfVal = idxOfVal + 1
        
    Loop While idxOfVal <= idxMxOfVal
    
    
    If (val1IsLarger = 0) Then  '2���͓������l
        substitution = String(idxMxOfVal, "0")
        minus = False
        Exit Function
        
    End If
    
    ReDim stringBuilder(idxMxOfVal - 1) '�̈�g��
    
    If (val1IsLarger = -1) Then 'val2�̕����傫�����l��������
        '2�������ւ���
        buf = val1
        val1 = val2
        val2 = buf
        
        wasMinus = True
        
    Else
        wasMinus = False
        
    End If
    
    '���Z���[�v
    carrier = 0
    For idxOfVal = idxMxOfVal To 1 Step -1
        
        val1Digit = CInt(Mid(val1, idxOfVal, 1))
        val2Digit = CInt(Mid(val2, idxOfVal, 1))
        
        '�J�艺����`�F�b�N
        If val1Digit = 0 And carrier = -1 Then
            carrier = -1
            val1Digit = 9
            
        Else
            val1Digit = val1Digit + carrier
            carrier = 0
            
            If (val1Digit < val2Digit) Then
                val1Digit = 10 + val1Digit
                carrier = -1
                
            End If
            
        End If
        
        stringBuilder(idxOfVal - 1) = CStr(val1Digit - val2Digit)
        
    Next idxOfVal
    
    substitution = Join(stringBuilder, vbNullString)
    minus = wasMinus
    
    
End Function

'
'1�����l�ɂ���Z������
'
Private Function multiple(ByVal multiplicand As String, ByVal multiplier As Byte) As String
    
    Dim carrier As Byte
    Dim digitOfMultiplicand As Byte
    Dim digitIdxOfMultiplicand As Long
    Dim stringBuilder() As String '����Z���ʊi�[�p
    Dim idxOfStringBuilder As Long
    
    '
    'multiplicand���L����10�i�l�ł��邩�̓`�F�b�N���Ȃ�
    'multiplier��1���ł��邱�Ƃ̓`�F�b�N���Ȃ�
    '
    
    
    If (multiplier = 0) Then '�~0�̏ꍇ��0��Ԃ�
        multiple = "0"
        Exit Function
    
    ElseIf (multiplier = 1) Then '�~1�̏ꍇ�͂��̂܂ܕԂ�
        multiple = multiplicand
        Exit Function
        
    End If
    
    digitIdxOfMultiplicand = Len(multiplicand)
    carrier = 0
    idxOfStringBuilder = 0
    
    Do
        digitOfMultiplicand = Mid(multiplicand, digitIdxOfMultiplicand, 1)
        
        tmpStr = Format(digitOfMultiplicand * multiplier + carrier, "00")
        
        carrier = CInt(Left(tmpStr, 1))
        
        ReDim Preserve stringBuilder(idxOfStringBuilder) '�̈�g��
        stringBuilder(idxOfStringBuilder) = Right(tmpStr, 1)
        
        digitIdxOfMultiplicand = digitIdxOfMultiplicand - 1
        idxOfStringBuilder = idxOfStringBuilder + 1
    
    Loop While digitIdxOfMultiplicand > 0 '��搔���c���Ă����
    
    If (carrier <> "0") Then
        ReDim Preserve stringBuilder(idxOfStringBuilder) '�̈�g��
        stringBuilder(idxOfStringBuilder) = carrier
        
    End If
    
    multiple = Join(invertStringArray(stringBuilder), vbNullString) '������A��
    
End Function

'
'1�����l�ɂ�鏜�Z������
'
'limitOfRepTimes:
'    Indivisible Number�ɑ΂���divide�񐔐���
'    (-)�l��ݒ肵���ꍇ�́A�����Ɋ��葱����
'
Private Function divide(ByVal dividend As String, ByVal divisor As Byte, ByVal limitOfRepTimes As Long) As String

    '�ϐ��錾
    Dim quot As Byte   '��
    Dim rmnd As Byte '�]��
    
    Dim repTimes As Long 'IndivisibleNumber�ɑ΂���divide��
    
    Dim digitOfDividend As Byte '�ꎞ�폜��
    
    Dim stringBuilder() As String '����Z���ʊi�[�p
    Dim digitIdxOfDividend As Long 'Division���ʕ�����
    
    '
    'dividend���L����10�i�l�ł��邩�̓`�F�b�N���Ȃ�
    'divisor��1���ł��邱�Ƃ̓`�F�b�N���Ȃ�
    '
    
    '1���`�F�b�N
    If divisor = 1 Then
        divide = dividend '1���̏ꍇ�͂��̂܂ܕԂ�
        Exit Function
        
    End If
    
    '������
    rmnd = 0
    digitIdxOfDividend = 1
    repTimes = 0
    
    '���s���[�v
    Do
        digitOfDividend = CByte(CStr(rmnd) & Mid(dividend, digitIdxOfDividend, 1)) '��ʌ��̗]�� & �Y����
        
        quot = digitOfDividend \ divisor '��
        rmnd = digitOfDividend Mod divisor '�]��
        
        ReDim Preserve stringBuilder(digitIdxOfDividend - 1) '�̈�g��
        stringBuilder(digitIdxOfDividend - 1) = CStr(quot) '����ǋL
        
        digitIdxOfDividend = digitIdxOfDividend + 1
        
        If (rmnd > 0) And (Len(dividend) < digitIdxOfDividend) Then '�]�肪���邯��ǁA���̌�������
            
            If (limitOfRepTimes > -1) And (repTimes < limitOfRepTimes) Then '�ċA�v�Z�񐔂��w��񐔈ȉ�
                dividend = dividend & "0" '"0"��t��
                repTimes = repTimes + 1
                
            End If
            
        End If
        
    Loop While digitIdxOfDividend <= Len(dividend) '�ŏI�����ɓ��B���Ȃ���
    
    divide = Join(stringBuilder, vbNullString) '������A��
    
End Function

'
'10�i������������2�i���������ɕϊ�����
'
'numOfDigits:
'    ���߂鏬���_�ȉ��̌���
'    0���w�肵���ꍇ�͋󕶎���Ԃ�
'
Private Function convFrcPrtOfDecPntToFrcPrtOfBinPnt(ByVal frcPt As String, ByVal numOfDigits As Long) As String
    
    Dim stringBuilder() As String 'bit�i�[�p
    Dim repTimes As Long
    
    '
    '�L��10�i�������񂩂ǂ����̓`�F�b�N���Ȃ�
    '
    
    If (frcPt = "") Or (numOfDigits = 0) Then '�󕶎��w�肩�A���߂錅��=0
        convFrcPrtOfDecPntToFrcPrtOfBinPnt = "" '�󕶎���ԋp
        Exit Function
        
    End If
    
    '�E����0����菜��
    Do While Right(frcPt, 1) = "0"
        frcPt = Left(frcPt, Len(frcPt) - 1)
        
    Loop
    
    If (frcPt = "") Then '�S��"0"��������
        convFrcPrtOfDecPntToFrcPrtOfBinPnt = "0"
        Exit Function
        
    End If
    
    '�|���Z����
    repTimes = 0
    sizeOfStringBuilder = 0
    Do
        
        tmp = multiple(frcPt, 2)
        
        ReDim Preserve stringBuilder(repTimes) '�̈�g��
        
        If (Len(tmp) > Len(frcPt)) Then '���オ�肪���������ꍇ
            stringBuilder(repTimes) = "1"
            frcPt = Right(tmp, Len(tmp) - 1)
            
        Else '���オ�肪�������Ȃ������ꍇ
            stringBuilder(repTimes) = "0"
            frcPt = tmp
        
        End If
        
        If frcPt = "0" Then 'bin�ϊ��I��
            Exit Do
            
        ElseIf Right(frcPt, 1) = "0" Then '�E���̂�"0"����������
            frcPt = Left(frcPt, Len(frcPt) - 1) '"0"�͏���
        
        End If
        
        repTimes = repTimes + 1
        
    Loop While IIf(numOfDigits < 0, True, (repTimes < numOfDigits)) '�J��Ԃ��񐔈ȉ�
    
    convFrcPrtOfDecPntToFrcPrtOfBinPnt = Join(stringBuilder, vbNullString) '������A��
    
End Function

'
'10�i������������2�i���������ɕϊ�����
'
Private Function convIntPrtOfDecPntToIntPrtOfBinPnt(ByVal intPt As String) As String
    
    Dim stringBuilder() As String 'bit�i�[�p
    Dim sizeOfStringBuilder As Long
    
    '
    '�L��10�i�������񂩂ǂ����̓`�F�b�N���Ȃ�
    '
    
    If (intPt = "") Then '�󕶎��̏ꍇ
        convIntPrtOfDecPntToIntPrtOfBinPnt = "0"
        Exit Function
        
    End If
    
    '������0����菜��
    Do While Left(intPt, 1) = "0"
        intPt = Right(intPt, Len(intPt) - 1)
        
    Loop
    
    If (intPt = "") Then '�S��"0"��������
        convIntPrtOfDecPntToIntPrtOfBinPnt = "0"
        Exit Function
        
    End If
    
    sizeOfStringBuilder = 0
    
    'bit����
    Do While (intPt <> "1")
        ret = divide(intPt, 2, 0)
        
        ReDim Preserve stringBuilder(sizeOfStringBuilder) '�̈�g��
        
        If (Right(intPt, 1) Like "[0,2,4,6,8]") Then '2�Ŋ��������܂��0
            stringBuilder(sizeOfStringBuilder) = "0"
            
        Else '2�Ŋ��������܂��1
            stringBuilder(sizeOfStringBuilder) = "1"
            
        End If
        
        sizeOfStringBuilder = sizeOfStringBuilder + 1
        
        If (Left(ret, 1) = "0") Then
            intPt = Right(ret, Len(ret) - 1)
            
        Else
            intPt = ret
            
        End If
        
    Loop
    
    '�ŏ��Bit�t��
    ReDim Preserve stringBuilder(sizeOfStringBuilder) '�̈�g��
    stringBuilder(sizeOfStringBuilder) = "1"
    
    convIntPrtOfDecPntToIntPrtOfBinPnt = Join(invertStringArray(stringBuilder), vbNullString) '������A��
    
End Function

'
'2�i������������10�i���������ɕϊ�����
'
Private Function convFrcPrtOfBinPntToFrcPrtOfDecPnt(ByVal frcPt As String) As String
    
    Dim lpMx As Long
    Dim decStr As String
    
    Dim minusNpowerOf2 As String
    
    '
    '�L��2�i�����񂩂ǂ����̓`�F�b�N���Ȃ�
    '
    
    '�����`�F�b�N
    If (frcPt = "") Then '�󕶎��w��̏ꍇ
        convFrcPrtOfBinPntToFrcPrtOfDecPnt = "0" '"0"��Ԃ�
        Exit Function
        
    End If
    
    lpMx = Len(frcPt)
    decStr = "0"
    minusNpowerOf2 = "5"
    
    '�������[�v
    For cnt = 1 To lpMx
        If (Mid(frcPt, cnt, 1) = "1") Then
            decStr = addition(minusNpowerOf2, decStr, False)
        End If
        
        minusNpowerOf2 = divide(minusNpowerOf2, 2, 1)
        
    Next cnt
    
    convFrcPrtOfBinPntToFrcPrtOfDecPnt = decStr
    
End Function

'
'2�i������������10�i���������ɕϊ�����
'
Private Function convIntPrtOfBinToIntPrtOfDecPrt(ByVal intPt As String) As String
    
    Dim nPowerOf2 As String
    Dim decStr As String
    
    '
    '�L��2�i�����񂩂ǂ����̓`�F�b�N���Ȃ�
    '
    
    '�����`�F�b�N
    If (intPt = "") Then '�󕶎��w��̏ꍇ
        convIntPrtOfBinToIntPrtOfDecPrt = "0" '"0"��Ԃ�
        Exit Function
        
    End If
    
    nPowerOf2 = "1"
    decStr = "0"
    
    For cnt = Len(intPt) To 1 Step -1
        
        If (Mid(intPt, cnt, 1) = "1") Then
            decStr = addition(nPowerOf2, decStr, True)
        End If
        
        nPowerOf2 = multiple(nPowerOf2, 2)
        
    Next cnt
    
    convIntPrtOfBinToIntPrtOfDecPrt = decStr
    
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



