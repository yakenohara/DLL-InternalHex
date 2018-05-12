Attribute VB_Name = "NumericString"
'�萔

Public Const DOT As String = "." '�����_�\�L

'����؂�Ȃ����l�ɑ΂��ĉ��񊄂�Z���邩
Const DEFAULT_DIV_TIMES_FOR_INDIVISIBLE As Long = 255

'10�i��n�i�ϊ����́A���������̌��E�Z�o����
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
'�����t��10�i�l���珬���t��n�i�l�ɕϊ�����
'
'numOfDigits
'    �����Z�o���̌��E���Z��
'
Public Function convDecPntToNPnt(ByVal pntStr As String, ByVal radix As Byte, Optional numOfDigits As Long = DEFAULT_FRC_DIGITS) As Variant
    
    Dim intPtOfBefore As String '������
    Dim frcPtOfBefore As String '������
    
    Dim intPtOfAfter As String '������
    Dim frcPtOfAfter As String '������
    Dim isMinus As Boolean
    Dim sign As String '����
    
    Dim ret As Long
    Dim idxOfDot As Long
    
    Dim retOfSeparatePnt As Variant
    
    '��`�F�b�N
    If (radix < 2) Or (16 < radix) Then
        convDecPntToNPnt = CVErr(xlErrValue) '#VALUE!��ԋp
        Exit Function
        
    End If
    
    'val1�̕�����`�F�b�N&�����A��������
    retOfSeparatePnt = separateToIntAndFrc(pntStr, intPtOfBefore, frcPtOfBefore, isMinus, 10)
    
    If (retOfSeparatePnt <> 0) Then 'val1��n�i�l�Ƃ��ĕs��
        convDecPntToNPnt = CVErr(xlErrValue) '#VALUE!��ԋp
        Exit Function
        
    End If
    
    '�������̕s�v��0����菜��
    Do While (Left(intPtOfBefore, 1) = "0")
        intPtOfBefore = Right(intPtOfBefore, Len(intPtOfBefore) - 1)
        
    Loop
    
    If (intPtOfBefore = "") Then '�S��"0"��������
        intPtOfBefore = "0"
        
    End If
    
    '�������̕s�v��0����菜��
    Do While (Right(frcPtOfBefore, 1) = "0")
        frcPtOfBefore = Left(frcPtOfBefore, Len(frcPtOfBefore) - 1)
        
    Loop
    
    '�}�C�i�X�l�`�F�b�N
    If (isMinus) Then
        sign = "-"
        
    Else
        sign = ""
        
    End If
    
    '��������n�i�ϊ�
    intPtOfAfter = convIntPrtOfDecPntToIntPrtOfNPnt(intPtOfBefore, radix)
    
    '��������n�i�ϊ�
    If (frcPtOfBefore = "") Then '�������͑��݂��Ȃ��ꍇ
        frcPtOfAfter = ""
        
    Else '�����������݂���ꍇ
        frcPtOfAfter = convFrcPrtOfDecPntToFrcPrtOfNPnt(frcPtOfBefore, numOfDigits, radix)
        
        If (frcPtOfAfter <> "") Then
            frcPtOfAfter = DOT & frcPtOfAfter
            
        End If
        
    End If
    
    '�����񌋍�
    convDecPntToNPnt = sign & intPtOfAfter & frcPtOfAfter
    
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
        intPtOfAfter = convIntPrtOfDecPntToIntPrtOfNPnt(intPtOfBefore, 2) '10�i�����ϊ��Ώۂ̎�
        
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
            frcPtOfAfter = convFrcPrtOfDecPntToFrcPrtOfNPnt(frcPtOfBefore, numOfDigits, 2) '10�i�����ϊ��Ώۂ̎�
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
'radix
'   2~16 �̂�
'
Public Function addNPntNPnt(ByVal value1 As String, ByVal value2 As String, Optional ByVal radix As Byte = 10) As Variant
    
    Dim intPrtOfVal1 As String
    Dim frcPrtOfVal1 As String
    Dim isMinusOfVal1 As Boolean
    
    Dim intPrtOfVal2 As String
    Dim frcPrtOfVal2 As String
    Dim isMinusOfVal2 As Boolean
    
    Dim retOfSeparatePnt As Variant
    
    Dim toRetSign As String
    
    Dim subtractionWasMinus As Boolean
    
    '��`�F�b�N
    If (radix < 2) Or (16 < radix) Then
        addNPntNPnt = CVErr(xlErrValue) '#VALUE!��ԋp
        Exit Function
        
    End If
    
    'val1�̕�����`�F�b�N&�����A��������
    retOfSeparatePnt = separateToIntAndFrc(value1, intPrtOfVal1, frcPrtOfVal1, isMinusOfVal1, radix)
    
    If (retOfSeparatePnt <> 0) Then 'val1��n�i�l�Ƃ��ĕs��
        addNPntNPnt = CVErr(xlErrValue) '#VALUE!��ԋp
        Exit Function
        
    End If
    
    'valw�̕�����`�F�b�N&�����A��������
    retOfSeparatePnt = separateToIntAndFrc(value2, intPrtOfVal2, frcPrtOfVal2, isMinusOfVal2, radix)
    
    If (retOfSeparatePnt <> 0) Then 'val2��n�i�l�Ƃ��ĕs��
        addNPntNPnt = CVErr(xlErrValue) '#VALUE!��ԋp
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
            tmpVal = add(tmpVal1, tmpVal2, radix)
            toRetSign = "-"
            
        Else 'value2�̓v���X�l
            tmpVal = subtract(tmpVal1, tmpVal2, subtractionWasMinus, radix)
            If (subtractionWasMinus) Then
                toRetSign = "-"
            Else
                toRetSign = ""
            End If
        
        End If
        
    Else 'value1�̓v���X�l
        If (isMinusOfVal2) Then 'value2�̓}�C�i�X�l
            tmpVal = subtract(tmpVal1, tmpVal2, subtractionWasMinus, radix)
            If (subtractionWasMinus) Then
                toRetSign = "-"
            Else
                toRetSign = ""
            End If
            
        Else 'value2�̓v���X�l
            tmpVal = add(tmpVal1, tmpVal2, radix)
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
    
    addNPntNPnt = toRetSign & intPrt & IIf(frcPrt = "", "", DOT & frcPrt)

End Function

'
'1st������2nd�����Ŋ|����
'2nd������1~9�̂݉�
'
'radix
'    2~16 �̂�
'
Public Function multipleNPntNPnt(ByVal multiplicand As String, ByVal multiplierCh As String, Optional radix As Byte = 10) As Variant

    Dim multiplicandIsMinus As Boolean
    Dim multiplierIsMinus As Boolean
    
    Dim intPrtOfMultiplicand As String
    Dim frcPrtOfMultiplicand As String
    
    Dim intPrtOfMultiplier As String
    Dim frcPrtOfMultiplier As String
    
    Dim intPrtOfAns As String
    Dim frcPrtOfAns As String
    Dim signOfAns As String
    
    Dim retOfSeparatePnt As Variant
    Dim retOfMultiple As String
    
    '�搔�̕�����`�F�b�N&�����A��������
    retOfSeparatePnt = separateToIntAndFrc(multiplierCh, intPrtOfMultiplier, frcPrtOfMultiplier, multiplierIsMinus, radix)
    If (retOfSeparatePnt <> 0) Then 'n�i�l�Ƃ��ĕs��
        multipleNPntNPnt = retOfSeparatePnt
        Exit Function
        
    End If
    
    '�����������݂��Ȃ��ꍇ�́A��菜��
    If (frcPrtOfMultiplier = "0") Then
        frcPrtOfMultiplier = ""
    End If
    
    '��搔�̕�����`�F�b�N&�����A��������
    retOfSeparatePnt = separateToIntAndFrc(multiplicand, intPrtOfMultiplicand, frcPrtOfMultiplicand, multiplicandIsMinus, radix)
    If (retOfSeparatePnt <> 0) Then 'n�i�l�Ƃ��ĕs��
        multipleNPntNPnt = retOfSeparatePnt
        Exit Function
        
    End If
    
    '�����������݂��Ȃ��ꍇ�́A��菜��
    If (frcPrtOfMultiplicand = "0") Then
        frcPrtOfMultiplicand = ""
    End If
    
    
    '��Z
    retOfMultiple = multiple(intPrtOfMultiplicand & frcPrtOfMultiplicand, intPrtOfMultiplier & frcPrtOfMultiplier, radix)
    
    intPrtOfAns = Left(retOfMultiple, Len(retOfMultiple) - (Len(frcPrtOfMultiplicand) + Len(frcPrtOfMultiplier)))
    frcPrtOfAns = Right(retOfMultiple, Len(frcPrtOfMultiplicand) + Len(frcPrtOfMultiplier))
    
    
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
        
        If (intPrtOfAns = "0") And (frcPrtOfAns = "") Then
            signOfAns = ""
            
        Else
            signOfAns = "-"
            
        End If
    Else
        signOfAns = ""
    
    End If
    
    multipleNPntNPnt = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)
    
End Function

'
'1st������2nd�����Ŋ���
'����؂�Ȃ����́A3rd�����Ŏw�肳�ꂽ�񐔂������������ʂ�Ԃ�
'2nd������1~9�̂݉�
'3rd������(-)�l�̏ꍇ�͍ی��Ȃ����葱����
'
Public Function divideNPntByOneDig(ByVal dividend As String, ByVal divisorCh As String, Optional ByVal radix As Byte = 10, Optional ByVal limitOfRepTimes As Long = DEFAULT_DIV_TIMES_FOR_INDIVISIBLE) As Variant

    Dim divisor As Integer
    
    Dim dividendIsMinus As Boolean
    Dim divisorIsMinus As Boolean
    
    Dim intPrtOfDividend As String
    Dim frcPrtOfDividend As String
    
    Dim intPrtOfDivisor As String
    Dim frcPrtOfDivisor As String
    
    Dim intPrtOfAns As String
    Dim frcPrtOfAns As String
    Dim signOfAns As String
    
    Dim retOfSeparatePnt As Variant
    Dim retOfDivide As String
    
    Dim rm As String
    
    '�����̕�����`�F�b�N&�����A��������
    retOfSeparatePnt = separateToIntAndFrc(divisorCh, intPrtOfDivisor, frcPrtOfDivisor, divisorIsMinus, radix)
    If (retOfSeparatePnt <> 0) Then 'n�i�l�Ƃ��ĕs��
        divideNPntByOneDig = retOfSeparatePnt
        Exit Function
        
    ElseIf (frcPrtOfDivisor <> "0") Or (Len(intPrtOfDivisor) > 1) Then '�����������݂��� / ��������2���ȏ�
        divideNPntByOneDig = CVErr(xlErrNum) '#NUM!��Ԃ�
        Exit Function
        
    End If
    
    divisor = convNCharToByte(intPrtOfDivisor)
    
    If (radix <= divisor) Then
        divideNPntByOneDig = CVErr(xlErrNum) '#NUM!��Ԃ�
        Exit Function
        
    ElseIf (divisor = 0) Then
        divideNPntByOneDig = CVErr(xlErrDiv0) '#DIV0!��Ԃ�
        Exit Function
        
    End If
    
    '�폜���̕�����`�F�b�N&�����A��������
    retOfSeparatePnt = separateToIntAndFrc(dividend, intPrtOfDividend, frcPrtOfDividend, dividendIsMinus, radix)
    If (retOfSeparatePnt <> 0) Then 'n�i�l�Ƃ��ĕs��
        divideNPntByOneDig = retOfSeparatePnt
        Exit Function
        
    End If
    
    '�����������݂��Ȃ��ꍇ�́A��菜��
    If (frcPrtOfDividend = "0") Then
        frcPrtOfDividend = ""
    End If
    
    '���Z
    retOfDivide = divideByOneDig(intPrtOfDividend & frcPrtOfDividend, intPrtOfDivisor, rm, limitOfRepTimes, radix)
    
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
        
        If (intPrtOfAns = "0") And (frcPrtOfAns = "") Then
            signOfAns = ""
            
        Else
            signOfAns = "-"
            
        End If
    Else
        signOfAns = ""
    
    End If
    
    divideNPntByOneDig = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)
    
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
        separateToIntAndFrc = CVErr(xlErrValue) '#VALUE!��Ԃ�
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
    
    'n�i�l�Ƃ��Đ��������`�F�b�N
    ret = checkNumeralPntStr(pnt, idxOfDot, radix)
    
    If (ret <> (Len(pnt) + 1)) Then 'value1��n�i�l�Ƃ��ĕs��
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
Private Function add(ByVal val1 As String, ByVal val2 As String, ByVal radix As Byte, Optional ByVal fill0Left As Boolean = True) As String
    
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
        
        tmpDigitOfVal1 = convNCharToByte(Mid(val1, idxOfVal, 1))
        tmpDigitOfVal2 = convNCharToByte(Mid(val2, idxOfVal, 1))
        
        tmpStr = convIntPrtOfDecPntToIntPrtOfNPnt(tmpDigitOfVal1 + tmpDigitOfVal2 + carrier, radix)
        
        If (Len(tmpStr) = 2) Then '������������
            carrier = 1
            
        Else
            carrier = 0
            
        End If
        
        stringBuilder(idxOfVal - 1) = Right(tmpStr, 1)
        
    Next idxOfVal
    
    ret = Join(stringBuilder, vbNullString)
    
    If (carrier > 0) Then '����������
        ret = CInt(carrier) & ret
        
    End If
    
    add = ret
    
End Function

'
'val1����val2�����Z����
'
Private Function subtract(ByVal val1 As String, ByVal val2 As String, ByRef minus As Boolean, ByVal radix As Byte) As String
    
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
        val1Digit = convNCharToByte(Mid(val1, idxOfVal, 1))
        val2Digit = convNCharToByte(Mid(val2, idxOfVal, 1))
        
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
        subtract = String(idxMxOfVal, "0")
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
        
        val1Digit = convNCharToByte(Mid(val1, idxOfVal, 1))
        val2Digit = convNCharToByte(Mid(val2, idxOfVal, 1))
        
        '�J�艺����`�F�b�N
        If val1Digit = 0 And carrier = -1 Then
            carrier = -1
            val1Digit = radix - 1
            
        Else
            val1Digit = val1Digit + carrier
            carrier = 0
            
            If (val1Digit < val2Digit) Then
                val1Digit = radix + val1Digit
                carrier = -1
                
            End If
            
        End If
        
        stringBuilder(idxOfVal - 1) = convIntPrtOfDecPntToIntPrtOfNPnt(val1Digit - val2Digit, radix)
        
    Next idxOfVal
    
    subtract = Join(stringBuilder, vbNullString)
    minus = wasMinus
    
    
End Function

'
'��Z������
'
Private Function multiple(ByVal multiplicand As String, ByVal multiplier As String, ByVal radix As Byte) As String

    Dim ansOfMultipleByOneDig As String
    Dim numOf0 As Long
    Dim tmpAns As String
    
    '
    '�L����n�i�l�ł��邩�̓`�F�b�N���Ȃ�
    '
    
    'multiplier�̕s�v��0����菜��
    Do While (Left(multiplier, 1) = "0")
        multiplier = Right(multiplier, Len(multiplier) - 1)
        
    Loop
    
    If (multiplier = "") Then '�S��"0"��������
        multiple = String(Len(multiplicand), "0")
        Exit Function
        
    ElseIf (multiplierCh = "1") Then '1�|���̏ꍇ�͂��̂܂ܕԂ�
        multiple = multiplicand
        Exit Function
        
    End If
    
    numOf0 = 0
    tmpAns = "0"
    
    '��Z���[�v
    For idx = Len(multiplier) To 1 Step -1
        
        ansOfMultipleByOneDig = multipleByOneDig(multiplicand, Mid(multiplier, idx, 1), radix)
        tmpAns = add(tmpAns, ansOfMultipleByOneDig & String(numOf0, "0"), radix)
        
        numOf0 = numOf0 + 1
        
    Next idx
    
    multiple = tmpAns
    
End Function

'
'1�����l�ɂ���Z������
'
Private Function multipleByOneDig(ByVal multiplicand As String, ByVal multiplierCh As String, ByVal radix As Byte) As String

    Dim carrier As Byte
    Dim digitOfMultiplicand As Byte
    Dim multiplier As Byte
    Dim digitIdxOfMultiplicand As Long
    Dim stringBuilder() As String '����Z���ʊi�[�p
    Dim idxOfStringBuilder As Long
    
    '
    'multiplicand���L����n�i�l�ł��邩�̓`�F�b�N���Ȃ�
    'multiplierCh��1���ł��邱�Ƃ̓`�F�b�N���Ȃ�
    '
    
    If (multiplierCh = "0") Then '0�|���̏ꍇ��0��Ԃ�
        multipleByOneDig = String(Len(multiplicand), "0")
        Exit Function
    
    ElseIf (multiplierCh = "1") Then '1�|���̏ꍇ�͂��̂܂ܕԂ�
        multipleByOneDig = multiplicand
        Exit Function
        
    End If
    
    multiplier = convNCharToByte(multiplierCh)
    digitIdxOfMultiplicand = Len(multiplicand)
    carrier = 0
    idxOfStringBuilder = 0
    
    Do
        digitOfMultiplicand = convNCharToByte(Mid(multiplicand, digitIdxOfMultiplicand, 1))
        
        tmpStr = convIntPrtOfDecPntToIntPrtOfNPnt(digitOfMultiplicand * multiplier + carrier, radix)
        
        'carrier����
        If (Len(tmpStr) = 2) Then '�����������ꍇ
            carrier = convNCharToByte(Left(tmpStr, 1))
            
        Else '���������Ȃ������ꍇ
            carrier = 0
            
        End If
        
        ReDim Preserve stringBuilder(idxOfStringBuilder) '�̈�g��
        stringBuilder(idxOfStringBuilder) = Right(tmpStr, 1)
        
        digitIdxOfMultiplicand = digitIdxOfMultiplicand - 1
        idxOfStringBuilder = idxOfStringBuilder + 1
    
    Loop While digitIdxOfMultiplicand > 0 '��搔���c���Ă����
    
    '���オ��`�F�b�N
    If (carrier > 0) Then
        ReDim Preserve stringBuilder(idxOfStringBuilder) '�̈�g��
        stringBuilder(idxOfStringBuilder) = convIntPrtOfDecPntToIntPrtOfNPnt(carrier, radix)
        
    End If
    
    multipleByOneDig = Join(invertStringArray(stringBuilder), vbNullString) '������A��
    
End Function

'
'1�����l�ɂ�鏜�Z������
'
'limitOfRepTimes:
'    Indivisible Number�ɑ΂���divide�񐔐���
'    (-)�l��ݒ肵���ꍇ�́A�����Ɋ��葱����
'remainder
'    ��]�B
'    �����_�ȉ��ƂȂ�ꍇ�́A
'    ��ԍ���1���ڂƂ��ď����_����菜��������������ƂȂ�B
'    ex:)
'    �y�O��z10 / 8 = 1.2 + 0.4
'    �y���s���@�zx = divideByOneDig("10", "8", rm, 1, 10)
'    �y���ʁz x:012
'            rm:04
'
Private Function divideByOneDig(ByVal dividend As String, ByVal divisorCh As String, ByRef remainder As String, ByVal limitOfRepTimes As Long, ByVal radix As Byte) As String

    '�ϐ��錾
    Dim divisor As Byte
    Dim quot As Byte   '��
    Dim rmnd As Byte '�]��
    
    Dim repTimes As Long 'IndivisibleNumber�ɑ΂���divide��
    
    Dim digitOfDividend As Byte '�ꎞ�폜��
    
    Dim stringBuilder() As String '���i�[�p
    Dim stringBuilderRM() As String '��]�i�[�p
    Dim digitIdxOfDividend As Long 'Division���ʕ�����
    
    '
    'dividend���L����n�i�l�ł��邩�̓`�F�b�N���Ȃ�
    'divisor��1���ł��邱�Ƃ̓`�F�b�N���Ȃ�
    '
    
    '1���`�F�b�N
    If divisorCh = "1" Then
        divideByOneDig = dividend '1���̏ꍇ�͂��̂܂ܕԂ�
        Exit Function
        
    End If
    
    '������
    rmnd = 0
    digitIdxOfDividend = 1
    repTimes = 0
    remainder = ""
    
    divisor = convNCharToByte(divisorCh)
    
    '���s���[�v
    Do
        digitOfDividend = rmnd * radix + convNCharToByte(Mid(dividend, digitIdxOfDividend, 1)) '��ʌ��̗]�� & �Y����
        
        quot = digitOfDividend \ divisor '��
        rmnd = digitOfDividend Mod divisor '�]��
        
        ReDim Preserve stringBuilder(digitIdxOfDividend - 1) '�̈�g��
        stringBuilder(digitIdxOfDividend - 1) = convIntPrtOfDecPntToIntPrtOfNPnt(quot, radix) '����ǋL
        
        digitIdxOfDividend = digitIdxOfDividend + 1
        
        If (rmnd > 0) And (Len(dividend) < digitIdxOfDividend) Then '�]�肪���邯��ǁA���̌�������
            
            If (limitOfRepTimes > -1) And (repTimes < limitOfRepTimes) Then '�ċA�v�Z�񐔂��w��񐔈ȉ�
                dividend = dividend & "0" '"0"��t��
                
                ReDim Preserve stringBuilderRM(repTimes) '�̈�g��
                stringBuilderRM(repTimes) = "0"
                
                repTimes = repTimes + 1
                
            End If
            
        End If
        
    Loop While digitIdxOfDividend <= Len(dividend) '�ŏI�����ɓ��B���Ȃ���
    
    If (rmnd = 0) Then '�]�肪0�̂Ƃ�
        remainder = "0"
        
    Else '�]�肪���݂��鎞
        
        ReDim Preserve stringBuilderRM(repTimes) '�̈�g��
        stringBuilderRM(repTimes) = convIntPrtOfDecPntToIntPrtOfNPnt(rmnd, radix)
        remainder = Join(stringBuilderRM, vbNullString) '������A��
    
    End If
    
    divideByOneDig = Join(stringBuilder, vbNullString) '������A��
    
    
End Function

'
'10�i��������n�i�������ɕϊ�����
'
'radix:
'    2~16 �̂�
'
Private Function convIntPrtOfDecPntToIntPrtOfNPnt(ByVal decInt As String, ByVal radix As Byte) As String
    
    Dim stringBuilder() As String '�ϊ��㕶���񐶐��p
    Dim sizeOfStringBuilder As Long
    Dim rm As String
    
    '
    '�L��10�i���l�����񂩂ǂ����̓`�F�b�N���Ȃ�
    '
    
    If (decInt = "") Then
        convIntPrtOfDecPntToIntPrtOfNPnt = "0"
        Exit Function
        
    End If
    
    '�����̕s�v��"0"����菜��
    Do While Left(decInt, 1) = "0"
        decInt = Right(decInt, Len(decInt) - 1)
        
    Loop
    
    If (decInt = "") Then '�S��"0"��������
        convIntPrtOfDecPntToIntPrtOfNPnt = "0"
        Exit Function
        
    End If
    
    If (radix = 10) Then '10�i��10�i�ϊ���������
        convIntPrtOfDecPntToIntPrtOfNPnt = decInt '�ϊ������ɕԂ�
        Exit Function
        
    End If
    
    sizeOfStringBuilder = 0
    strLenOfRadix = Len(CStr(radix))
    
    If (radix > 9) Then 'A~F�ɊY�����鐔�l�̏ꍇ
        
        strOfRadix = Chr((radix - 10) + Asc("A"))
        
    Else '0~9�ɊY�����鐔�l�̏ꍇ
        strOfRadix = CStr(radix)
        
    End If
    
    'bit����
    Do While True
        
        If (Len(decInt) <= strLenOfRadix) Then
            
            If (CByte(decInt) < radix) Then '��Ŋ���鐔���Ȃ��Ȃ���
                Exit Do
                
            End If
        End If
        
        decInt = divideByOneDig(decInt, strOfRadix, rm, 0, 10)
        
        '�����̕s�v��"0"����菜��
        Do While Left(decInt, 1) = "0"
            decInt = Right(decInt, Len(decInt) - 1)
            
        Loop
        If (decInt = "") Then '�S��"0"��������
            decInt = "0"
            
        End If
        
        ReDim Preserve stringBuilder(sizeOfStringBuilder) '�̈�g��
        stringBuilder(sizeOfStringBuilder) = convByteToNChar(rm)
        
        sizeOfStringBuilder = sizeOfStringBuilder + 1
        
    Loop
    
    '�ŏ��Bit�t��
    ReDim Preserve stringBuilder(sizeOfStringBuilder) '�̈�g��
    stringBuilder(sizeOfStringBuilder) = convByteToNChar(decInt)
    
    convIntPrtOfDecPntToIntPrtOfNPnt = Join(invertStringArray(stringBuilder), vbNullString) '������A��
    
End Function

'
'10�i������������n�i���������ɕϊ�����
'
'numOfDigits:
'    ���߂鏬���_�ȉ��̌���
'    0���w�肵���ꍇ�͋󕶎���Ԃ�
'
Private Function convFrcPrtOfDecPntToFrcPrtOfNPnt(ByVal frcPt As String, ByVal numOfDigits As Long, ByVal radix As Byte) As String
    
    Dim stringBuilder() As String '�ϊ����ʊi�[�p
    Dim repTimes As Long
    
    '
    '�L��10�i�������񂩂ǂ����̓`�F�b�N���Ȃ�
    '
    
    If (frcPt = "") Or (numOfDigits = 0) Then '�󕶎��w�肩�A���߂錅��=0
        convFrcPrtOfDecPntToFrcPrtOfNPnt = "" '�󕶎���ԋp
        Exit Function
        
    End If
    
    '�E����0����菜��
    Do While Right(frcPt, 1) = "0"
        frcPt = Left(frcPt, Len(frcPt) - 1)
        
    Loop
    
    If (frcPt = "") Then '�S��"0"��������
        convFrcPrtOfDecPntToFrcPrtOfNPnt = "0"
        Exit Function
        
    End If
    
    '�|���Z����
    repTimes = 0
    sizeOfStringBuilder = 0
    Do
        tmp = multiple(frcPt, CStr(radix), 10)
        
        ReDim Preserve stringBuilder(repTimes) '�̈�g��
        
        lenDiff = Len(tmp) - Len(frcPt)
        
        If (lenDiff > 0) Then '���オ�肪���������ꍇ
            stringBuilder(repTimes) = convByteToNChar(Left(tmp, lenDiff))
            frcPt = Right(tmp, Len(tmp) - lenDiff)
            
        Else '���オ�肪�������Ȃ������ꍇ
            stringBuilder(repTimes) = "0"
            frcPt = tmp
        
        End If
        
        '�E���̕s�v��"0"����菜��
        Do While (Right(frcPt, 1) = "0")
            frcPt = Left(frcPt, Len(frcPt) - 1)
            
        Loop
        
        If frcPt = "" Then '�S��"0"��������
            Exit Do
            
        End If
        
        repTimes = repTimes + 1
        
    Loop While IIf(numOfDigits < 0, True, (repTimes < numOfDigits)) '�J��Ԃ��񐔈ȉ�
    
    convFrcPrtOfDecPntToFrcPrtOfNPnt = Join(stringBuilder, vbNullString) '������A��
    
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
            decStr = add(nPowerOf2, decStr, 10, True)
        End If
        
        nPowerOf2 = multiple(nPowerOf2, "2", 10)
        
    Next cnt
    
    convIntPrtOfBinToIntPrtOfDecPrt = decStr
    
End Function

'
'2�i������������10�i���������ɕϊ�����
'
Private Function convFrcPrtOfBinPntToFrcPrtOfDecPnt(ByVal frcPt As String) As String
    
    Dim lpMx As Long
    Dim decStr As String
    
    Dim minusNpowerOf2 As String
    
    Dim rm As String
    
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
            decStr = add(minusNpowerOf2, decStr, 10, False)
        End If
        
        minusNpowerOf2 = divideByOneDig(minusNpowerOf2, "2", rm, 1, 10)
        
    Next cnt
    
    convFrcPrtOfBinPntToFrcPrtOfDecPnt = decStr
    
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


'
'���l������10�i�l�ł�������Ԃ�
'
Private Function convNCharToByte(ByVal ch As String) As Byte
    
    Dim toRetByte As Byte
    Dim ascOfA As Integer
    Dim ascOfG As Integer
    
    '
    '�L�����l�������ǂ����̓`�F�b�N���Ȃ�
    '
    
    ascOfA = Asc("A")
    ascOfG = Asc("G")
    ascOfCh = Asc(ch)
    
    If (ascOfA <= ascOfCh) And (ascOfCh <= ascOfG) Then 'A~G�̏ꍇ
        toRetByte = 10 + (ascOfCh - ascOfA)
    
    Else '0~9�̏ꍇ
        toRetByte = CByte(ch)
    
    End If
    
    convNCharToByte = toRetByte
    
End Function

'
'10�i�l���琔�l������Ԃ�
'
Private Function convByteToNChar(ByVal byt As Byte) As String
    
    Dim toRetStr As String
    
    '
    '�L�����l�������ǂ����̓`�F�b�N���Ȃ�
    '
    
    If (byt > 9) Then 'A~G�̏ꍇ
        toRetStr = Chr((byt - 10) + Asc("A"))
    
    Else '0~9�̏ꍇ
        toRetStr = Chr(byt + Asc("0"))
    
    End If
    
    convByteToNChar = toRetStr
    
End Function



