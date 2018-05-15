Attribute VB_Name = "NumericString"
'定数

Public Const DOT As String = "." '小数点表記

'割り切れない数値に対して何回割り算するか
Const DEFAULT_DIV_TIMES_FOR_INDIVISIBLE As Long = 255

'10進→n進変換時の、小数部分の限界算出桁数
Const DEFAULT_FRC_DIGITS As Long = 255

'n進小数→10進小数変換時の変換精度
Const PREC_OF_CONV As Long = 255

'
'n進文字列を指定桁数分シフトする
'
'radix
'    基数(2~16まで)
'
Public Function shiftNumeralPnt(ByVal str As String, ByVal shift As Long, Optional radix As Byte = 10) As Variant
    
    Dim idxOfDot As Long
    Dim ret As Long
    Dim toRet As Variant
    Dim sign As String
    
    Dim intPt As String
    Dim prcPt As String
    
    '引数チェック
    If str = "" Then '空文字の場合
        shiftNumeralPnt = CVErr(xlErrValue) '#VALUE!を返す
        Exit Function
        
    End If
    
    '符号を取り除く
    If Left(str, 1) = "-" Then '(-)値の時
        str = Right(str, Len(str) - 1)
        sign = "-"
        
    Else
        sign = ""
    
    End If
    
    'チェック
    ret = checkNumeralPntStr(str, radix, idxOfDot)
    
    If (ret = (Len(str) + 1)) Then 'n進文字列だった場合
        
        If (idxOfDot <= Len(str)) Then '文字列中に小数点文字があった場合
            str = Replace(str, DOT, "")
            
        End If
        
        shift = idxOfDot + shift
        lenOfStr = Len(str)
        
        If (shift <= 1) Then '左側を0埋めをする
            intPt = "0"
            frcPt = String(1 - shift, "0") & str
            
        ElseIf (shift > (lenOfStr + 1)) Then '右側を0埋めする
            intPt = str & String(shift - lenOfStr - 1, "0")
            frcPt = ""
            
        ElseIf (shift = (lenOfStr + 1)) Then '小数点位置が記載不要の場合
            intPt = str
            frcPt = ""
            
        Else '文字列中に小数点を挿入する
            intPt = Left(str, shift - 1)
            frcPt = Right(str, lenOfStr - Len(intPt))
            
        End If
        
        '整数部の不要な"0"を取り除く
        If (intPt <> "0") Then
            Do While Left(intPt, 1) = "0"
                intPt = Right(intPt, Len(intPt) - 1)
                
            Loop
        End If
        
        '小数部の不要な"0"を取り除く
        Do While Right(frcPt, 1) = "0"
            frcPt = Left(frcPt, Len(frcPt) - 1)
            
        Loop
        
        toRet = sign & intPt & IIf(frcPt = "", "", DOT & frcPt)
    
    Else 'n進文字列で無かった場合
        toRet = CVErr(xlErrNum) '#NUM!を返す
    
    End If
    
    shiftNumeralPnt = toRet
    
End Function

'
'文字列がn進文字列かどうかを返す
'
'radix
'    基数(2~16まで)
'
Public Function isNumeralPnt(ByVal decStr As String, Optional radix As Byte = 10) As Boolean
    
    Dim idxOfDot As Long
    Dim ret As Long
    Dim toRet As Boolean
    
    'チェック
    ret = checkNumeralPntStr(decStr, radix, idxOfDot)
    
    If (ret = (Len(decStr) + 1)) Then 'n進文字列だった場合
        toRet = True
        
    Else
        toRet = False
    
    End If
    
    isNumeralPnt = toRet
    
    
End Function

'
'小数付き10進値から小数付きn進値に変換する
'
'numOfDigits
'    小数算出時の限界除算回数
'
Public Function convDecPntToNPnt(ByVal pntStr As String, ByVal radix As Byte, Optional numOfDigits As Long = DEFAULT_FRC_DIGITS) As Variant
    
    Dim intPtOfBefore As String '整数部
    Dim frcPtOfBefore As String '小数部
    
    Dim intPtOfAfter As String '整数部
    Dim frcPtOfAfter As String '小数部
    Dim isMinus As Boolean
    Dim sign As String '符号
    
    Dim ret As Long
    Dim idxOfDot As Long
    
    Dim retOfSeparatePnt As Variant
    
    '基数チェック
    If (radix < 2) Or (16 < radix) Then
        convDecPntToNPnt = CVErr(xlErrValue) '#VALUE!を返却
        Exit Function
        
    End If
    
    'pntStrの文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(pntStr, 10, intPtOfBefore, frcPtOfBefore, isMinus)
    
    If (retOfSeparatePnt <> 0) Then 'pntStrはn進値として不正
        convDecPntToNPnt = CVErr(xlErrValue) '#VALUE!を返却
        Exit Function
        
    End If
    
    '整数部の不要な0を取り除く
    Do While (Left(intPtOfBefore, 1) = "0")
        intPtOfBefore = Right(intPtOfBefore, Len(intPtOfBefore) - 1)
        
    Loop
    
    If (intPtOfBefore = "") Then '全部"0"だったら
        intPtOfBefore = "0"
        
    End If
    
    '少数部の不要な0を取り除く
    Do While (Right(frcPtOfBefore, 1) = "0")
        frcPtOfBefore = Left(frcPtOfBefore, Len(frcPtOfBefore) - 1)
        
    Loop
    
    'マイナス値チェック
    If (isMinus) Then
        sign = "-"
        
    Else
        sign = ""
        
    End If
    
    '整数部をn進変換
    intPtOfAfter = convIntPrtOfDecPntToIntPrtOfNPnt(intPtOfBefore, radix)
    
    '小数部をn進変換
    If (frcPtOfBefore = "") Then '小数部は存在しない場合
        frcPtOfAfter = ""
        
    Else '小数部が存在する場合
        frcPtOfAfter = convFrcPrtOfDecPntToFrcPrtOfNPnt(frcPtOfBefore, radix, numOfDigits)
        
        If (frcPtOfAfter <> "") Then
            frcPtOfAfter = DOT & frcPtOfAfter
            
        End If
        
    End If
    
    '文字列結合
    convDecPntToNPnt = sign & intPtOfAfter & frcPtOfAfter
    
End Function

'
'小数付きn進値から小数付き10進値に変換する
'
'radix
'    基数(2~16まで)
'
'precisionOfConv
'    変換精度
'    ex:)
'    【前提】0.01(3進数) = 0.111111111111..(10進数)←1の繰り返し
'    【実行方法】convNPntToDecPnt("0.01", 3, precOfConv)
'    【結果】
'            precOfConv=2で実行した場合: 返却値:0.11
'            precOfConv=3で実行した場合: 返却値:0.111
'
Public Function convNPntToDecPnt(ByVal pntStr As String, ByVal radix As Byte, Optional precisionOfConv As Long = PREC_OF_CONV) As Variant

    Dim intPtOfBefore As String '整数部
    Dim frcPtOfBefore As String '小数部
    
    Dim intPtOfAfter As String '整数部
    Dim frcPtOfAfter As String '小数部
    Dim isMinus As Boolean
    Dim sign As String '符号
    
    Dim ret As Long
    Dim idxOfDot As Long
    
    Dim retOfSeparatePnt As Variant
    
    Dim ansIT As String
    
    '基数チェック
    If (radix < 2) Or (16 < radix) Then
        convNPntToDecPnt = CVErr(xlErrValue) '#VALUE!を返却
        Exit Function
        
    End If
    
    '変換精度チェック
    If (precisionOfConv < 0) Then
        convNPntToDecPnt = CVErr(xlErrValue) '#VALUE!を返却
        Exit Function
        
    End If
    
    'pntStrの文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(pntStr, radix, intPtOfBefore, frcPtOfBefore, isMinus)
    
    If (retOfSeparatePnt <> 0) Then 'pntStrはn進値として不正
        convNPntToDecPnt = CVErr(xlErrValue) '#VALUE!を返却
        Exit Function
        
    End If
    
    '整数部の不要な0を取り除く
    Do While (Left(intPtOfBefore, 1) = "0")
        intPtOfBefore = Right(intPtOfBefore, Len(intPtOfBefore) - 1)
        
    Loop
    
    If (intPtOfBefore = "") Then '全部"0"だったら
        intPtOfBefore = "0"
        
    End If
    
    '少数部の不要な0を取り除く
    Do While (Right(frcPtOfBefore, 1) = "0")
        frcPtOfBefore = Left(frcPtOfBefore, Len(frcPtOfBefore) - 1)
        
    Loop
    
    If (precisionOfConv = 0) Then '小数点以下の求める精度が0桁の場合
        frcPtOfBefore = ""
        
    End If
    
    'マイナス値チェック
    If (isMinus) Then
        sign = "-"
        
    Else
        sign = ""
        
    End If
    
    '整数部変換
    intPtOfAfter = convIntPrtOfNPntToIntPrtOfDecPnt(intPtOfBefore, radix)
    
    '小数部をn進変換
    If (frcPtOfBefore = "") Then '小数部は存在しない場合
        frcPtOfAfter = ""
        
    Else '小数部が存在する場合
        frcPtOfAfter = convFrcPrtOfNPntToFrcPrtOfDecPnt(frcPtOfBefore, radix, precisionOfConv, ansIT)
        
        If (frcPtOfAfter <> "") Then
            frcPtOfAfter = DOT & frcPtOfAfter
            
        End If
        
    End If
    
    '文字列結合
    convNPntToDecPnt = sign & intPtOfAfter & frcPtOfAfter

End Function

'
'1の補数を得る
'
Public Function get1sComplement() As Variant
    
    'todo
    
End Function

'
'2数を加算する
'
'radix
'   2~16 のみ
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
    
    '基数チェック
    If (radix < 2) Or (16 < radix) Then
        addNPntNPnt = CVErr(xlErrValue) '#VALUE!を返却
        Exit Function
        
    End If
    
    'val1の文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(value1, radix, intPrtOfVal1, frcPrtOfVal1, isMinusOfVal1)
    
    If (retOfSeparatePnt <> 0) Then 'val1はn進値として不正
        addNPntNPnt = CVErr(xlErrValue) '#VALUE!を返却
        Exit Function
        
    End If
    
    'valwの文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(value2, radix, intPrtOfVal2, frcPrtOfVal2, isMinusOfVal2)
    
    If (retOfSeparatePnt <> 0) Then 'val2はn進値として不正
        addNPntNPnt = CVErr(xlErrValue) '#VALUE!を返却
        Exit Function
        
    End If
    
    '小数部の桁数合わせ
    lenOfVal1FrcPrt = Len(frcPrtOfVal1)
    lenOfVal2FrcPrt = Len(frcPrtOfVal2)
    If (lenOfVal1FrcPrt > lenOfVal2FrcPrt) Then
        frcPrtOfVal2 = frcPrtOfVal2 & String(lenOfVal1FrcPrt - lenOfVal2FrcPrt, "0")
        
    Else
        frcPrtOfVal1 = frcPrtOfVal1 & String(lenOfVal2FrcPrt - lenOfVal1FrcPrt, "0")
        
    End If
    
    tmpVal1 = intPrtOfVal1 & frcPrtOfVal1
    tmpVal2 = intPrtOfVal2 & frcPrtOfVal2
    
    '加算or減算
    If (isMinusOfVal1) Then 'value1はマイナス値
        If (isMinusOfVal2) Then 'value2はマイナス値
            tmpVal = add(tmpVal1, tmpVal2, radix)
            toRetSign = "-"
            
        Else 'value2はプラス値
            tmpVal = subtract(tmpVal1, tmpVal2, radix, subtractionWasMinus)
            If (subtractionWasMinus) Then
                toRetSign = "-"
            Else
                toRetSign = ""
            End If
        
        End If
        
    Else 'value1はプラス値
        If (isMinusOfVal2) Then 'value2はマイナス値
            tmpVal = subtract(tmpVal1, tmpVal2, radix, subtractionWasMinus)
            If (subtractionWasMinus) Then
                toRetSign = "-"
            Else
                toRetSign = ""
            End If
            
        Else 'value2はプラス値
            tmpVal = add(tmpVal1, tmpVal2, radix)
            toRetSign = ""
        
        End If
    
    End If
    
    '小数点復活
    intPrt = Left(tmpVal, Len(tmpVal) - Len(frcPrtOfVal1))
    frcPrt = Right(tmpVal, Len(frcPrtOfVal1))
    
    '小数部の不要"0"削除
    Do While (Right(frcPrt, 1) = "0")
        frcPrt = Left(frcPrt, Len(frcPrt) - 1)
    Loop
    
    addNPntNPnt = toRetSign & intPrt & IIf(frcPrt = "", "", DOT & frcPrt)

End Function

'
'1st引数を2nd引数で掛ける
'2nd引数は1~9のみ可
'
'radix
'    2~16 のみ
'
Public Function multipleNPntNPnt(ByVal multiplicand As String, ByVal multiplier As String, Optional radix As Byte = 10) As Variant

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
    
    '乗数の文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(multiplier, radix, intPrtOfMultiplier, frcPrtOfMultiplier, multiplierIsMinus)
    If (retOfSeparatePnt <> 0) Then 'n進値として不正
        multipleNPntNPnt = retOfSeparatePnt
        Exit Function
        
    End If
    
    '小数部が存在しない場合は、取り除く
    If (frcPrtOfMultiplier = "0") Then
        frcPrtOfMultiplier = ""
    End If
    
    '被乗数の文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(multiplicand, radix, intPrtOfMultiplicand, frcPrtOfMultiplicand, multiplicandIsMinus)
    If (retOfSeparatePnt <> 0) Then 'n進値として不正
        multipleNPntNPnt = retOfSeparatePnt
        Exit Function
        
    End If
    
    '小数部が存在しない場合は、取り除く
    If (frcPrtOfMultiplicand = "0") Then
        frcPrtOfMultiplicand = ""
    End If
    
    
    '乗算
    retOfMultiple = multiple(intPrtOfMultiplicand & frcPrtOfMultiplicand, intPrtOfMultiplier & frcPrtOfMultiplier, radix)
    
    intPrtOfAns = Left(retOfMultiple, Len(retOfMultiple) - (Len(frcPrtOfMultiplicand) + Len(frcPrtOfMultiplier)))
    frcPrtOfAns = Right(retOfMultiple, Len(frcPrtOfMultiplicand) + Len(frcPrtOfMultiplier))
    
    
    '整数部の不要な0を取り除く
    Do While (Left(intPrtOfAns, 1) = "0")
        intPrtOfAns = Right(intPrtOfAns, Len(intPrtOfAns) - 1)
    Loop
    If (intPrtOfAns = "") Then '全部0だったら
        intPrtOfAns = "0"
    End If
    
    '小数部の不要な0を取り除く
    Do While (Right(frcPrtOfAns, 1) = "0")
        frcPrtOfAns = Left(frcPrtOfAns, Len(frcPrtOfAns) - 1)
    Loop
    
    '符号判定
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
'1st引数を2nd引数で割る
'割り切れない時は、3rd引数で指定された回数だけ割った結果を返す
'3rd引数が(-)値の場合は際限なく割り続ける
'
'radix
'    基数(2~16まで)
'
Public Function divideNPntNPnt(ByVal dividend As String, ByVal divisor As String, Optional ByVal radix As Byte = 10, Optional ByVal limitOfRepTimes As Long = DEFAULT_DIV_TIMES_FOR_INDIVISIBLE) As Variant

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
    Dim errOfDvide As Variant
    
    '除数の文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(divisor, radix, intPrtOfDivisor, frcPrtOfDivisor, divisorIsMinus)
    If (retOfSeparatePnt <> 0) Then 'n進値として不正
        divideNPntNPnt = retOfSeparatePnt
        Exit Function
        
    End If
    
    '整数の不要な0を取り除く
    Do While (Left(intPrtOfDivisor, 1) = "0")
        intPrtOfDivisor = Right(intPrtOfDivisor, Len(intPrtOfDivisor) - 1)
        
    Loop
    
    If intPrtOfDivisor = "" Then '全部0だったら
        intPrtOfDivisor = "0"
        
    End If
    
    '小数の不要な0を取り除く
    Do While (Right(frcPrtOfDivisor, 1) = "0")
        frcPrtOfDivisor = Left(frcPrtOfDivisor, Len(frcPrtOfDivisor) - 1)
        
    Loop
    
    '0割チェック
    If (intPrtOfDivisor = "0") And (frcPrtOfDivisor = "") Then
        divideNPntNPnt = CVErr(xlErrDiv0) '#DIV0!を返す
        Exit Function
        
    End If
    
    '被除数の文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(dividend, radix, intPrtOfDividend, frcPrtOfDividend, dividendIsMinus)
    If (retOfSeparatePnt <> 0) Then 'n進値として不正
        divideNPntNPnt = retOfSeparatePnt
        Exit Function
        
    End If
    
    '整数の不要な0を取り除く
    Do While (Left(intPrtOfDividend, 1) = "0")
        intPrtOfDividend = Right(intPrtOfDividend, Len(intPrtOfDividend) - 1)
        
    Loop
    
    If intPrtOfDividend = "" Then '全部0だったら
        intPrtOfDividend = "0"
        
    End If
    
    '小数の不要な0を取り除く
    Do While (Right(frcPrtOfDividend, 1) = "0")
        frcPrtOfDividend = Left(frcPrtOfDividend, Len(frcPrtOfDividend) - 1)
        
    Loop
    
    
    '除算
    retOfDivide = divide(intPrtOfDividend & frcPrtOfDividend, intPrtOfDivisor & frcPrtOfDivisor, radix, limitOfRepTimes, rm, errOfDvide)
    
    If (retOfDivide = "") Then 'オーバーフローの場合
        divideNPntNPnt = errOfDvide
        Exit Function
        
    End If
    
    digitsOfIntPrtOfAns = Len(intPrtOfDividend) + Len(frcPrtOfDivisor)
    lenOfRetOfDivide = Len(retOfDivide)
    
    If (lenOfRetOfDivide < digitsOfIntPrtOfAns) Then
        intPrtOfAns = retOfDivide & String(digitsOfIntPrtOfAns - lenOfRetOfDivide, "0") '整数部の生成
        frcPrtOfAns = "0"
        
    Else
        intPrtOfAns = Left(retOfDivide, digitsOfIntPrtOfAns) '小数部の生成
        frcPrtOfAns = Right(retOfDivide, lenOfRetOfDivide - digitsOfIntPrtOfAns)
        
    End If
    
    '整数部の不要な0を取り除く
    Do While (Left(intPrtOfAns, 1) = "0")
        intPrtOfAns = Right(intPrtOfAns, Len(intPrtOfAns) - 1)
    Loop
    
    If (intPrtOfAns = "") Then '全部0だったら
        intPrtOfAns = "0"
    End If
    
    '小数部の不要な0を取り除く
    Do While (Right(frcPrtOfAns, 1) = "0")
        frcPrtOfAns = Left(frcPrtOfAns, Len(frcPrtOfAns) - 1)
    Loop
    
    '符号判定
    If (dividendIsMinus Xor divisorIsMinus) Then
        
        If (intPrtOfAns = "0") And (frcPrtOfAns = "") Then
            signOfAns = ""
            
        Else
            signOfAns = "-"
            
        End If
    Else
        signOfAns = ""
    
    End If
    
    divideNPntNPnt = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)
    
End Function

'
'小数部と整数部に分解する
'
'成功の場合は0を返却する
'失敗の場合はCvErrを返却する
'
Private Function separateToIntAndFrc(ByVal pnt As String, ByVal radix As Byte, ByRef intPrt As String, ByRef frcPrt, ByRef isMinus As Boolean) As Variant
    
    Dim idxOfDot As Long
    
    Dim value1IsMinus
    
    Dim retToIsMinus As Boolean
    Dim retToIntPrt As String
    Dim retToFrcPrt As String
    
    '字列長チェック
    If (Len(pnt) < 1) Then '文字列長が0
        separateToIntAndFrc = CVErr(xlErrValue) '#VALUE!を返す
        Exit Function
        
    End If
    
    '符号を取り除く
    If (Left(pnt, 1) = "-") Then
        retToIsMinus = True
        pnt = Right(pnt, Len(pnt) - 1)
        If (pnt = "") Then
            pnt = "0"
        End If
        
    Else
        retToIsMinus = False
        
    End If
    
    'n進値として正しいかチェック
    ret = checkNumeralPntStr(pnt, radix, idxOfDot)
    
    If (ret <> (Len(pnt) + 1)) Then 'value1はn進値として不正
        separateToIntAndFrc = CVErr(xlErrNum) '#NUM!を返す
        Exit Function
        
    End If
    
    '整数部と小数部に分ける
    
    '整数部を抽出する
    retToIntPrt = Left(pnt, idxOfDot - 1)
    If (retToIntPrt = "") Then '整数部の記載がなかった場合
        retToIntPrt = "0"
    End If
    
    '小数部を抽出する
    If (idxOfDot < Len(pnt)) Then '小数部の記載がある
        retToFrcPrt = Right(pnt, Len(pnt) - idxOfDot)
        
    Else '小数部の記載がない
        retToFrcPrt = "0"
    
    End If
    
    intPrt = retToIntPrt
    frcPrt = retToFrcPrt
    isMinus = retToIsMinus
    separateToIntAndFrc = 0
    
End Function

'
'文字列がn進値文字列かどうかチェックする
'
'返却値
'    n進値文字列だったの場合は文字長 + 1
'    そうでない場合は、最初に見つかった10進文字以外の文字位置
'    空文字を指定された場合は0を返す
'
'idxOfDot
'    小数点文字位置
'    小数点が無かった場合は最終文字位置+1を格納する
'
'radix
'    基数(2~16のみ)
'
Private Function checkNumeralPntStr(ByVal decStr As String, ByVal radix As Byte, ByRef idxOfDot As Long) As Long
    
    Dim foundIdxOfDot As Long '小数点文字が最初に見つかった文字位置
    Dim cnt As Long
    Dim lpMx As Long
    Dim stCnt As Long
    
    Dim minOkChar1 As Integer
    Dim maxOkChar1 As Integer
    Dim minOkChar2 As Integer
    Dim maxOkChar2 As Integer
    
    Dim radixIsBiggerThan10 As Boolean
    
    lpMx = Len(decStr)
    
    '引数チェック
    If (lpMx = 0) Or (radix < 2) Or (16 < radix) Then
        checkNumeralPntStr = 0
        Exit Function
        
    End If
    
    '基数からOKな文字コード範囲を作る
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
    
    If (Left(decStr, 1) = "-") Then '最初の(-)符号は無視する
        stCnt = 2
        
    Else
        stCnt = 1
        
    End If
    
    For cnt = stCnt To lpMx
        
        ch = Mid(decStr, cnt, 1)
        chCode = Asc(ch)
        
        If (chCode < minOkChar1) Or (maxOkChar1 < chCode) Then  '文字は0~9いずれでもない
            If IIf(radixIsBiggerThan10, (chCode < minOkChar2) Or (maxOkChar2 < chCode), True) Then '文字はA~Fいずれでもない
            
                If (ch = DOT) Then
                    If (foundIdxOfDot = 0) Then '小数点文字は1回目
                        foundIdxOfDot = cnt
                    
                    Else '小数点文字は2回目
                        Exit For
                        
                    End If
                
                Else '文字は0~9いずれでもなく、小数点文字でもない
                    Exit For
                    
                End If
            End If
        End If
    Next cnt
    
    
    If foundIdxOfDot = 0 Then '小数点文字が見つからなかった場合
        idxOfDot = lpMx + 1
        
    Else '小数点文字が見つかった場合
        idxOfDot = foundIdxOfDot
    
    End If
    
    checkNumeralPntStr = cnt
    
End Function

'
'2数を和算する
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
    '有効な10進値であるかはチェックしない
    '
    
    '文字列長確認
    lenOfVal1 = Len(val1)
    lenOfVal2 = Len(val2)
    
    
    '0埋め確認
    If (lenOfVal1 > lenOfVal2) Then
        
        If fill0Left Then '左を0埋めする
            val2 = String(lenOfVal1 - lenOfVal2, "0") & val2
            
        Else
            val2 = val2 & String(lenOfVal1 - lenOfVal2, "0")
            
        End If
        
        lenOfVal2 = lenOfVal1
        
    Else
        If fill0Left Then '左を0埋めする
            val1 = String(lenOfVal2 - lenOfVal1, "0") & val1
            
        Else
            val1 = val1 & String(lenOfVal2 - lenOfVal1, "0")
            
        End If
        
        lenOfVal1 = lenOfVal2
        
    End If
    
    ReDim stringBuilder(lenOfVal1 - 1) '領域拡張
    
    carrier = 0
    
    'additionループ
    For idxOfVal = lenOfVal1 To 1 Step -1
        
        tmpDigitOfVal1 = convNCharToByte(Mid(val1, idxOfVal, 1))
        tmpDigitOfVal2 = convNCharToByte(Mid(val2, idxOfVal, 1))
        
        tmpStr = convIntPrtOfDecPntToIntPrtOfNPnt(tmpDigitOfVal1 + tmpDigitOfVal2 + carrier, radix)
        
        If (Len(tmpStr) = 2) Then '桁が増えたか
            carrier = 1
            
        Else
            carrier = 0
            
        End If
        
        stringBuilder(idxOfVal - 1) = Right(tmpStr, 1)
        
    Next idxOfVal
    
    ret = Join(stringBuilder, vbNullString)
    
    If (carrier > 0) Then '桁増えたか
        ret = CInt(carrier) & ret
        
    End If
    
    add = ret
    
End Function

'
'val1からval2を減算する
'
Private Function subtract(ByVal val1 As String, ByVal val2 As String, ByVal radix As Byte, ByRef resultIsMinus As Boolean) As String
    
    '変数宣言
    Dim lenOfVal1 As Long
    Dim lenOfVal2 As Long
    
    Dim val1IsLarger As Integer '0:不明, 1:yes, -1:no
    
    Dim idxOfVal As Long
    Dim idxMxOfVal As Long
    
    Dim stringBuilder() As String
    
    Dim wasMinus As Boolean
    
    '
    '有効な10進値かどうかはチェックしない
    '
    
    
    '文字列長確認
    lenOfVal1 = Len(val1)
    lenOfVal2 = Len(val2)
    
    '0埋め確認
    If (lenOfVal1 > lenOfVal2) Then
        val2 = String(lenOfVal1 - lenOfVal2, "0") & val2
        
    Else
        val1 = String(lenOfVal2 - lenOfVal1, "0") & val1
        
    End If
    
    
    '大小比較チェック
    idxOfVal = 1
    val1IsLarger = 0
    idxMxOfVal = Len(val1)
    Do
        val1Digit = convNCharToByte(Mid(val1, idxOfVal, 1))
        val2Digit = convNCharToByte(Mid(val2, idxOfVal, 1))
        
        'どちらかが大きかったら break
        If val1Digit > val2Digit Then
            val1IsLarger = 1
            Exit Do
        
        ElseIf val1Digit < val2Digit Then
            val1IsLarger = -1
            Exit Do
        
        End If
        
        idxOfVal = idxOfVal + 1
        
    Loop While idxOfVal <= idxMxOfVal
    
    
    If (val1IsLarger = 0) Then  '2数は同じ数値
        subtract = String(idxMxOfVal, "0")
        resultIsMinus = False
        Exit Function
        
    End If
    
    ReDim stringBuilder(idxMxOfVal - 1) '領域拡張
    
    If (val1IsLarger = -1) Then 'val2の方が大きい数値だったら
        '2数を入れ替える
        buf = val1
        val1 = val2
        val2 = buf
        
        wasMinus = True
        
    Else
        wasMinus = False
        
    End If
    
    '減算ループ
    carrier = 0
    For idxOfVal = idxMxOfVal To 1 Step -1
        
        val1Digit = convNCharToByte(Mid(val1, idxOfVal, 1))
        val2Digit = convNCharToByte(Mid(val2, idxOfVal, 1))
        
        '繰り下がりチェック
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
    resultIsMinus = wasMinus
    
    
End Function

'
'乗算をする
'
Private Function multiple(ByVal multiplicand As String, ByVal multiplier As String, ByVal radix As Byte) As String

    Dim ansOfMultipleByOneDig As String
    Dim numOf0 As Long
    Dim tmpAns As String
    
    '
    '有効なn進値であるかはチェックしない
    '
    
    'multiplierの不要な0を取り除く
    Do While (Left(multiplier, 1) = "0")
        multiplier = Right(multiplier, Len(multiplier) - 1)
        
    Loop
    
    If (multiplier = "") Then '全部"0"だったら
        multiple = String(Len(multiplicand), "0")
        Exit Function
        
    ElseIf (multiplier = "1") Then '1掛けの場合はそのまま返す
        multiple = multiplicand
        Exit Function
        
    End If
    
    numOf0 = 0
    tmpAns = "0"
    
    '乗算ループ
    For idx = Len(multiplier) To 1 Step -1
        
        ansOfMultipleByOneDig = multipleByOneDig(multiplicand, Mid(multiplier, idx, 1), radix)
        tmpAns = add(tmpAns, ansOfMultipleByOneDig & String(numOf0, "0"), radix)
        
        numOf0 = numOf0 + 1
        
    Next idx
    
    multiple = tmpAns
    
End Function

'
'1桁数値による乗算をする
'
Private Function multipleByOneDig(ByVal multiplicand As String, ByVal multiplierCh As String, ByVal radix As Byte) As String

    Dim carrier As Byte
    Dim digitOfMultiplicand As Byte
    Dim multiplier As Byte
    Dim digitIdxOfMultiplicand As Long
    Dim stringBuilder() As String '割り算結果格納用
    Dim idxOfStringBuilder As Long
    
    '
    'multiplicandが有効なn進値であるかはチェックしない
    'multiplierChは1桁であることはチェックしない
    '
    
    If (multiplierCh = "0") Then '0掛けの場合は0を返す
        multipleByOneDig = String(Len(multiplicand), "0")
        Exit Function
    
    ElseIf (multiplierCh = "1") Then '1掛けの場合はそのまま返す
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
        
        'carrier判定
        If (Len(tmpStr) = 2) Then '桁が増えた場合
            carrier = convNCharToByte(Left(tmpStr, 1))
            
        Else '桁が増えなかった場合
            carrier = 0
            
        End If
        
        ReDim Preserve stringBuilder(idxOfStringBuilder) '領域拡張
        stringBuilder(idxOfStringBuilder) = Right(tmpStr, 1)
        
        digitIdxOfMultiplicand = digitIdxOfMultiplicand - 1
        idxOfStringBuilder = idxOfStringBuilder + 1
    
    Loop While digitIdxOfMultiplicand > 0 '被乗数が残っている間
    
    '桁上がりチェック
    If (carrier > 0) Then
        ReDim Preserve stringBuilder(idxOfStringBuilder) '領域拡張
        stringBuilder(idxOfStringBuilder) = convIntPrtOfDecPntToIntPrtOfNPnt(carrier, radix)
        
    End If
    
    multipleByOneDig = Join(invertStringArray(stringBuilder), vbNullString) '文字列連結
    
End Function

'
'除算をする
'
'以下の場合は空文字を返却し、
'errCodeにエラーコードを格納する
'    ┣0割の場合。(エラーコードは#DIV/0!)
'    ┗dividend / divisor にlong型で取り扱えない大きな数値がある場合。(エラーコードは#NUM!)
'
'remainder
'    剰余。
'    小数点以下となる場合は、
'    一番左を1桁目として小数点を取り除いた小数文字列となる。
'    ex:)
'    【前提】10 / 8 = 1.2 余り 0.4
'    【実行方法】x = divide("10", "8", 10, 1, rm, code)
'    【結果】 x:012
'            rm:04
'
'limitOfRepTimes:
'    Indivisible Numberに対するdivide回数制限
'    (-)値を設定した場合は、無限に割り続ける
'
Private Function divide(ByVal dividend As String, ByVal divisor As String, ByVal radix As Byte, ByVal limitOfRepTimes As Long, ByRef remainder As String, ByRef errCode As Variant) As String

    '変数宣言
    Dim quot As Long '商
    Dim rmnd As Long '余り
    
    Dim repTimes As Long 'IndivisibleNumberに対するdivide回数
    
    Dim digitOfDividend As Long '一時被除数
    
    Dim stringBuilder() As String '商格納用
    Dim stringBuilderRM() As String '剰余格納用
    Dim digitIdxOfDividend As Long 'Division結果文字列長
    
    Dim divisorDec As Long
    
    Dim dividingFrc As Boolean
    
    '
    'dividend, divisor が有効なn進値であるかはチェックしない
    '
    
    'divisorの不要な0を取り除く
    Do While (Left(divisor, 1) = "0")
        divisor = Right(divisor, Len(divisor) - 1)
        
    Loop
    
    If divisor = "" Then '全部0だったら
        divide = ""
        errCode = CVErr(xlErrNum) '#NUM!を返す
        Exit Function
        
    End If
    
    'divisorの10進変換
    tmp = convIntPrtOfNPntToIntPrtOfDecPnt(divisor, radix)
    
    'divisorのLong型変換
    On Error GoTo OVERFLOW
    divisorDec = CLng(tmp)
    
    '1割チェック
    If divisorDec = 1 Then
        divide = dividend '1割の場合はそのまま返す
        Exit Function
        
    ElseIf divisorDec = 0 Then '0割チェック
        divide = ""
        errCode = CVErr(xlErrDiv0) '#DIV0!を返す
        Exit Function
        
    End If
    
    '初期化
    rmnd = 0
    digitIdxOfDividend = 1
    repTimes = 0
    
    dividingFrc = False '小数点以下に対する割り算に突入したか
    
    '実行ループ
    Do
        digitOfDividend = rmnd * radix + convNCharToByte(Mid(dividend, digitIdxOfDividend, 1)) '上位桁の余り & 該当桁
        
        quot = digitOfDividend \ divisorDec '商
        rmnd = digitOfDividend Mod divisorDec '余り
        
        ReDim Preserve stringBuilder(digitIdxOfDividend - 1) '領域拡張
        stringBuilder(digitIdxOfDividend - 1) = convIntPrtOfDecPntToIntPrtOfNPnt(quot, radix) '商を追記
        
        digitIdxOfDividend = digitIdxOfDividend + 1
        
        If (rmnd > 0) And (Len(dividend) < digitIdxOfDividend) Then '余りがあるけれど、次の桁が無い
        
            If (limitOfRepTimes > -1) And (repTimes < limitOfRepTimes) Then '再帰計算回数が指定回数以下
                dividend = dividend & "0" '"0"を付加
                
                ReDim Preserve stringBuilderRM(repTimes) '領域拡張
                stringBuilderRM(repTimes) = "0"
                
                repTimes = repTimes + 1
                
            End If
            
        End If
        
    Loop While digitIdxOfDividend <= Len(dividend) '最終文字に到達しない間
    
    If (rmnd = 0) Then '余りが0のとき
        remainder = "0"
        
    Else '余りが存在する時
        
        ReDim Preserve stringBuilderRM(repTimes) '領域拡張
        stringBuilderRM(repTimes) = convIntPrtOfDecPntToIntPrtOfNPnt(rmnd, radix)
        remainder = Join(stringBuilderRM, vbNullString) '文字列連結
    
    End If
    
    divide = Join(stringBuilder, vbNullString) '文字列連結
    
    Exit Function
    
OVERFLOW: 'オーバーフローの場合
    divide = ""
    errCode = CVErr(xlErrNum) '#NUM!を返す
    Exit Function
    
End Function

'
'10進整数部をn進整数部に変換する
'
'radix:
'    2~16 のみ
'
Private Function convIntPrtOfDecPntToIntPrtOfNPnt(ByVal decInt As String, ByVal radix As Byte) As String
    
    Dim stringBuilder() As String '変換後文字列生成用
    Dim sizeOfStringBuilder As Long
    Dim rm As String
    Dim errOfDvide As Variant
    
    '
    '有効10進数値文字列かどうかはチェックしない
    '
    
    If (decInt = "") Then
        convIntPrtOfDecPntToIntPrtOfNPnt = "0"
        Exit Function
        
    End If
    
    '左側の不要な"0"を取り除く
    Do While Left(decInt, 1) = "0"
        decInt = Right(decInt, Len(decInt) - 1)
        
    Loop
    
    If (decInt = "") Then '全部"0"だったら
        convIntPrtOfDecPntToIntPrtOfNPnt = "0"
        Exit Function
        
    End If
    
    '10進→10進変換だったら
    If (radix = 10) Then
        convIntPrtOfDecPntToIntPrtOfNPnt = decInt '変換せずに返す
        Exit Function
        
    End If
    
    sizeOfStringBuilder = 0
    strLenOfRadix = Len(CStr(radix))
    
    '字列生成
    Do While True
        
        If (Len(decInt) <= strLenOfRadix) Then
            
            If (CByte(decInt) < radix) Then '基数で割れる数がなくなった
                Exit Do
                
            End If
        End If
        
        decInt = divide(decInt, radix, 10, 0, rm, errOfDvide)
        'オーバーフローは発生し得ない
        
        '左側の不要な"0"を取り除く
        Do While Left(decInt, 1) = "0"
            decInt = Right(decInt, Len(decInt) - 1)
            
        Loop
        If (decInt = "") Then '全部"0"だったら
            decInt = "0"
            
        End If
        
        ReDim Preserve stringBuilder(sizeOfStringBuilder) '領域拡張
        stringBuilder(sizeOfStringBuilder) = convByteToNChar(rm)
        
        sizeOfStringBuilder = sizeOfStringBuilder + 1
        
    Loop
    
    '最上位Bit付加
    ReDim Preserve stringBuilder(sizeOfStringBuilder) '領域拡張
    stringBuilder(sizeOfStringBuilder) = convByteToNChar(decInt)
    
    convIntPrtOfDecPntToIntPrtOfNPnt = Join(invertStringArray(stringBuilder), vbNullString) '文字列連結
    
End Function

'
'10進小数部分からn進小数部分に変換する
'
'numOfDigits:
'    求める小数点以下の桁数
'    0を指定した場合は空文字を返す
'
Private Function convFrcPrtOfDecPntToFrcPrtOfNPnt(ByVal frcPt As String, ByVal radix As Byte, ByVal numOfDigits As Long) As String
    
    Dim stringBuilder() As String '変換結果格納用
    Dim repTimes As Long
    
    '
    '有効10進数文字列かどうかはチェックしない
    '
    
    If (frcPt = "") Or (numOfDigits = 0) Then '空文字指定か、求める桁数=0
        convFrcPrtOfDecPntToFrcPrtOfNPnt = "" '空文字を返却
        Exit Function
        
    End If
    
    '右側の0を取り除く
    Do While Right(frcPt, 1) = "0"
        frcPt = Left(frcPt, Len(frcPt) - 1)
        
    Loop
    
    If (frcPt = "") Then '全部"0"だったら
        convFrcPrtOfDecPntToFrcPrtOfNPnt = "0"
        Exit Function
        
    End If
    
    '10進→10進変換だったら
    If (radix = 10) Then
        convFrcPrtOfDecPntToFrcPrtOfNPnt = frcPt
        Exit Function
        
    End If
    
    '字列生成ループ
    repTimes = 0
    sizeOfStringBuilder = 0
    Do
        tmp = multiple(frcPt, CStr(radix), 10)
        
        ReDim Preserve stringBuilder(repTimes) '領域拡張
        
        lenDiff = Len(tmp) - Len(frcPt)
        
        If (lenDiff > 0) Then '桁上がりが発生した場合
            stringBuilder(repTimes) = convByteToNChar(Left(tmp, lenDiff))
            frcPt = Right(tmp, Len(tmp) - lenDiff)
            
        Else '桁上がりが発生しなかった場合
            stringBuilder(repTimes) = "0"
            frcPt = tmp
        
        End If
        
        '右側の不要な"0"を取り除く
        Do While (Right(frcPt, 1) = "0")
            frcPt = Left(frcPt, Len(frcPt) - 1)
            
        Loop
        
        If frcPt = "" Then '全部"0"だったら
            Exit Do
            
        End If
        
        repTimes = repTimes + 1
        
    Loop While IIf(numOfDigits < 0, True, (repTimes < numOfDigits)) '繰り返し回数以下
    
    convFrcPrtOfDecPntToFrcPrtOfNPnt = Join(stringBuilder, vbNullString) '文字列連結
    
End Function

'
'n進整数部分から10進整数部分に変換する
'
Private Function convIntPrtOfNPntToIntPrtOfDecPnt(ByVal intPt As String, ByVal radix As Byte) As String
    
    Dim xPowerOfRadix As String
    Dim decStr As String
    
    '
    '有効n進文字列かどうかはチェックしない
    '
    
    '引数チェック
    If (intPt = "") Then '空文字指定の場合
        convIntPrtOfNPntToIntPrtOfDecPnt = "0" '"0"を返す
        Exit Function
        
    End If
    
    '10進→10進変換だったら
    If (radix = 10) Then
        convIntPrtOfNPntToIntPrtOfDecPnt = intPt
        Exit Function
        
    End If
    
    strOfRadix = CStr(radix)
    xPowerOfRadix = "1"
    decStr = "0"
    
    For cnt = Len(intPt) To 1 Step -1
        ch = Mid(intPt, cnt, 1)
        
        If (ch <> "0") Then
            tmp = multiple(xPowerOfRadix, CStr(convNCharToByte(ch)), 10)
            decStr = add(decStr, tmp, 10, True)
            
        End If
        
        xPowerOfRadix = multiple(xPowerOfRadix, strOfRadix, 10)
        
    Next cnt
    
    convIntPrtOfNPntToIntPrtOfDecPnt = decStr
    
End Function

'
'n進小数部分から10進小数部分に変換する
'
'numOfSignificantDigits
'    有効桁数
'
'precisionOfConv
'    変換精度
'    ex:)
'    【前提】0.1(3進数) = 0.33333333333333..(10進数)←3の繰り返し
'    【実行方法】convFrcPrtOfNPntToFrcPrtOfDecPnt("1", 3, precOfConv, ansIT)
'    【結果】
'            precOfConv=2で実行した場合: 返却値:33
'            precOfConv=3で実行した場合: 返却値:333
'
'ansIncTrunc(※参照渡し)
'    変換誤差を含めた最大の解
'    ex:)
'    【前提】0.01(3進数) = 0.1111111111111..(10進数)←1の繰り返し
'    【実行方法】convFrcPrtOfNPntToFrcPrtOfDecPnt("01", 3, 2, ansIT)
'    【結果】返却値:11
'            ansIT :1134
'       ※解は0.11以上、0.1134未満である事を表す
'
Private Function convFrcPrtOfNPntToFrcPrtOfDecPnt(ByVal frcPt As String, ByVal radix As Byte, ByVal precisionOfConv As Long, ByRef ansIncTrunc As String) As String
    
    Dim lpMx As Long
    Dim decStr As String
    
    Dim minusXpowerOfRadix As String
    Dim minusXpowerOfRadixT As String
    
    Dim rm As String
    Dim errOfDvide As Variant
    Dim strOfRadix As String
    
    Dim trunc As String '誤差
    
    '
    '有効n進文字列かどうかはチェックしない
    '変換精度がマイナスかどうかはチェックしない
    '
    
    '引数チェック
    If (frcPt = "") Then '空文字指定の場合
        convFrcPrtOfNPntToFrcPrtOfDecPnt = "0" '"0"を返す
        Exit Function
        
    End If
    
    strOfRadix = CStr(radix)
    lpMx = Len(frcPt)
    decStr = "0"
    trunc = "0"
    
    tmp = divide("1", strOfRadix, 10, precisionOfConv, rm, errOfDivide)
    'オーバーフローは発生し得ない
    minusXpowerOfRadix = Right(tmp, Len(tmp) - 1) '一番左の0を取り除く
    
    If (rm <> "0") Then
        rm = Right(rm, Len(rm) - 1) '小数点以下部分のみにする
        
    End If
    
    minusXpowerOfRadixT = add(minusXpowerOfRadix, rm, 10, True)
    
    '生成ループ
    For cnt = 1 To lpMx
        tmpCh = Mid(frcPt, cnt, 1)
        
        If (tmpCh <> "0") Then
            decOfTmpCh = convNCharToByte(tmpCh)
            toAdd = multiple(minusXpowerOfRadix, decOfTmpCh, 10)
            decStr = add(toAdd, decStr, 10, False)
            
            toAdd = multiple(minusXpowerOfRadixT, decOfTmpCh, 10)
            trunc = add(toAdd, trunc, 10, False)
            
        End If
        
        minusXpowerOfRadix = divide(minusXpowerOfRadix, strOfRadix, 10, precisionOfConv, rm, errOfDvide)
        'オーバーフローは発生し得ない
        
        minusXpowerOfRadixT = divide(minusXpowerOfRadixT, strOfRadix, 10, precisionOfConv, rm, errOfDvide)
        'オーバーフローは発生し得ない
        
        minusXpowerOfRadixT = add(minusXpowerOfRadixT, rm, 10, True)
        
    Next cnt
    
    ansIncTrunc = trunc
    convFrcPrtOfNPntToFrcPrtOfDecPnt = decStr
    
End Function

'
'String配列の順番を入替える
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
'数値文字が10進値でいくつかを返す
'
Private Function convNCharToByte(ByVal ch As String) As Byte
    
    Dim toRetByte As Byte
    Dim ascOfA As Integer
    Dim ascOfG As Integer
    
    '
    '有効数値文字かどうかはチェックしない
    '
    
    ascOfA = Asc("A")
    ascOfG = Asc("G")
    ascOfCh = Asc(ch)
    
    If (ascOfA <= ascOfCh) And (ascOfCh <= ascOfG) Then 'A~Gの場合
        toRetByte = 10 + (ascOfCh - ascOfA)
    
    Else '0~9の場合
        toRetByte = CByte(ch)
    
    End If
    
    convNCharToByte = toRetByte
    
End Function

'
'10進値から数値文字を返す
'
Private Function convByteToNChar(ByVal byt As Byte) As String
    
    Dim toRetStr As String
    
    '
    '有効数値文字かどうかはチェックしない
    '
    
    If (byt > 9) Then 'A~Gの場合
        toRetStr = Chr((byt - 10) + Asc("A"))
    
    Else '0~9の場合
        toRetStr = Chr(byt + Asc("0"))
    
    End If
    
    convByteToNChar = toRetStr
    
End Function


