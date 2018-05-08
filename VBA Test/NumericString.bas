Attribute VB_Name = "NumericString"
'wish
'n進数計算ができるようにしたい

Public Const DOT As String = "." '小数点表記

'割り切れない数値に対して何回割り算するか
Const DEFAULT_DIV_TIMES_FOR_INDIVISIBLE As Long = 255

'2進小数部分の算出桁数
Const DEFAULT_FRC_DIGITS As Long = 255

'
'n進文字列を指定桁数分シフトする
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
    ret = checkNumeralPntStr(str, idxOfDot, radix)
    
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
    ret = checkNumeralPntStr(decStr, idxOfDot, radix)
    
    If (ret = (Len(decStr) + 1)) Then 'n進文字列だった場合
        toRet = True
        
    Else
        toRet = False
    
    End If
    
    isNumeralPnt = toRet
    
    
End Function

'
'小数付き10進値から小数付き2進値に変換する
'
'numOfDigits
'    小数算出時の限界除算回数
'
Public Function convDecPntToBinPnt(ByVal decStr As String, Optional numOfDigits As Long = DEFAULT_FRC_DIGITS) As Variant
    
    convDecPntToBinPnt = convNPntToNPnt(decStr, 10, numOfDigits)
    
End Function

'
'小数付き2進値から小数付き10進値に変換する
'
Public Function convBinPntToDecPnt(ByVal binStr As String) As Variant
    
    convBinPntToDecPnt = convNPntToNPnt(binStr, 2, -1)
    
End Function

'
'10進⇔2進変換用共通関数
'
Private Function convNPntToNPnt(ByVal pntStr As String, ByVal radix As Byte, numOfDigits As Long) As Variant

    Dim intPtOfBefore As String '整数部
    Dim frcPtOfBefore As String '小数部
    
    Dim intPtOfAfter As String '整数部
    Dim frcPtOfAfter As String '小数部
    Dim sign As String '符号
    
    Dim ret As Long
    Dim idxOfDot As Long
    
    '引数チェック
    If (Len(pntStr) = 0) Then
        convNPntToNPnt = CVErr(xlErrValue) '#VALUE!を返却
        Exit Function
        
    End If
    
    'マイナス値チェック
    If (Left(pntStr, 1) = "-") Then
        If (Len(pntStr) < 2) Then
            convNPntToNPnt = CVErr(xlErrNum) '#NUM!を返却
            Exit Function
        End If
        
        sign = "-"
        pntStr = Right(pntStr, Len(pntStr) - 1)
        
    Else
        sign = ""
        
    End If
    
    '数値列かどうかチェック
    ret = checkNumeralPntStr(pntStr, idxOfDot, radix)
    
    If (ret <> (Len(pntStr) + 1)) Then '文字列はn進値ではない
        convNPntToNPnt = CVErr(xlErrNum) '#NUM!を返却
        Exit Function
        
    End If
    
    '整数部&小数部分解
    intPtOfBefore = Left(pntStr, idxOfDot - 1)
    
    '整数部を2進変換
    If (radix = 2) Then
        intPtOfAfter = convIntPrtOfBinToIntPrtOfDecPrt(intPtOfBefore) '2進数が変換対象の時
        
    Else
        intPtOfAfter = convIntPrtOfDecPntToIntPrtOfBinPnt(intPtOfBefore) '10進数が変換対象の時
        
    End If
    
    If (idxOfDot = ret) Then '小数部は存在しない場合
        frcPtOfBefore = ""
        frcPtOfAfter = ""
        
    Else '小数部が存在する場合
        frcPtOfBefore = Right(pntStr, Len(pntStr) - idxOfDot)
        
        '小数部を2進変換
        If (radix = 2) Then
            frcPtOfAfter = convFrcPrtOfBinPntToFrcPrtOfDecPnt(frcPtOfBefore) '2進数が変換対象の時
        Else
            frcPtOfAfter = convFrcPrtOfDecPntToFrcPrtOfBinPnt(frcPtOfBefore, numOfDigits) '10進数が変換対象の時
        End If
        
        If frcPtOfAfter <> "" Then '小数点付加が必要な場合
            frcPtOfAfter = DOT & frcPtOfAfter
            
        End If
        
    End If
    
    '文字列結合
    convNPntToNPnt = sign & intPtOfAfter & frcPtOfAfter

End Function

'
'2数を加算する
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
    
    'val1の文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(value1, intPrtOfVal1, frcPrtOfVal1, isMinusOfVal1, 10)
    
    If (retOfSeparatePnt <> 0) Then 'val1は10進値として不正
        addDecPntDecPnt = addDecPntDecPnt
        Exit Function
        
    End If
    
    'valwの文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(value2, intPrtOfVal2, frcPrtOfVal2, isMinusOfVal2, 10)
    
    If (retOfSeparatePnt <> 0) Then 'val2は10進値として不正
        addDecPntDecPnt = addDecPntDecPnt
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
            tmpVal = addition(tmpVal1, tmpVal2)
            toRetSign = "-"
            
        Else 'value2はプラス値
            tmpVal = substitution(tmpVal1, tmpVal2, substitutionWasMinus)
            If (substitutionWasMinus) Then
                toRetSign = ""
            Else
                toRetSign = "-"
            End If
        
        End If
        
    Else 'value1はプラス値
        If (isMinusOfVal2) Then 'value2はマイナス値
            tmpVal = substitution(tmpVal1, tmpVal2, substitutionWasMinus)
            If (substitutionWasMinus) Then
                toRetSign = "-"
            Else
                toRetSign = ""
            End If
            
        Else 'value2はプラス値
            tmpVal = addition(tmpVal1, tmpVal2)
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
    
    addDecPntDecPnt = toRetSign & intPrt & IIf(frcPrt = "", "", DOT & frcPrt)

End Function

'
'1st引数を2nd引数で掛ける
'2nd引数は1~9のみ可
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
    
    '引数チェック
    If Not ((-9 <= multiplier) And (multiplier <= 9)) Then
        multipleDecPntByOneDig = CVErr(xlErrNum) '#NUM!を返す
        Exit Function
        
    End If
    
    '被乗数の文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(multiplicand, intPrtOfMultiplicand, frcPrtOfMultiplicand, multiplicandIsMinus, 10)
    If (retOfSeparatePnt <> 0) Then '10進値として不正
        multipleDecPntByOneDig = retOfSeparatePnt
        Exit Function
        
    End If
    
    '乗数の符号チェック
    If (multiplier = 0) Then '0掛けの場合
        multipleDecPntByOneDig = "0" '0を返す
        Exit Function
        
    ElseIf (multiplier < 0) Then '乗数がマイナス値
        multiplierIsMinus = True
        multiplier = Abs(multiplier)
        
    Else '乗数はプラス値
        multiplierIsMinus = False
        
    End If
    
    '乗算
    retOfMultiple = multiple(intPrtOfMultiplicand & frcPrtOfMultiplicand, CByte(multiplier))
    
    intPrtOfAns = Left(retOfMultiple, Len(retOfMultiple) - Len(frcPrtOfMultiplicand))
    frcPrtOfAns = Right(retOfMultiple, Len(frcPrtOfMultiplicand))
    
    
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
        signOfAns = "-"
    Else
        signOfAns = ""
    
    End If
    
    multipleDecPntByOneDig = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)
    
End Function

'
'1st引数を2nd引数で割る
'割り切れない時は、3rd引数で指定された回数だけ割った結果を返す
'2nd引数は1~9のみ可
'3rd引数が(-)値の場合は際限なく割り続ける
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
    
    '引数チェック
    If Not ((-9 <= divisor) And (divisor <= 9)) Then
        divideDecPntByOneDig = CVErr(xlErrNum) '#NUM!を返す
        Exit Function
        
    End If
    
    If (divisor = 0) Then
        divideDecPntByOneDig = CVErr(xlErrDiv0) '#DIV0!を返す
        Exit Function
        
    End If
    
    
    '被除数の文字列チェック&小数、整数分解
    retOfSeparatePnt = separateToIntAndFrc(dividend, intPrtOfDividend, frcPrtOfDividend, dividendIsMinus, 10)
    If (retOfSeparatePnt <> 0) Then '10進値として不正
        divideDecPntByOneDig = retOfSeparatePnt
        Exit Function
        
    End If
    
    '除数の符号チェック
    If (divisor = 0) Then '0掛けの場合
        divideDecPntByOneDig = "0" '0を返す
        Exit Function
        
    ElseIf (divisor < 0) Then '除数がマイナス値
        divisorIsMinus = True
        divisor = Abs(divisor)
        
    Else '除数はプラス値
        divisorIsMinus = False
        
    End If
    
    '除算
    retOfDivide = divide(intPrtOfDividend & frcPrtOfDividend, CByte(divisor), limitOfRepTimes)
    
    intPrtOfAns = Left(retOfDivide, Len(intPrtOfDividend))
    frcPrtOfAns = Right(retOfDivide, Len(retOfDivide) - Len(intPrtOfDividend))
    
    
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
        signOfAns = "-"
    Else
        signOfAns = ""
    
    End If
    
    divideDecPntByOneDig = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)
    
End Function

'
'小数部と整数部に分解する
'
'成功の場合は0を返却する
'失敗の場合はCvErrを返却する
'
Private Function separateToIntAndFrc(ByVal pnt As String, ByRef intPrt As String, ByRef frcPrt, ByRef isMinus As Boolean, ByVal radix As Byte) As Variant
    
    Dim idxOfDot As Long
    
    Dim value1IsMinus
    
    Dim retToIsMinus As Boolean
    Dim retToIntPrt As String
    Dim retToFrcPrt As String
    
    '字列長チェック
    If (Len(pnt) < 1) Then '文字列長が0
        separateToIntAndFrc = CVErr(xlErrValue) '#NUM!を返す
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
    
    '10進値として正しいかチェック
    ret = checkNumeralPntStr(pnt, idxOfDot, radix)
    
    If (ret <> (Len(pnt) + 1)) Then 'value1は10進値として不正
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
Private Function checkNumeralPntStr(ByVal decStr As String, ByRef idxOfDot As Long, ByRef radix As Byte) As Long
    
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
Private Function addition(ByVal val1 As String, ByVal val2 As String, Optional ByVal fill0Left As Boolean = True) As String
    
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
        tmpStr = Format(CInt(Mid(val1, idxOfVal, 1)) + CInt(Mid(val2, idxOfVal, 1)) + carrier, "00")
        carrier = CInt(Left(tmpStr, 1))
        stringBuilder(idxOfVal - 1) = Right(tmpStr, 1)
        
    Next idxOfVal
    
    ret = Join(stringBuilder, vbNullString)
    
    If (carrier > 0) Then '桁増えたか
        ret = CInt(carrier) & ret
        
    End If
    
    addition = ret
    
End Function

'
'val1からval2を減算する
'
Private Function substitution(ByVal val1 As String, ByVal val2 As String, ByRef minus As Boolean) As String
    
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
        val1Digit = CInt(Mid(val1, idxOfVal, 1))
        val2Digit = CInt(Mid(val2, idxOfVal, 1))
        
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
        substitution = String(idxMxOfVal, "0")
        minus = False
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
        
        val1Digit = CInt(Mid(val1, idxOfVal, 1))
        val2Digit = CInt(Mid(val2, idxOfVal, 1))
        
        '繰り下がりチェック
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
'1桁数値による乗算をする
'
Private Function multiple(ByVal multiplicand As String, ByVal multiplier As Byte) As String
    
    Dim carrier As Byte
    Dim digitOfMultiplicand As Byte
    Dim digitIdxOfMultiplicand As Long
    Dim stringBuilder() As String '割り算結果格納用
    Dim idxOfStringBuilder As Long
    
    '
    'multiplicandが有効な10進値であるかはチェックしない
    'multiplierは1桁であることはチェックしない
    '
    
    
    If (multiplier = 0) Then '×0の場合は0を返す
        multiple = "0"
        Exit Function
    
    ElseIf (multiplier = 1) Then '×1の場合はそのまま返す
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
        
        ReDim Preserve stringBuilder(idxOfStringBuilder) '領域拡張
        stringBuilder(idxOfStringBuilder) = Right(tmpStr, 1)
        
        digitIdxOfMultiplicand = digitIdxOfMultiplicand - 1
        idxOfStringBuilder = idxOfStringBuilder + 1
    
    Loop While digitIdxOfMultiplicand > 0 '被乗数が残っている間
    
    If (carrier <> "0") Then
        ReDim Preserve stringBuilder(idxOfStringBuilder) '領域拡張
        stringBuilder(idxOfStringBuilder) = carrier
        
    End If
    
    multiple = Join(invertStringArray(stringBuilder), vbNullString) '文字列連結
    
End Function

'
'1桁数値による除算をする
'
'limitOfRepTimes:
'    Indivisible Numberに対するdivide回数制限
'    (-)値を設定した場合は、無限に割り続ける
'
Private Function divide(ByVal dividend As String, ByVal divisor As Byte, ByVal limitOfRepTimes As Long) As String

    '変数宣言
    Dim quot As Byte   '商
    Dim rmnd As Byte '余り
    
    Dim repTimes As Long 'IndivisibleNumberに対するdivide回数
    
    Dim digitOfDividend As Byte '一時被除数
    
    Dim stringBuilder() As String '割り算結果格納用
    Dim digitIdxOfDividend As Long 'Division結果文字列長
    
    '
    'dividendが有効な10進値であるかはチェックしない
    'divisorは1桁であることはチェックしない
    '
    
    '1割チェック
    If divisor = 1 Then
        divide = dividend '1割の場合はそのまま返す
        Exit Function
        
    End If
    
    '初期化
    rmnd = 0
    digitIdxOfDividend = 1
    repTimes = 0
    
    '実行ループ
    Do
        digitOfDividend = CByte(CStr(rmnd) & Mid(dividend, digitIdxOfDividend, 1)) '上位桁の余り & 該当桁
        
        quot = digitOfDividend \ divisor '商
        rmnd = digitOfDividend Mod divisor '余り
        
        ReDim Preserve stringBuilder(digitIdxOfDividend - 1) '領域拡張
        stringBuilder(digitIdxOfDividend - 1) = CStr(quot) '商を追記
        
        digitIdxOfDividend = digitIdxOfDividend + 1
        
        If (rmnd > 0) And (Len(dividend) < digitIdxOfDividend) Then '余りがあるけれど、次の桁が無い
            
            If (limitOfRepTimes > -1) And (repTimes < limitOfRepTimes) Then '再帰計算回数が指定回数以下
                dividend = dividend & "0" '"0"を付加
                repTimes = repTimes + 1
                
            End If
            
        End If
        
    Loop While digitIdxOfDividend <= Len(dividend) '最終文字に到達しない間
    
    divide = Join(stringBuilder, vbNullString) '文字列連結
    
End Function

'
'10進小数部分から2進小数部分に変換する
'
'numOfDigits:
'    求める小数点以下の桁数
'    0を指定した場合は空文字を返す
'
Private Function convFrcPrtOfDecPntToFrcPrtOfBinPnt(ByVal frcPt As String, ByVal numOfDigits As Long) As String
    
    Dim stringBuilder() As String 'bit格納用
    Dim repTimes As Long
    
    '
    '有効10進数文字列かどうかはチェックしない
    '
    
    If (frcPt = "") Or (numOfDigits = 0) Then '空文字指定か、求める桁数=0
        convFrcPrtOfDecPntToFrcPrtOfBinPnt = "" '空文字を返却
        Exit Function
        
    End If
    
    '右側の0を取り除く
    Do While Right(frcPt, 1) = "0"
        frcPt = Left(frcPt, Len(frcPt) - 1)
        
    Loop
    
    If (frcPt = "") Then '全部"0"だったら
        convFrcPrtOfDecPntToFrcPrtOfBinPnt = "0"
        Exit Function
        
    End If
    
    '掛け算する
    repTimes = 0
    sizeOfStringBuilder = 0
    Do
        
        tmp = multiple(frcPt, 2)
        
        ReDim Preserve stringBuilder(repTimes) '領域拡張
        
        If (Len(tmp) > Len(frcPt)) Then '桁上がりが発生した場合
            stringBuilder(repTimes) = "1"
            frcPt = Right(tmp, Len(tmp) - 1)
            
        Else '桁上がりが発生しなかった場合
            stringBuilder(repTimes) = "0"
            frcPt = tmp
        
        End If
        
        If frcPt = "0" Then 'bin変換終了
            Exit Do
            
        ElseIf Right(frcPt, 1) = "0" Then '右桁のみ"0"があったら
            frcPt = Left(frcPt, Len(frcPt) - 1) '"0"は消す
        
        End If
        
        repTimes = repTimes + 1
        
    Loop While IIf(numOfDigits < 0, True, (repTimes < numOfDigits)) '繰り返し回数以下
    
    convFrcPrtOfDecPntToFrcPrtOfBinPnt = Join(stringBuilder, vbNullString) '文字列連結
    
End Function

'
'10進整数部分から2進整数部分に変換する
'
Private Function convIntPrtOfDecPntToIntPrtOfBinPnt(ByVal intPt As String) As String
    
    Dim stringBuilder() As String 'bit格納用
    Dim sizeOfStringBuilder As Long
    
    '
    '有効10進数文字列かどうかはチェックしない
    '
    
    If (intPt = "") Then '空文字の場合
        convIntPrtOfDecPntToIntPrtOfBinPnt = "0"
        Exit Function
        
    End If
    
    '左側の0を取り除く
    Do While Left(intPt, 1) = "0"
        intPt = Right(intPt, Len(intPt) - 1)
        
    Loop
    
    If (intPt = "") Then '全部"0"だったら
        convIntPrtOfDecPntToIntPrtOfBinPnt = "0"
        Exit Function
        
    End If
    
    sizeOfStringBuilder = 0
    
    'bit生成
    Do While (intPt <> "1")
        ret = divide(intPt, 2, 0)
        
        ReDim Preserve stringBuilder(sizeOfStringBuilder) '領域拡張
        
        If (Right(intPt, 1) Like "[0,2,4,6,8]") Then '2で割ったあまりは0
            stringBuilder(sizeOfStringBuilder) = "0"
            
        Else '2で割ったあまりは1
            stringBuilder(sizeOfStringBuilder) = "1"
            
        End If
        
        sizeOfStringBuilder = sizeOfStringBuilder + 1
        
        If (Left(ret, 1) = "0") Then
            intPt = Right(ret, Len(ret) - 1)
            
        Else
            intPt = ret
            
        End If
        
    Loop
    
    '最上位Bit付加
    ReDim Preserve stringBuilder(sizeOfStringBuilder) '領域拡張
    stringBuilder(sizeOfStringBuilder) = "1"
    
    convIntPrtOfDecPntToIntPrtOfBinPnt = Join(invertStringArray(stringBuilder), vbNullString) '文字列連結
    
End Function

'
'2進小数部分から10進小数部分に変換する
'
Private Function convFrcPrtOfBinPntToFrcPrtOfDecPnt(ByVal frcPt As String) As String
    
    Dim lpMx As Long
    Dim decStr As String
    
    Dim minusNpowerOf2 As String
    
    '
    '有効2進文字列かどうかはチェックしない
    '
    
    '引数チェック
    If (frcPt = "") Then '空文字指定の場合
        convFrcPrtOfBinPntToFrcPrtOfDecPnt = "0" '"0"を返す
        Exit Function
        
    End If
    
    lpMx = Len(frcPt)
    decStr = "0"
    minusNpowerOf2 = "5"
    
    '生成ループ
    For cnt = 1 To lpMx
        If (Mid(frcPt, cnt, 1) = "1") Then
            decStr = addition(minusNpowerOf2, decStr, False)
        End If
        
        minusNpowerOf2 = divide(minusNpowerOf2, 2, 1)
        
    Next cnt
    
    convFrcPrtOfBinPntToFrcPrtOfDecPnt = decStr
    
End Function

'
'2進整数部分から10進整数部分に変換する
'
Private Function convIntPrtOfBinToIntPrtOfDecPrt(ByVal intPt As String) As String
    
    Dim nPowerOf2 As String
    Dim decStr As String
    
    '
    '有効2進文字列かどうかはチェックしない
    '
    
    '引数チェック
    If (intPt = "") Then '空文字指定の場合
        convIntPrtOfBinToIntPrtOfDecPrt = "0" '"0"を返す
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



