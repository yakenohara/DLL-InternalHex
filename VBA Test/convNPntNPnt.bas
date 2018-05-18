Attribute VB_Name = "convNPntNPnt"
'<PrivateFunction用テスト関数>---------------------------------------------------------------------------------------------------------------------
'
Public Function TESTadd(ByVal val1 As String, ByVal val2 As String, ByVal radix As Byte) As Variant
    TESTadd = add(val1, val2, radix)
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
'--------------------------------------------------------------------------------------------------------------------</PrivateFunction用テスト関数>

'
'2数を和算する
'
'!CAUTION!
'    val1, val2 が有効なn進値であるかはチェックしない
'    radixは2~16の範囲内である事はチェックしない
'
Private Function add(ByVal val1 As String, ByVal val2 As String, ByVal radix As Byte) As String
    
    '変数宣言
    Dim lenOfVal1 As Integer
    Dim lenOfVal2 As Integer
    Dim idxOfVal As Long
    Dim stringBuilder() As String
    Dim decDigitOfVal1 As Integer
    Dim decDigitOfVal2 As Integer
    Dim decCarrier As Integer
    Dim decDigitOfAns  As Integer
    Dim stsOfSub As Variant
    
    '数値列長取得
    lenOfVal1 = Len(val1)
    lenOfVal2 = Len(val2)
    
    '0埋め確認
    If (lenOfVal1 > lenOfVal2) Then
        val2 = String(lenOfVal1 - lenOfVal2, "0") & val2
        idxOfVal = lenOfVal1
        
    Else
        val1 = String(lenOfVal2 - lenOfVal1, "0") & val1
        idxOfVal = lenOfVal2
        
    End If
    
    'ループ前初期化
    ReDim stringBuilder(idxOfVal) '領域確保
    decCarrier = 0
    
    '解の生成ループ
    Do While (idxOfVal > 0)
        
        '対象桁の和算
        decDigitOfVal1 = convNCharToByte(Mid(val1, idxOfVal, 1))
        decDigitOfVal2 = convNCharToByte(Mid(val2, idxOfVal, 1))
        decDigitOfAns = decDigitOfVal1 + decDigitOfVal2 + decCarrier
        
        '繰り上がり&解格納
        If (decDigitOfAns >= radix) Then '繰り上がりあり
            decCarrier = 1
            decDigitOfAns = decDigitOfAns - radix
            
        Else '繰り上がりなし
            decCarrier = 0
            
        End If
        
        stringBuilder(idxOfVal) = convByteToNChar(decDigitOfAns) '解を格納
        
        idxOfVal = idxOfVal - 1 'decrement
        
    Loop
    
    '最上位桁格納
    stringBuilder(idxOfVal) = IIf(decCarrier > 0, "1", "")
    
    add = Join(stringBuilder, vbNullString)
    
End Function

'
'乗算をする
'
'!CAUTION!
'    multiplicand, multiplier が有効なn進値であるかはチェックしない
'    radixは2~16の範囲内である事はチェックしない
'
Private Function multiple(ByVal multiplicand As String, ByVal multiplier As String, ByVal radix As Byte) As String

    Dim ansOfMultipleByOneDigit As String
    Dim numOfShift As Long
    Dim tmpAns As String
    Dim stsOfSub As Variant
    Dim idxOfMultiplier As Long
    
    'multiplierの不要な0を取り除く
    multiplier = removeLeft0(multiplier)
    
    numOfShift = 0
    tmpAns = String(Len(multiplicand), "0")
    
    '乗算ループ
    For idxOfMultiplier = Len(multiplier) To 1 Step -1
        
        digitOfMultiplier = Mid(multiplier, idxOfMultiplier, 1)
        
        If (digitOfMultiplier <> "0") Then '1以上の数値の時だけ、解に足し合わせる
            ansOfMultipleByOneDigit = multipleByOneDigit(multiplicand, digitOfMultiplier, radix)
            tmpAns = add(tmpAns, ansOfMultipleByOneDigit & String(numOfShift, "0"), radix)
            
        End If
        
        numOfShift = numOfShift + 1
        
    Next idxOfMultiplier
    
    multiple = tmpAns
    
End Function

'
'1桁数値による乗算をする
'
'!CAUTION!
'    multiplicand, multiplierCh が有効なn進値であるかはチェックしない
'    radixは2~16の範囲内である事はチェックしない
'
Private Function multipleByOneDigit(ByVal multiplicand As String, ByVal multiplierCh As String, ByVal radix As Byte) As String

    Dim decMultiplier As Byte
    Dim stsOfSub As Variant
    
    Dim decDigitOfMultiplicand As Byte
    Dim decCarrier As Byte
    Dim decDigitOfAns  As Byte
    
    Dim digitIdxOfMultiplicand As Long
    Dim stringBuilder() As String '割り算結果格納用
    
    '乗数の10進変換
    decMultiplier = convNCharToByte(multiplierCh)
    
    '0掛け&1掛けチェック
    If (decMultiplier = 0) Then
        multipleByOneDigit = String(Len(multiplicand), "0") '0掛けの場合は0を返す
        Exit Function
    
    ElseIf (multiplierCh = "1") Then '1掛けの場合はそのまま返す
        multipleByOneDigit = multiplicand
        Exit Function
        
    End If
    
    'ループ前初期化
    digitIdxOfMultiplicand = Len(multiplicand)
    ReDim stringBuilder(digitIdxOfMultiplicand) '領域確保
    decCarrier = 0
    
    Do While (digitIdxOfMultiplicand > 0) '被乗数が残っている間
        
        '対象桁の乗算
        decDigitOfMultiplicand = convNCharToByte(Mid(multiplicand, digitIdxOfMultiplicand, 1))
        decDigitOfAns = decDigitOfMultiplicand * decMultiplier + decCarrier
        
        digitOfAns = convIntPrtOfNPntToIntPrtOfNPnt(decDigitOfAns, 10, radix) '10進→n進変換
        
        '繰り上がり&解格納
        If (Len(digitOfAns) = 2) Then '繰り上がりあり
            decCarrier = convNCharToByte(Left(digitOfAns, 1))
            digitOfAns = Right(digitOfAns, 1)
            
        Else '繰り上がりなし
            decCarrier = 0
            
        End If
        
        '解格納
        stringBuilder(digitIdxOfMultiplicand) = digitOfAns
        
        digitIdxOfMultiplicand = digitIdxOfMultiplicand - 1 'decrement
        
    Loop
    
    '最上位桁格納
    stringBuilder(digitIdxOfMultiplicand) = IIf(decCarrier > 0, convByteToNChar(decCarrier), "")
    
    multipleByOneDigit = Join(stringBuilder, vbNullString) '文字列連結
    
End Function

'
'除算をする
'
'以下の場合は空文字を返却し、
'errCodeにエラーコードを格納する
'    ・0割の場合。(エラーコードは#DIV/0!)
'    ・dividend / divisor にlong型で取り扱えない大きな数値がある場合。(エラーコードは#NUM!)
'
'numOfFrcDigits:
'    求める小数点以下の桁数
'    指定桁数で除算を打ち切る
'    (-)値を設定した場合は、小数点以下は求めない
'
'remainder
'    剰余
'    (numOfFrcDigits > 0)の場合は、
'    numOfFrcDigitsでの剰余を格納する
'    ex:)
'    【前提】10 / 8 = 1.2 余り 0.4
'    【実行方法】x = divide("10", "8", 10, 1, rm, code)
'    【結果】 x:012
'            rm:4
'
'!CAUTION!
'    dividend, divisor が有効なn進値であるかはチェックしない
'    radixは2~16の範囲内である事はチェックしない
'
Private Function divide(ByVal dividend As String, ByVal divisor As String, ByVal radix As Byte, ByVal numOfFrcDigits As Long, ByRef remainder As String, ByRef errCode As Variant) As String

    '変数宣言
    Dim quot As Long '商
    Dim rmnd As Long '余り
    Dim repTimes As Long 'IndivisibleNumberに対するdivide回数
    Dim digitOfDividend As Long '一時被除数
    Dim stringBuilder() As String '商格納用
    Dim digitIdxOfDividend As Long 'Division結果文字列長
    Dim divisorDec As Long
    Dim stsOfSub As Variant
    
    'divisorの不要な0を取り除く
    divisor = removeLeft0(divisor)
    
    '<divisorの10進変換>------------------------------------------------------------
    
    tmp = convIntPrtOfNPntToIntPrtOfNPnt(divisor, radix, 10)
    
    'divisorのLong型変換
    On Error GoTo OVERFLOW
    divisorDec = CLng(tmp)
    
    If divisorDec = 0 Then '0割チェック
        divide = ""
        errCode = CVErr(xlErrDiv0) '#DIV0!を返す
        Exit Function
    
    ElseIf divisorDec = 1 Then
        divide = dividend '1割の場合はそのまま返す
        Exit Function
        
    End If
    
    '-----------------------------------------------------------</divisorの10進変換>
    
    '初期化
    rmnd = 0
    digitIdxOfDividend = 1
    repTimes = 0
    
    '実行ループ
    Do
        digitOfDividend = rmnd * radix + convNCharToByte(Mid(dividend, digitIdxOfDividend, 1)) '上位桁の余り & 該当桁
        
        quot = digitOfDividend \ divisorDec '商
        rmnd = digitOfDividend Mod divisorDec '余り
        
        ReDim Preserve stringBuilder(digitIdxOfDividend - 1) '領域拡張
        
        '商を追記
        'ある桁に対する除算の商は、基数未満しか取り柄ない
        stringBuilder(digitIdxOfDividend - 1) = convByteToNChar(quot)
        
        digitIdxOfDividend = digitIdxOfDividend + 1
        
        If (rmnd > 0) And (Len(dividend) < digitIdxOfDividend) Then '余りがあるけれど、次の桁が無い
        
            If (numOfFrcDigits > -1) And (repTimes < numOfFrcDigits) Then '再帰計算回数が指定回数以下
                dividend = dividend & "0" '"0"を付加
                
                repTimes = repTimes + 1
                
            End If
            
        End If
        
    Loop While digitIdxOfDividend <= Len(dividend) '最終文字に到達しない間
    
    If (rmnd = 0) Then '余りが0のとき
        remainder = "0"
        
    Else '余りが存在する時
        
        remainder = convIntPrtOfNPntToIntPrtOfNPnt(rmnd, 10, radix)
    
    End If
    
    divide = Join(stringBuilder, vbNullString) '文字列連結
    
    Exit Function
    
OVERFLOW: 'オーバーフローの場合
    divide = ""
    errCode = CVErr(xlErrNum) '#NUM!を返す
    Exit Function
    
End Function

'
'n進整数部をn進整数部に変換する
'
'!CAUTION!
'    intStrが有効な(fromRadix)進値であるかはチェックしない
'    fromRadix,toRadixは2~16の範囲内である事はチェックしない
'
Private Function convIntPrtOfNPntToIntPrtOfNPnt(ByVal intStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte) As String
    
    Dim stsOfSub As Variant
    Dim retOfTryConvRadix As String
    Dim strLenOfToRadix As Long
    Dim stringBuilder() As String '変換後文字列生成用
    Dim sizeOfStringBuilder As Long
    Dim rm As String
    
    intStr = removeLeft0(intStr)
    
    '変換する基数が同じ場合
    If (fromRadix = toRadix) Then
        convIntPrtOfNPntToIntPrtOfNPnt = intStr '"0"を取り除いただけの値を返す
        Exit Function
        
    End If
    
    'convRadixで解決可能かどうかチェック
    retOfTryConvRadix = tryConvRadix(intStr, fromRadix, toRadix, stsOfSub)
    
    If (retOfTryConvRadix <> "") Then '基数変換用テーブルに解があった
        convIntPrtOfNPntToIntPrtOfNPnt = retOfTryConvRadix
        Exit Function
        
    End If
    
    '生成ループ前初期化
    sizeOfStringBuilder = 0
    chOfToRadix = convRadix(fromRadix, toRadix)
    strLenOfToRadix = Len(chOfToRadix)
    
    '生成ループ - toRadixによる除算によって解を求める -
    Do While True
        
        If (Len(intStr) <= strLenOfToRadix) Then
            
            retOfTryConvRadix = tryConvRadix(intStr, fromRadix, 10, stsOfSub)
            
            If (retOfTryConvRadix <> "") Then
                
                If (CByte(retOfTryConvRadix) < toRadix) Then '基数で割れる数がなくなった ※必ず retOfTryConvRadix > 0 とはなる ※
                    Exit Do
                    
                End If
                
            End If
            
        End If
        
        intStr = divide(intStr, chOfToRadix, fromRadix, 0, rm, stsOfSub) '16(10進値)以下による除算なので、オーバーフローは発生し得ない
        
        intStr = removeLeft0(intStr) '左側の不要な"0"を取り除く
        
        ReDim Preserve stringBuilder(sizeOfStringBuilder) '領域拡張
        
        '剰余を(toRadix)進値に変換した結果が算出Digit
        stringBuilder(sizeOfStringBuilder) = convIntPrtOfNPntToIntPrtOfNPnt(rm, fromRadix, toRadix)
        
        sizeOfStringBuilder = sizeOfStringBuilder + 1
        
    Loop
    
    ReDim Preserve stringBuilder(sizeOfStringBuilder) '領域拡張
    
    '最上位Bit付加
    '剰余を(toRadix)進値に変換した結果が算出Digit
    stringBuilder(sizeOfStringBuilder) = convIntPrtOfNPntToIntPrtOfNPnt(intStr, fromRadix, toRadix)
    
    convIntPrtOfNPntToIntPrtOfNPnt = Join(invertStringArray(stringBuilder), vbNullString) '文字列連結
    
End Function

'
'n進小数部をn進小数部に変換する
'
'numOfDigits:
'    求める桁数
'    0以下を指定した場合は、空文字を返却する
'
'!CAUTION!
'    frcStrが有効な(fromRadix)進値であるかはチェックしない
'    fromRadix,toRadixは2~16の範囲内である事はチェックしない
'
Private Function convFrcPrtOfNPntToFrcPrtOfNPnt(ByVal frcStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByVal numOfDigits As Long) As String
    
    Dim stsOfSub As Variant
    Dim stringBuilder() As String '変換後文字列生成用
    Dim sizeOfStringBuilder As Long
    Dim retOfMultiple As String
    
    frcStr = removeRight0(frcStr)
    
    '変換する基数が同じ場合
    If (fromRadix = toRadix) Then
        convFrcPrtOfNPntToFrcPrtOfNPnt = frcStr '"0"を取り除いただけの値を返す
        Exit Function
        
    End If
    
    '"0"を変換する場合
    If (frcStr = "0") Then
        convFrcPrtOfNPntToFrcPrtOfNPnt = "0" '"0"を返す
        Exit Function
        
    End If
    
    '生成ループ前初期化
    strOfToRadix = convRadix(fromRadix, toRadix)
    sizeOfStringBuilder = 0
    lenOfFrcStrB = Len(frcStr)
    
    '生成ループ - toRadixによる乗算によって解を求める -
    Do While (sizeOfStringBuilder < numOfDigits)
        
        '小数の積が0になったら終了
        If (frcStr = "0") Then
            Exit Do
            
        End If
        
        frcStr = multiple(frcStr, strOfToRadix, fromRadix)
        
        ReDim Preserve stringBuilder(sizeOfStringBuilder) '領域拡張
        
        '増えた桁を拾う
        lenOfFrcStrA = Len(frcStr)
        
        If (lenOfFrcStrA > lenOfFrcStrB) Then
            tmp = Left(frcStr, lenOfFrcStrA - lenOfFrcStrB)
            frcStr = Right(frcStr, lenOfFrcStrB)
            increasedDigits = convIntPrtOfNPntToIntPrtOfNPnt(tmp, fromRadix, toRadix)
            
        Else
            increasedDigits = "0"
        
        End If
        
        stringBuilder(sizeOfStringBuilder) = increasedDigits '解を追記
        
        frcStr = removeRight0(frcStr) ' 右側の不要な0を取り除く
        
        lenOfFrcStrB = Len(frcStr)
        sizeOfStringBuilder = sizeOfStringBuilder + 1
        
    Loop
    
    convFrcPrtOfNPntToFrcPrtOfNPnt = Join(stringBuilder, vbNullString)

End Function

'
'数値文字が10進値でいくつかを返す
'
'!CAUTION!
'    chは2~Fの範囲内である事はチェックしない
'
Private Function convNCharToByte(ByVal ch As String) As Byte
    
    Dim toRetByte As Byte
    Dim ascOfA As Integer
    
    ascOfA = Asc("A")
    ascOfCh = Asc(ch)
    
    If (ascOfA <= ascOfCh) Then 'A~Fの場合
        toRetByte = 10 + (ascOfCh - ascOfA)
    
    Else '0~9の場合
        toRetByte = CByte(ch)
    
    End If
    
    convNCharToByte = toRetByte
    
End Function

'
'10進値から数値文字を返す
'
'!CAUTION!
'    bytは0~16の範囲内である事はチェックしない
'
Private Function convByteToNChar(ByVal byt As Byte) As String
    
    Dim toRetStr As String
    
    If (byt > 9) Then 'A~Fの場合
        toRetStr = Chr((byt - 10) + Asc("A"))
    
    Else '0~9の場合
        toRetStr = Chr(byt + Asc("0"))
    
    End If
    
    convByteToNChar = toRetStr
    
End Function

'
'基数変換で必要な文字列を得る
'
'!CAUTION!
'    fromRadix,toRadixは2~16の範囲内である事はチェックしない
'
Private Function convRadix(ByVal fromRadix As Byte, ByVal toRadix As Byte) As String
    
    Dim radixTable As Variant
    
    '基数変換用テーブル
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
'convRadixを使ってN進→N進変換をトライする
'変換成功の場合は、変換後のN進値を返す
'失敗の場合は、endStatusに#N/A!を格納し、空文字を返す
'
'!CAUTION!
'    fromRadix,toRadixは2~16の範囲内である事はチェックしない
'
Private Function tryConvRadix(ByVal intStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByRef endStatus As Variant) As String
    
    Dim idxOfRTable As Byte
    Dim toRetStr As String
    
    'convRadixで解決可能かどうかチェック
    For idxOfRTable = 0 To 16
        If (intStr = convRadix(fromRadix, idxOfRTable)) Then '基数変換用テーブルに解があった
            toRetStr = convRadix(toRadix, idxOfRTable) '基数変換テーブルから解を返す
            Exit For
            
        End If
        
    Next idxOfRTable
    
    If (idxOfRTable > 16) Then '見つからなかった場合
        endStatus = CVErr(xlErrNA)
        toRetStr = ""
    
    End If
    
    tryConvRadix = toRetStr
    
End Function

'
'左側の不要な"0"を取り除く
'
'以下の場合は、"0"を返す
'    ・空文字を指定した場合
'    ・すべて"0"(正規表現で表す"0+")な文字列の場合
'
'!CAUTION!
'    intStrが有効な数値文字列かどうかはチェックしない
'
Private Function removeLeft0(ByVal intStr As String) As String
    
    Dim lpIdx As Long
    Dim lpMx As Long
    Dim toRetStr As String
    
    lpMx = Len(intStr)
    lpIdx = 1
    
    '文字列捜査ループ
    Do While (lpIdx <= lpMx)
        
        If (Mid(intStr, lpIdx, 1) <> "0") Then '捜査対象文字は"0"でない
            Exit Do
            
        End If
        
        lpIdx = lpIdx + 1 'increment
        
    Loop
    
    If (lpIdx > lpMx) Then '空文字 or すべて"0"な文字列
        toRetStr = "0"
        
    Else
        toRetStr = Right(intStr, lpMx - lpIdx + 1)
        
    End If
    
    removeLeft0 = toRetStr
    
End Function

'
'右側の不要な"0"を取り除く
'
'以下の場合は、"0"を返す
'    ・空文字を指定した場合
'    ・すべて"0"(正規表現で表す"0+")な文字列の場合
'
'!CAUTION!
'    intStrが有効な数値文字列かどうかはチェックしない
'
Private Function removeRight0(ByVal intStr As String) As String
    
    Dim lpIdx As Long
    Dim toRetStr As String
    
    lpIdx = Len(intStr)
    
    '文字列捜査ループ
    Do While (lpIdx > 0)
        
        If (Mid(intStr, lpIdx, 1) <> "0") Then '捜査対象文字は"0"でない
            Exit Do
            
        End If
        
        lpIdx = lpIdx - 1 'decrement
        
    Loop
    
    If (lpIdx = 0) Then  '空文字 or すべて"0"な文字列
        toRetStr = "0"
        
    Else
        toRetStr = Left(intStr, lpIdx)
        
    End If
    
    removeRight0 = toRetStr
    
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



