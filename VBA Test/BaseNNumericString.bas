Attribute VB_Name = "BaseNNumericString"
'<定数>------------------------------------------------------------------------------------------

Private Const DOT As String = "." '小数点表記

'割り切れない数値に対して何回割り算するか
Const DEFAULT_LIMIT_OF_FRC_DIGITS As Long = 30
'
'-----------------------------------------------------------------------------------------</定数>

'
'2数を加算する
'
'引数が不正の場合は、以下に応じたCvErrを返却する
'    ・radixが2~16以外か、数値列はn進値として不正の場合(エラーコードは#NUM!)
'    ・数値列が空文字かNullの場合(エラーコードは#NULL!)
'
Public Function addBaseNNumber(ByVal val1 As String, ByVal val2 As String, Optional ByVal radix As Byte = 10) As Variant
    
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
    
    Dim tmpVal1 As String
    Dim tmpVal2 As String
    Dim tmpAns As String
    
    Dim signOfAns As String
    Dim intPrtOfAns As String
    Dim frcPrtOfAns As String
    
    'val1の文字列チェック&小数、整数分解
    stsOfSub = separateToIntAndFrc(val1, radix, True, intPrtOfVal1, frcPrtOfVal1, isMinusOfVal1)
    If IsError(stsOfSub) Then 'val1はn進値として不正
        addBaseNNumber = stsOfSub 'checkBaseNNumberのエラーコードを返す
        Exit Function
        
    End If
    
    'val2の文字列チェック&小数、整数分解
    stsOfSub = separateToIntAndFrc(val2, radix, True, intPrtOfVal2, frcPrtOfVal2, isMinusOfVal2)
    If IsError(stsOfSub) Then 'val2はn進値として不正
        addBaseNNumber = stsOfSub 'checkBaseNNumberのエラーコードを返す
        Exit Function
        
    End If
    
    '小数部の桁数合わせ
    lenOfVal1FrcPrt = Len(frcPrtOfVal1)
    lenOfVal2FrcPrt = Len(frcPrtOfVal2)
    If (lenOfVal1FrcPrt > lenOfVal2FrcPrt) Then 'val1の桁数が大きい
        frcPrtOfVal2 = frcPrtOfVal2 & String(lenOfVal1FrcPrt - lenOfVal2FrcPrt, "0") 'val2の右側を0埋め
        lenOfVal2FrcPrt = Len(frcPrtOfVal2)
        
    Else 'val2の桁数が大きい
        frcPrtOfVal1 = frcPrtOfVal1 & String(lenOfVal2FrcPrt - lenOfVal1FrcPrt, "0") 'val1の右側を0埋め
        lenOfVal1FrcPrt = Len(frcPrtOfVal1)
        
    End If
    
    tmpVal1 = intPrtOfVal1 & frcPrtOfVal1
    tmpVal2 = intPrtOfVal2 & frcPrtOfVal2
    
    '加算or減算
    If (isMinusOfVal1) Then 'val1はマイナス値
        If (isMinusOfVal2) Then 'val2はマイナス値
            tmpAns = add(tmpVal1, tmpVal2, radix)
            signOfAns = "-"
            
        Else 'val2はプラス値
            tmpAns = subtract(tmpVal1, tmpVal2, radix, subtractionWasMinus)
            If (subtractionWasMinus) Then
                signOfAns = ""
            Else
                signOfAns = "-"
            End If
        
        End If
        
    Else 'val1はプラス値
        If (isMinusOfVal2) Then 'val2はマイナス値
            tmpAns = subtract(tmpVal1, tmpVal2, radix, subtractionWasMinus)
            If (subtractionWasMinus) Then
                signOfAns = "-"
            Else
                signOfAns = ""
            End If
            
        Else 'val2はプラス値
            tmpAns = add(tmpVal1, tmpVal2, radix)
            signOfAns = ""
        
        End If
    
    End If
    
    '小数点復活
    intPrtOfAns = Left(tmpAns, Len(tmpAns) - lenOfVal1FrcPrt)
    frcPrtOfAns = Right(tmpAns, lenOfVal1FrcPrt)
    
    '不要な"0"を削除
    intPrtOfAns = removeLeft0(intPrtOfAns)
    If (frcPrtOfAns <> "") Then
        frcPrtOfAns = removeRight0(frcPrtOfAns)
        If (frcPrtOfAns = "0") Then
            frcPrtOfAns = ""
        End If
    End If
    
    '-0確認
    If ((intPrtOfAns & frcPrtOfAns) = "0") Then
        signOfAns = ""
    End If
    
    addBaseNNumber = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)

End Function

'
'2数を乗算する
'
'引数が不正の場合は、以下に応じたCvErrを返却する
'    ・radixが2~16以外か、数値列はn進値として不正の場合(エラーコードは#NUM!)
'    ・数値列が空文字かNullの場合(エラーコードは#NULL!)
'
Public Function multipleBaseNNumber(ByVal multiplicand As String, ByVal multiplier As String, Optional radix As Byte = 10) As Variant

    Dim intPrtOfVal1 As String
    Dim frcPrtOfVal1 As String
    Dim isMinusOfVal1 As Boolean
    
    Dim intPrtOfVal2 As String
    Dim frcPrtOfVal2 As String
    Dim isMinusOfVal2 As Boolean
    
    Dim stsOfSub As Variant
    
    Dim toCutLenOfFrcPrtOfTmpAns As Long
    Dim lenOfTmpAns As Long
    Dim tmpAns As String
    
    Dim signOfAns As String
    Dim intPrtOfAns As String
    Dim frcPrtOfAns As String
    
    'val1の文字列チェック&小数、整数分解
    stsOfSub = separateToIntAndFrc(multiplicand, radix, True, intPrtOfVal1, frcPrtOfVal1, isMinusOfVal1)
    If IsError(stsOfSub) Then 'val1はn進値として不正
        multipleBaseNNumber = stsOfSub 'checkBaseNNumberのエラーコードを返す
        Exit Function
        
    End If
    
    'val2の文字列チェック&小数、整数分解
    stsOfSub = separateToIntAndFrc(multiplier, radix, True, intPrtOfVal2, frcPrtOfVal2, isMinusOfVal2)
    If IsError(stsOfSub) Then 'val2はn進値として不正
        multipleBaseNNumber = stsOfSub 'checkBaseNNumberのエラーコードを返す
        Exit Function
        
    End If
    
    '乗算
    tmpAns = multiple(intPrtOfVal1 & frcPrtOfVal1, intPrtOfVal2 & frcPrtOfVal2, radix)
    
    toCutLenOfFrcPrtOfTmpAns = Len(frcPrtOfVal1) + Len(frcPrtOfVal2)
    lenOfTmpAns = Len(tmpAns)
    
    '整数&小数部の切り出し
    If (lenOfTmpAns > toCutLenOfFrcPrtOfTmpAns) Then '少数部分の桁数はtmpAns内
        intPrtOfAns = Left(tmpAns, (lenOfTmpAns - toCutLenOfFrcPrtOfTmpAns))
        frcPrtOfAns = Right(tmpAns, toCutLenOfFrcPrtOfTmpAns)
        
    Else '少数部分の桁数はtmpAns内に収まらない
        intPrtOfAns = "0"
        frcPrtOfAns = String((toCutLenOfFrcPrtOfTmpAns - lenOfTmpAns), "0") & tmpAns
        
    End If
    
    '不要な"0"を削除
    intPrtOfAns = removeLeft0(intPrtOfAns)
    If (frcPrtOfAns <> "") Then
        frcPrtOfAns = removeRight0(frcPrtOfAns)
        If (frcPrtOfAns = "0") Then
            frcPrtOfAns = ""
        End If
    End If
    
    '符号判定
    If (isMinusOfVal1 Xor isMinusOfVal2) Then
        
        If (intPrtOfAns = "0") And (frcPrtOfAns = "") Then
            signOfAns = ""
            
        Else
            signOfAns = "-"
            
        End If
    Else
        signOfAns = ""
    
    End If
    
    '-0確認
    If ((intPrtOfAns & frcPrtOfAns) = "0") Then
        signOfAns = ""
    End If
    
    multipleBaseNNumber = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)

End Function

'
'1st引数を2nd引数で除算する
'
'引数が不正の場合は、以下に応じたCvErrを返却する
'    ・radixが2~16以外か、数値列はn進値として不正の場合(エラーコードは#NUM!)
'    ・数値列が空文字かNullの場合(エラーコードは#NULL!)
'    ・limitOfFrcDigitsが(-)値 (エラーコードは#NUM!)
'
'以下の場合は、エラーコードを返却する
'    ・0割の場合(エラーコードは#DIV/0!)
'    ・dividend / divisor の数値列内に、Long型で取り扱えない大きな数値がある場合。(エラーコードは#NUM!)
'
'limitOfFrcDigits(Optional)
'    求める小数点以下桁数
'
Public Function divideBaseNNumber(ByVal dividend As String, ByVal divisor As String, Optional ByVal radix As Byte = 10, Optional ByVal limitOfFrcDigits As Long = DEFAULT_LIMIT_OF_FRC_DIGITS) As Variant

    Dim intPrtOfVal1 As String
    Dim frcPrtOfVal1 As String
    Dim isMinusOfVal1 As Boolean
    
    Dim intPrtOfVal2 As String
    Dim frcPrtOfVal2 As String
    Dim isMinusOfVal2 As Boolean
    
    Dim stsOfSub As Variant
    Dim rm As String
    
    Dim toCutLenOfIntPrtOfTmpAns As Long
    Dim lenOfTmpAns As Long
    Dim tmpAns As String
    
    Dim signOfAns As String
    Dim intPrtOfAns As String
    Dim frcPrtOfAns As String
    
    'val1の文字列チェック&小数、整数分解
    stsOfSub = separateToIntAndFrc(dividend, radix, True, intPrtOfVal1, frcPrtOfVal1, isMinusOfVal1)
    If IsError(stsOfSub) Then 'val1はn進値として不正
        divideBaseNNumber = stsOfSub 'checkBaseNNumberのエラーコードを返す
        Exit Function
        
    End If
    
    'val2の文字列チェック&小数、整数分解
    stsOfSub = separateToIntAndFrc(divisor, radix, True, intPrtOfVal2, frcPrtOfVal2, isMinusOfVal2)
    If IsError(stsOfSub) Then 'val2はn進値として不正
        divideBaseNNumber = stsOfSub 'checkBaseNNumberのエラーコードを返す
        Exit Function
        
    End If
    
    If (limitOfFrcDigits < 0) Then '小数点以下桁数指定が0未満
        divideBaseNNumber = CVErr(xlErrNum) '#NUM!を返す
        Exit Function
        
    End If
    
    '除算 - 乗数に小数点以下がある場合は、その桁数を小数点以下算出回数に加えないといけない -
    tmpAns = divide(intPrtOfVal1 & frcPrtOfVal1, intPrtOfVal2 & frcPrtOfVal2, radix, limitOfFrcDigits + Len(frcPrtOfVal2), rm, stsOfSub)
    
    If IsError(stsOfSub) Then '除算処理でエラー
        divideBaseNNumber = stsOfSub 'divideのエラーコードを返す
        Exit Function
        
    End If
    
    toCutLenOfIntPrtOfTmpAns = Len(intPrtOfVal1) + Len(frcPrtOfVal2)
    lenOfTmpAns = Len(tmpAns)
    
    '整数&小数部の切り出し
    If (toCutLenOfIntPrtOfTmpAns <= lenOfTmpAns) Then '整数部分の桁数はtmpAns内
        intPrtOfAns = Left(tmpAns, toCutLenOfIntPrtOfTmpAns)
        frcPrtOfAns = Right(tmpAns, Len(tmpAns) - toCutLenOfIntPrtOfTmpAns)
        
    Else '整数部分の桁数はtmpAns内に収まらない
        intPrtOfAns = tmpAns & String(toCutLenOfIntPrtOfTmpAns - lenOfTmpAns, "0")
        frcPrtOfAns = ""
        
    End If
    
    '不要な"0"を削除
    intPrtOfAns = removeLeft0(intPrtOfAns)
    If (frcPrtOfAns <> "") And (Len(frcPrtOfAns) <> limitOfFrcDigits) Then '小数点以下桁数は指定桁数での算出終了ではない
        frcPrtOfAns = removeRight0(frcPrtOfAns)
        If (frcPrtOfAns = "0") Then
            frcPrtOfAns = ""
        End If
    End If
    
    '符号判定
    If (isMinusOfVal1 Xor isMinusOfVal2) Then
        
        If (intPrtOfAns = "0") And (frcPrtOfAns = "") Then
            signOfAns = ""
            
        Else
            signOfAns = "-"
            
        End If
    Else
        signOfAns = ""
    
    End If
    
    '-0確認
    If ((intPrtOfAns & frcPrtOfAns) = "0") Then
        signOfAns = ""
    End If
    
    divideBaseNNumber = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)

End Function

'
'n進数をn進数に変換する
'引数が不正の場合は、以下に応じたCvErrを返却する
'    ・fromRadix or toRadix が2~16以外か、数値列はfromRadix進値として不正の場合(エラーコードは#NUM!)
'    ・変換元数値列が空文字かNullの場合(エラーコードは#NULL!)
'    ・limitOfFrcDigitsが(-)値 (エラーコードは#NUM!)
'
'fromRadix:
'    変換元数値列の基数
'
'toRadix:
'    変換先数値列の基数
'
'limitOfFrcDigits(Optional):
'    小数点以下の求める桁数
'
Public Function convRadix(ByVal baseNNumericStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, Optional ByVal limitOfFrcDigits As Long = DEFAULT_LIMIT_OF_FRC_DIGITS) As Variant
    
    Dim intPrtOfFrom As String
    Dim frcPrtOfFrom As String
    Dim isMinusOfFrom As Boolean
    
    Dim stsOfSub As Variant
    
    Dim signOfAns As String
    Dim intPrtOfAns As String
    Dim frcPrtOfAns As String
    
    '文字列チェック&小数、整数分解
    stsOfSub = separateToIntAndFrc(baseNNumericStr, fromRadix, True, intPrtOfFrom, frcPrtOfFrom, isMinusOfFrom)
    If IsError(stsOfSub) Then 'n進値として不正
        convRadix = stsOfSub 'checkBaseNNumberのエラーコードを返す
        Exit Function
        
    End If
    
    'toRadixの範囲チェック
    If ((toRadix < 2) Or (16 < toRadix)) Then '変換先基数は2~16の範囲外
        convRadix = CVErr(xlErrNum) '#NUM!を返す
        Exit Function
        
    End If
    
    '整数部のn進→n進変換
    intPrtOfAns = convRadixOfInt(intPrtOfFrom, fromRadix, toRadix)
    
    '小数部のn進→n進変換
    If (frcPrtOfFrom = "") Then '小数部が存在しない場合
        frcPrtOfAns = ""
        
    Else '小数部が存在する場合
        frcPrtOfAns = convRadixOfFrc(frcPrtOfFrom, fromRadix, toRadix, limitOfFrcDigits)
        
    End If
    
    '符号判定
    If isMinusOfFrom Then
        signOfAns = "-"
        
    Else
        signOfAns = ""
    
    End If
    
    convRadix = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)
    
End Function

'
'減基数の補数を得る
'
'引数が不正の場合は、以下に応じたCvErrを返却する
'    ・radixが2~16以外か、数値列はn進値として不正の場合(エラーコードは#NUM!)
'    ・数値列が空文字かNullの場合(エラーコードは#NULL!)
'
Public Function getDiminishedRadixComplement(ByVal baseNNumericStr As String, ByVal radix As Byte) As Variant
    
    Dim intPrtOfVal1 As String
    Dim frcPrtOfVal1 As String
    Dim isMinusOfVal1 As Boolean
    
    Dim stsOfSub As Variant
    Dim tmpVal1 As String
    Dim lenOfTmpVal1 As Long
    Dim stringBuilder() As String
    Dim tmpAns As String
    Dim lpCnt As Long
    
    Dim signOfAns As String
    Dim intPrtOfAns As String
    Dim frcPrtOfAns As String
    
    '文字列チェック&小数、整数分解
    stsOfSub = separateToIntAndFrc(baseNNumericStr, radix, False, intPrtOfVal1, frcPrtOfVal1, isMinusOfVal1)
    If IsError(stsOfSub) Then 'val1はn進値として不正
        getDiminishedRadixComplement = stsOfSub 'checkBaseNNumberのエラーコードを返す
        Exit Function
        
    End If
    
    tmpVal1 = intPrtOfVal1 & frcPrtOfVal1
    lenOfTmpVal1 = Len(tmpVal1)
    ReDim stringBuilder(lenOfTmpVal1)
    
    '補数を求めるループ
    For lpCnt = lenOfTmpVal1 To 1 Step -1
        stringBuilder(lpCnt) = convByteToNChar((radix - 1) - convNCharToByte(Mid(tmpVal1, lpCnt, 1)))
        
    Next lpCnt
    
    tmpAns = Join(stringBuilder, vbNullString)
    
    '整数&小数部の切り出し
    intPrtOfAns = Left(tmpAns, Len(intPrtOfVal1))
    frcPrtOfAns = Right(tmpAns, Len(frcPrtOfVal1))
    
    '符号判定
    If (isMinusOfVal1) Then '(-)値の場合
        signOfAns = "-"
        
    Else
        signOfAns = ""
        
    End If
    
    getDiminishedRadixComplement = signOfAns & intPrtOfAns & IIf(frcPrtOfAns = "", "", DOT & frcPrtOfAns)
    

End Function

'
'数値列がn進数値列かどうかチェックして、
'整数部と小数部に分解する
'小数部の記載がない場合は、小数部は空文字を格納する
'
'成功の場合は0を返却する
'
'失敗の場合は以下に応じたCvErrを返却する
'    　・radixが2~16以外か、数値列はn進値として不正の場合(エラーコードは#NUM!)
'    　・数値列が空文字かNullの場合(エラーコードは#NULL!)
'
'radix
'    基数(2~16のみ)
'
'remove0
'    不要な0(整数部は左側の0、小数部は右側の0)を取り除くかどうか
'    TRUEを指定して小数部が全て0の場合、小数部は空文字を格納する
'
Private Function separateToIntAndFrc(ByVal baseNNumericStr As String, ByVal radix As Byte, ByVal remove0 As Boolean, ByRef intPrt As String, ByRef frcPrt, ByRef isMinus As Boolean) As Variant
    
    Dim retOfCheckBaseNNumber As Long
    Dim idxOfDot As Long
    Dim stsOfSub As Variant
    Dim lenOfBaseNNumericStr As Long
    Dim toRetIsMinus As Boolean
    Dim toRetIntPrt As String
    Dim toRetFrcPrt As String
    
    lenOfBaseNNumericStr = Len(baseNNumericStr)
    
    'n進値として正しいかチェック&符号判定&小数点位置取得
    retOfCheckBaseNNumber = checkBaseNNumber(baseNNumericStr, radix, toRetIsMinus, idxOfDot, stsOfSub)
    If IsError(stsOfSub) Then 'n進値として不正
        separateToIntAndFrc = stsOfSub 'checkBaseNNumberのエラーコードを返す
        Exit Function
        
    End If
    
    '整数部抽出開始位置の判定
    If (toRetIsMinus) Then '(-)値の場合
        stIdxOfIntPrt = 2
    Else
        stIdxOfIntPrt = 1
    End If
    
    '抽出
    If (idxOfDot = 0) Then '小数部の記載がない場合
        toRetIntPrt = Mid(baseNNumericStr, stIdxOfIntPrt, (lenOfBaseNNumericStr - stIdxOfIntPrt) + 1)
        toRetFrcPrt = ""
        
    Else '小数部あり
        toRetIntPrt = Mid(baseNNumericStr, stIdxOfIntPrt, idxOfDot - stIdxOfIntPrt)
        toRetFrcPrt = Right(baseNNumericStr, lenOfBaseNNumericStr - idxOfDot)
        
    End If
    
    '0削除
    If (remove0) Then
        toRetIntPrt = removeLeft0(toRetIntPrt)
        
        If (toRetFrcPrt <> "") Then
            
            toRetFrcPrt = removeRight0(toRetFrcPrt)
                
            If (toRetFrcPrt = "0") Then 'すべて0だったら
                toRetFrcPrt = ""
            End If
        End If
        
    End If
    
    '返却
    intPrt = toRetIntPrt
    frcPrt = toRetFrcPrt
    isMinus = toRetIsMinus
    separateToIntAndFrc = 0
    
End Function


'
'数値列がn進数値列かどうかチェックする
'
'返却値
'    n進値文字列だったの場合はerrCodeに0を格納し、文字長 + 1を返却する
'    そうでない場合は、errCodeに#NUM!を格納し、
'    最初に見つかった10進文字以外の文字位置を返却する
'
'    以下の場合は、errCodeにエラーコードを格納し、0を返却する
'    　・radixが2~16以外の場合(エラーコードは#NUM!)
'    　・引数が空文字かNullの場合(エラーコードは#NULL!)
'
'radix
'    基数(2~16のみ)
'
'idxOfDot(ByRef)
'    小数点文字位置
'    小数点が無かった場合0
'
Private Function checkBaseNNumber(ByVal baseNNumericStr As String, ByVal radix As Byte, ByRef isMinus As Boolean, ByRef idxOfDot As Long, ByRef errCode As Variant) As Long
    
    Dim minOkChar1 As Integer
    Dim maxOkChar1 As Integer
    Dim minOkChar2 As Integer
    Dim maxOkChar2 As Integer
    Dim radixIsBiggerThan10 As Boolean
    Dim cnt As Long
    Dim lpMx As Long
    Dim stCnt As Long
    Dim foundIdxOfDot As Long '小数点文字が最初に見つかった文字位置
    Dim ngIdx As Long
    Dim numOfDigits As Long
    
    '引数チェック
    If (radix < 2) Or (16 < radix) Then
        errCode = CVErr(xlErrNum) '#NUM!を格納する
        checkBaseNNumber = 0
        Exit Function
        
    End If
    
    lpMx = Len(baseNNumericStr)
    
    If (lpMx = 0) Then
        errCode = CVErr(xlErrNull) '#NULL!を格納する
        checkBaseNNumber = 0
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
    
    '符号存在チェック
    If (Left(baseNNumericStr, 1) = "-") Then '符号は(-)
        isMinus = True
        stCnt = 2
        
    Else '符号は(+)
        isMinus = False
        stCnt = 1
        
    End If
    
    '文字列検査ループ
    foundIdxOfDot = 0
    ngIdx = 0
    numOfDigits = 0
    For cnt = stCnt To lpMx
        
        ch = Mid(baseNNumericStr, cnt, 1)
        chCode = Asc(ch)
        
        If (chCode < minOkChar1) Or (maxOkChar1 < chCode) Then  '文字は0~9いずれでもない
            If IIf(radixIsBiggerThan10, (chCode < minOkChar2) Or (maxOkChar2 < chCode), True) Then '文字はA~Fいずれでもない
                
                If (ch = DOT) Then '小数点文字の場合
                    If (foundIdxOfDot = 0) Then '小数点文字の出現は1回目
                    
                        If (numOfDigits = 0) Then '整数部の桁数が0
                            ngIdx = cnt
                            Exit For
                            
                        End If
                        
                        foundIdxOfDot = cnt
                        numOfDigits = 0
                    
                    Else '小数点文字の出現は2回目
                        ngIdx = cnt
                        Exit For
                        
                    End If
                
                Else '文字は数値文字でもなく、小数点文字でもない
                    ngIdx = cnt
                    Exit For
                    
                End If
                
            Else '文字はA~F
                numOfDigits = numOfDigits + 1 'increment
            End If
            
        Else '文字は0~9
            numOfDigits = numOfDigits + 1 'increment
        End If
        
    Next cnt
    
    If (numOfDigits = 0) And (ngIdx = 0) Then '数値が見つからない場合
        ngIdx = cnt - 1
        
    End If
    
    If (ngIdx > 0) Then 'NG文字が存在する場合
        errCode = CVErr(xlErrNum) '#NUM!を格納する
        checkBaseNNumber = ngIdx 'NG文字位置を返却
        
    Else 'すべてOKな場合
    
        idxOfDot = foundIdxOfDot
        errCode = 0
        checkBaseNNumber = cnt '文字列長 + 1を返却
        
    End If
    
End Function

'
'2数を和算する
'
'!CAUTION!
'    val1, val2 が有効なn進値であるかはチェックしない
'    radixは2~16の範囲内である事はチェックしない
'
Private Function add(ByVal val1 As String, ByVal val2 As String, ByVal radix As Byte) As String
    
    '変数宣言
    Dim lenOfVal1 As Long
    Dim lenOfVal2 As Long
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
        
        '繰り上がりチェック
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
'val1からval2を減算する
'減算結果が(-)値の場合は、resultIsMinusにTRUEを格納する
'減算結果が(+)値の場合は、resultIsMinusにFALSEを格納する
'
'!CAUTION!
'    val1, val2 が有効なn進値であるかはチェックしない
'    radixは2~16の範囲内である事はチェックしない
'
Private Function subtract(ByVal val1 As String, ByVal val2 As String, ByVal radix As Byte, ByRef resultIsMinus As Boolean) As String
    
    '変数宣言
    Dim idxMxOfVal As Long
    Dim diffIdx As Long
    Dim val1IsLarger As Integer '0:不明, 1:yes, -1:no
    Dim lenOfVal1 As Long
    Dim lenOfVal2 As Long
    Dim idxOfVal As Long
    Dim stringBuilder() As String
    Dim decDigitOfVal1 As Integer
    Dim decDigitOfVal2 As Integer
    Dim decCarrier As Integer
    Dim decDigitOfAns  As Integer
    
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
    
    '<大小比較チェック>--------------------------------------------------------------------
    
    diffIdx = 1
    val1IsLarger = 0
    Do While (diffIdx <= idxOfVal)
        
        decDigitOfVal1 = convNCharToByte(Mid(val1, diffIdx, 1))
        decDigitOfVal2 = convNCharToByte(Mid(val2, diffIdx, 1))
        
        'どちらかが大きかったら break
        If decDigitOfVal1 > decDigitOfVal2 Then
            val1IsLarger = 1
            Exit Do
        
        ElseIf decDigitOfVal1 < decDigitOfVal2 Then
            val1IsLarger = -1
            Exit Do
        
        End If
        
        diffIdx = diffIdx + 1
        
    Loop
    
    'val1とval2数は同じ数値の場合
    If (val1IsLarger = 0) Then
        resultIsMinus = False
        subtract = String(idxOfVal, "0") '0を返却
        Exit Function
        
    End If
    
    
    If (val1IsLarger = 1) Then 'val1の方が大きい数値の場合
        resultIsMinus = False '(+)を格納
        
    Else 'val2の方が大きい数値の場合
        
        '2数を入れ替える
        buf = val1
        val1 = val2
        val2 = buf
        
        resultIsMinus = True '(-)を格納
        
    End If
    
    '-------------------------------------------------------------------</大小比較チェック>
    
    'ループ前初期化
    ReDim stringBuilder(idxOfVal) '領域確保
    decCarrier = 0
    
    '解の生成ループ
    Do While (idxOfVal > 0)
        
        '対象桁の減算
        decDigitOfVal1 = convNCharToByte(Mid(val1, idxOfVal, 1))
        decDigitOfVal2 = convNCharToByte(Mid(val2, idxOfVal, 1))
        
        '繰り下がりチェック
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
        
        stringBuilder(idxOfVal) = convByteToNChar(decDigitOfAns) '解を格納
        
        idxOfVal = idxOfVal - 1 'decrement
        
    Loop
    
    '最上位桁格納
    stringBuilder(idxOfVal) = ""
    
    subtract = Join(stringBuilder, vbNullString)
    
    
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
        
        digitOfAns = convRadixOfInt(decDigitOfAns, 10, radix) '10進→n進変換
        
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
    
    tmp = convRadixOfInt(divisor, radix, 10)
    
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
        
        remainder = convRadixOfInt(rmnd, 10, radix)
    
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
Private Function convRadixOfInt(ByVal intStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte) As String
    
    Dim stsOfSub As Variant
    Dim retOfTryConvBasicRadix As String
    Dim strLenOfToRadix As Long
    Dim stringBuilder() As String '変換後文字列生成用
    Dim sizeOfStringBuilder As Long
    Dim rm As String
    
    intStr = removeLeft0(intStr)
    
    '変換する基数が同じ場合
    If (fromRadix = toRadix) Then
        convRadixOfInt = intStr '"0"を取り除いただけの値を返す
        Exit Function
        
    End If
    
    'convBasicRadixで解決可能かどうかチェック
    retOfTryConvBasicRadix = tryConvBasicRadix(intStr, fromRadix, toRadix, stsOfSub)
    
    If (retOfTryConvBasicRadix <> "") Then '基数変換用テーブルに解があった
        convRadixOfInt = retOfTryConvBasicRadix
        Exit Function
        
    End If
    
    '生成ループ前初期化
    sizeOfStringBuilder = 0
    chOfToRadix = convBasicRadix(fromRadix, toRadix)
    strLenOfToRadix = Len(chOfToRadix)
    
    '生成ループ - toRadixによる除算によって解を求める -
    Do While True
        
        If (Len(intStr) <= strLenOfToRadix) Then
            
            retOfTryConvBasicRadix = tryConvBasicRadix(intStr, fromRadix, 10, stsOfSub)
            
            If (retOfTryConvBasicRadix <> "") Then
                
                If (CByte(retOfTryConvBasicRadix) < toRadix) Then '基数で割れる数がなくなった ※必ず retOfTryConvBasicRadix > 0 とはなる ※
                    Exit Do
                    
                End If
                
            End If
            
        End If
        
        intStr = divide(intStr, chOfToRadix, fromRadix, 0, rm, stsOfSub) '16(10進値)以下による除算なので、オーバーフローは発生し得ない
        
        intStr = removeLeft0(intStr) '左側の不要な"0"を取り除く
        
        ReDim Preserve stringBuilder(sizeOfStringBuilder) '領域拡張
        
        '剰余を(toRadix)進値に変換した結果が算出Digit
        stringBuilder(sizeOfStringBuilder) = convRadixOfInt(rm, fromRadix, toRadix)
        
        sizeOfStringBuilder = sizeOfStringBuilder + 1
        
    Loop
    
    ReDim Preserve stringBuilder(sizeOfStringBuilder) '領域拡張
    
    '最上位Bit付加
    '剰余を(toRadix)進値に変換した結果が算出Digit
    stringBuilder(sizeOfStringBuilder) = convRadixOfInt(intStr, fromRadix, toRadix)
    
    convRadixOfInt = Join(invertStringArray(stringBuilder), vbNullString) '文字列連結
    
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
'    numOfDigitsが0以上であるかはチェックしない
'
Private Function convRadixOfFrc(ByVal frcStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByVal numOfDigits As Long) As String
    
    Dim stsOfSub As Variant
    Dim stringBuilder() As String '変換後文字列生成用
    Dim sizeOfStringBuilder As Long
    Dim retOfMultiple As String
    
    frcStr = removeRight0(frcStr)
    
    '変換する基数が同じ場合
    If (fromRadix = toRadix) Then
        
        If (numOfDigits > 0) Then
            convRadixOfFrc = frcStr '"0"を取り除いただけの値を返す
        Else
            convRadixOfFrc = "" '空文字を返す
        End If
        
        Exit Function
    End If
    
    '"0"を変換する場合
    If (frcStr = "0") Then
        
        If (numOfDigits > 0) Then
            convRadixOfFrc = "0" '"0"を返す
        Else
            convRadixOfFrc = "" '空文字を返す
        End If
        
        Exit Function
    End If
    
    '生成ループ前初期化
    strOfToRadix = convBasicRadix(fromRadix, toRadix)
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
            increasedDigits = convRadixOfInt(tmp, fromRadix, toRadix)
            
        Else
            increasedDigits = "0"
        
        End If
        
        stringBuilder(sizeOfStringBuilder) = increasedDigits '解を追記
        
        frcStr = removeRight0(frcStr) ' 右側の不要な0を取り除く
        
        lenOfFrcStrB = Len(frcStr)
        sizeOfStringBuilder = sizeOfStringBuilder + 1
        
    Loop
    
    convRadixOfFrc = Join(stringBuilder, vbNullString)

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
Private Function convBasicRadix(ByVal fromRadix As Byte, ByVal toRadix As Byte) As String
    
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
    
    convBasicRadix = radixTable(fromRadix)(toRadix)
    
End Function

'
'convBasicRadixを使ってN進→N進変換をトライする
'変換成功の場合は、変換後のN進値を返す
'失敗の場合は、endStatusに#N/A!を格納し、空文字を返す
'
'!CAUTION!
'    fromRadix,toRadixは2~16の範囲内である事はチェックしない
'
Private Function tryConvBasicRadix(ByVal intStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByRef endStatus As Variant) As String
    
    Dim idxOfRTable As Byte
    Dim toRetStr As String
    
    'convBasicRadixで解決可能かどうかチェック
    For idxOfRTable = 0 To 16
        If (intStr = convBasicRadix(fromRadix, idxOfRTable)) Then '基数変換用テーブルに解があった
            toRetStr = convBasicRadix(toRadix, idxOfRTable) '基数変換テーブルから解を返す
            Exit For
            
        End If
        
    Next idxOfRTable
    
    If (idxOfRTable > 16) Then '見つからなかった場合
        endStatus = CVErr(xlErrNA)
        toRetStr = ""
    
    End If
    
    tryConvBasicRadix = toRetStr
    
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


