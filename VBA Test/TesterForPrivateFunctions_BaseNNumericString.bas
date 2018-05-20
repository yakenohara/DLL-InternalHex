'
'以下関数を「BaseNNumericString.bas」に追加して、
'「Tester_BaseNNumericString.xlsm」にインポートします
'

'<PrivateFunction用テスト関数>---------------------------------------------------------------------------------------------------------------------
'
Public Function TESTseparateToIntAndFrc(ByVal baseNNumericStr As String, ByVal radix As Byte, ByVal remove0 As Boolean) As Variant
    Dim intPrt As String
    Dim frcPrt As String
    Dim isMinus As Boolean
    TESTseparateToIntAndFrc = separateToIntAndFrc(baseNNumericStr, radix, remove0, intPrt, frcPrt, isMinus)
End Function

Public Function TESTseparateToIntAndFrcByRef1(ByVal baseNNumericStr As String, ByVal radix As Byte, ByVal remove0 As Boolean) As Variant
    Dim intPrt As String
    Dim frcPrt As String
    Dim isMinus As Boolean
    x = separateToIntAndFrc(baseNNumericStr, radix, remove0, intPrt, frcPrt, isMinus)
    TESTseparateToIntAndFrcByRef1 = intPrt
End Function

Public Function TESTseparateToIntAndFrcByRef2(ByVal baseNNumericStr As String, ByVal radix As Byte, ByVal remove0 As Boolean) As Variant
    Dim intPrt As String
    Dim frcPrt As String
    Dim isMinus As Boolean
    x = separateToIntAndFrc(baseNNumericStr, radix, remove0, intPrt, frcPrt, isMinus)
    TESTseparateToIntAndFrcByRef2 = frcPrt
End Function

Public Function TESTseparateToIntAndFrcByRef3(ByVal baseNNumericStr As String, ByVal radix As Byte, ByVal remove0 As Boolean) As Variant
    Dim intPrt As String
    Dim frcPrt As String
    Dim isMinus As Boolean
    x = separateToIntAndFrc(baseNNumericStr, radix, remove0, intPrt, frcPrt, isMinus)
    TESTseparateToIntAndFrcByRef3 = isMinus
End Function

Public Function TESTcheckBaseNNumber(ByVal baseNNumericStr As String, ByVal radix As Byte) As Variant
    Dim isMinus As Boolean
    Dim idxOfDot As Long
    Dim stsOfSub As Variant
    TESTcheckBaseNNumber = checkBaseNNumber(baseNNumericStr, radix, isMinus, idxOfDot, stsOfSub)
End Function

Public Function TESTcheckBaseNNumberByRef1(ByVal baseNNumericStr As String, ByVal radix As Byte) As Variant
    Dim isMinus As Boolean
    Dim idxOfDot As Long
    Dim stsOfSub As Variant
    x = checkBaseNNumber(baseNNumericStr, radix, isMinus, idxOfDot, stsOfSub)
    TESTcheckBaseNNumberByRef1 = isMinus
End Function

Public Function TESTcheckBaseNNumberByRef2(ByVal baseNNumericStr As String, ByVal radix As Byte) As Variant
    Dim isMinus As Boolean
    Dim idxOfDot As Long
    Dim stsOfSub As Variant
    x = checkBaseNNumber(baseNNumericStr, radix, isMinus, idxOfDot, stsOfSub)
    TESTcheckBaseNNumberByRef2 = idxOfDot
End Function

Public Function TESTcheckBaseNNumberByRef3(ByVal baseNNumericStr As String, ByVal radix As Byte) As Variant
    Dim isMinus As Boolean
    Dim idxOfDot As Long
    Dim stsOfSub As Variant
    x = checkBaseNNumber(baseNNumericStr, radix, isMinus, idxOfDot, stsOfSub)
    TESTcheckBaseNNumberByRef3 = stsOfSub
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

Public Function TESTconvRadixOfInt(ByVal intStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte) As Variant
    TESTconvRadixOfInt = convRadixOfInt(intStr, fromRadix, toRadix)
End Function

Public Function TESTconvRadixOfFrc(ByVal frcStr As String, ByVal fromRadix As Byte, ByVal toRadix As Byte, ByVal numOfDigits As Long) As Variant
    TESTconvRadixOfFrc = convRadixOfFrc(frcStr, fromRadix, toRadix, numOfDigits)
End Function
'
'--------------------------------------------------------------------------------------------------------------------</PrivateFunction用テスト関数>
