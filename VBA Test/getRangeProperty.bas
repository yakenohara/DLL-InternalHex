Attribute VB_Name = "getRangeProperty"
Public Function getText(ByVal rng As Range) As Variant
    
    getText = rng.Text
    
End Function

Public Function getFormula(ByVal rng As Range) As Variant
    
    getFormula = rng.Formula
    
End Function
