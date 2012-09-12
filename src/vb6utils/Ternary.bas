Attribute VB_Name = "Ternary"
Option Explicit

Public Function Choose(expression As Boolean, trueValue As Variant, falseValue As Variant) As Variant
  If expression Then
    If IsObject(trueValue) Then
      Set Choose = trueValue
    Else
      Choose = trueValue
    End If
    Exit Function
  End If
  
  If IsObject(falseValue) Then
    Set Choose = falseValue
  Else
    Choose = falseValue
  End If
End Function

Public Function DbNullCoalesce(fieldReference As Variant, defaultIfNull) As Variant
  DbNullCoalesce = Choose(IsNull(fieldReference), defaultIfNull, fieldReference)
End Function
