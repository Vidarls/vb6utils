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

Public Function DefaultIfNull(value As Variant, Optional defaultValue) As Variant
  If IsMissing(defaultValue) Then
    defaultValue = GetDefaultNonNullValueForTypeOfGivenVariable(value)
  End If
  DefaultIfNull = Choose(ValueIsNullNothingOrEmpty(value), defaultValue, value)
End Function

Public Function ValueIsNullNothingOrEmpty(value As Variant) As Boolean
  ValueIsNullNothingOrEmpty = True
  If typename(value) = "Nothing" Then Exit Function
  If typename(value) = "Empty" Then Exit Function
  If IsNull(value) Then Exit Function
  If IsEmpty(value) Then Exit Function
  If IsObject(value) Then
    If value Is Nothing Then Exit Function
  End If
  ValueIsNullNothingOrEmpty = False
End Function

Private Function GetDefaultNonNullValueForTypeOfGivenVariable(givenValue As Variant) As Variant
  If typename(givenValue) = "Field" Then
    GetDefaultNonNullValueForTypeOfGivenVariable = GetDefaultNonNullValueForVarType(AdoTypeToVbType(givenValue.Type), typename(givenValue))
    Exit Function
  End If
  
  GetDefaultNonNullValueForTypeOfGivenVariable = GetDefaultNonNullValueForVarType(vartype(givenValue), typename(givenValue))
End Function

Private Function GetDefaultNonNullValueForVarType(vartypeValue As Integer, typename As String) As Variant
Select Case vartypeValue
    Case vbInteger
      GetDefaultNonNullValueForVarType = CInt(0)
    Case vbLong
      GetDefaultNonNullValueForVarType = CLng(0)
    Case vbSingle
      GetDefaultNonNullValueForVarType = CSng(0)
    Case vbDouble
      GetDefaultNonNullValueForVarType = CDbl(0)
    Case vbCurrency
      GetDefaultNonNullValueForVarType = CCur(0)
    Case vbDate
      Dim d As Date
      GetDefaultNonNullValueForVarType = d
    Case vbString
      GetDefaultNonNullValueForVarType = ""
    Case vbBoolean
      GetDefaultNonNullValueForVarType = False
    Case vbDecimal
      GetDefaultNonNullValueForVarType = CDec(0)
    Case vbByte
      GetDefaultNonNullValueForVarType = CByte(0)
    Case Else
      Err.Raise 13, "Ternary.DefaultIfNull", "Can not infer non-null default value for variables of type " + typename + ". Please provide explicit non null default value"
  End Select
End Function

Private Function AdoTypeToVbType(adoType As ADODB.DataTypeEnum) As Integer
  Select Case adoType
    Case adArray:             AdoTypeToVbType = vbArray
    Case adBigInt:            AdoTypeToVbType = vbLong
    Case adBinary:            AdoTypeToVbType = vbArray + vbByte
    Case adBoolean:           AdoTypeToVbType = vbBoolean
    Case adBSTR:              AdoTypeToVbType = vbString
    Case adChar:              AdoTypeToVbType = vbString
    Case adCurrency:          AdoTypeToVbType = vbCurrency
    Case adDate:              AdoTypeToVbType = vbDate
    Case adDBDate:            AdoTypeToVbType = vbDate
    Case adDBTime:            AdoTypeToVbType = vbDate
    Case adDBTimeStamp:       AdoTypeToVbType = vbDate
    Case adDecimal:           AdoTypeToVbType = vbDecimal
    Case adDouble:            AdoTypeToVbType = vbDouble
    Case adEmpty:             AdoTypeToVbType = vbEmpty
    Case adInteger:           AdoTypeToVbType = vbInteger
    Case adLongVarBinary:     AdoTypeToVbType = vbArray + vbByte
    Case adLongVarChar:       AdoTypeToVbType = vbString
    Case adLongVarWChar:      AdoTypeToVbType = vbString
    Case adNumeric:           AdoTypeToVbType = vbLong
    Case adSingle:            AdoTypeToVbType = vbSingle
    Case adSmallInt:          AdoTypeToVbType = vbInteger
    Case adTinyInt:           AdoTypeToVbType = vbInteger
    Case adUnsignedBigInt:    AdoTypeToVbType = vbLong
    Case adUnsignedInt:       AdoTypeToVbType = vbInteger
    Case adUnsignedSmallInt:  AdoTypeToVbType = vbInteger
    Case adUnsignedTinyInt:   AdoTypeToVbType = vbInteger
    Case adVarBinary:         AdoTypeToVbType = vbArray + vbByte
    Case adVarChar:           AdoTypeToVbType = vbString
    Case adVariant:           AdoTypeToVbType = vbVariant
    Case adVarNumeric:        AdoTypeToVbType = vbLong
    Case adVarWChar:          AdoTypeToVbType = vbString
    Case adWChar:             AdoTypeToVbType = vbString
    Case Else
      Err.Raise 13, "Ternary.DefaultIfNull", "Can not infer default value from ado type #" + Str(adoType) + ". Please provide explicit default value"
  End Select
End Function
