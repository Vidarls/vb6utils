Attribute VB_Name = "Strng"
Option Explicit

Public Function Frmat(ByVal formatString As String, ParamArray args() As Variant) As String
  Dim placeholders As Collection
  Dim total As Integer
  Dim i As Integer
  Dim stringValue As String
  
  Frmat = formatString
  If IsMissing(args) Then Exit Function
  
  Set placeholders = List.From(FindPlaceholders(formatString)).Distinct().ToCollection()
  If placeholders.Count < 1 Then Exit Function
  
  total = placeholders.Count - 1
  If UBound(args) < total Then total = UBound(args)
  For i = 0 To total
    If IsObject(args(i)) Then stringValue = "Not supported"
    stringValue = CStr(args(i))
    formatString = Replace(formatString, placeholders(i + 1), stringValue)
  Next i
  Frmat = formatString
End Function

Private Function FindPlaceholders(text As String) As Collection
   Dim startPosition As Long
   Dim stopPosition As Long
   Dim done As Boolean
   Dim placeholders As New Collection
   Dim i As Integer
   
   Set FindPlaceholders = placeholders
   
   startPosition = InStr(1, text, "{")
   If startPosition = 0 Then Exit Function
   
   stopPosition = InStr(startPosition, text, "}")
   If stopPosition = 0 Then Exit Function
   
   placeholders.Add mid(text, startPosition, ((stopPosition + 2) - startPosition))
End Function

