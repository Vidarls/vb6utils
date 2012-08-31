Attribute VB_Name = "List"
Option Explicit

Public Function From(enumerable) As Lst
  Dim item As Variant
  Dim newList As Lst
  Set newList = New Lst
  For Each item In enumerable
    newList.Add item
  Next item
  Set From = newList
End Function
