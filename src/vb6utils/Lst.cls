VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private internalCollection As Collection
Private internalConcatenationSeparator As String
Private internalPropertyName As String
Private internalDistinctCollection As Collection

Private Function item(ByVal Index As Variant) As Variant
  If IsObject(internalCollection(Index)) Then
    Set item = internalCollection(Index)
  Else
    item = internalCollection(Index)
  End If
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
  Set NewEnum = internalCollection.[_NewEnum]
End Property

Public Property Get Count() As Long
  Count = internalCollection.Count
End Property

Private Sub Class_initialize()
  Set internalCollection = New Collection
  internalConcatenationSeparator = ""
End Sub

Public Sub Add(item As Variant)
  internalCollection.Add item
End Sub

Private Sub Remove(Index As Variant)
  internalCollection.Remove Index
End Sub

Private Sub Class_terminate()
  Set internalCollection = Nothing
End Sub

Public Function Map(objectContainingMappingFunction, nameOfMappingFunction) As Lst
  Dim newList As Lst
  Dim entry As Variant
  Set newList = New Lst
  For Each entry In internalCollection
    newList.Add CallByName(objectContainingMappingFunction, nameOfMappingFunction, VbMethod, entry)
  Next entry
  Set Map = newList
End Function

Public Function Fold(objectContainingFoldingFunction, nameOfFoldingFunction, ByVal initialValue) As Variant
  Dim foldingResult
  Dim entry As Variant
  
  foldingResult = initialValue
  For Each entry In internalCollection
    If IsObject(foldingResult) Then
      Set foldingResult = CallByName(objectContainingFoldingFunction, nameOfFoldingFunction, VbMethod, foldingResult, entry)
    Else
      foldingResult = CallByName(objectContainingFoldingFunction, nameOfFoldingFunction, VbMethod, foldingResult, entry)
    End If
  Next entry
  
  If IsObject(foldingResult) Then
    Set Fold = foldingResult
  Else
    Fold = foldingResult
  End If
End Function

Public Function Filter(objectContainingFilterFunction, nameOfFilterFunction, ParamArray args() As Variant) As Lst
  Dim filteredList As New Lst
  Dim entry As Variant
  
  For Each entry In internalCollection
    If CallByName(objectContainingFilterFunction, nameOfFilterFunction, VbMethod, entry, args) Then Call filteredList.Add(entry)
  Next entry
  Set Filter = filteredList
End Function

Public Function ToCollection() As Collection
  Set ToCollection = internalCollection
End Function

Public Function Concat(Optional separator As String = "") As String
  internalConcatenationSeparator = separator
  Concat = CStr(Me.Fold(Me, "internalConcatenator", ""))
End Function

Public Function Contains(entry As Variant) As Boolean
  Contains = (Filter(Me, "internalContainsFilter", entry).Count > 0)
End Function

Public Function Except(somethingEnumerable As Variant) As Lst
  Dim newList As New Lst
  Dim entry As Variant
  Dim comparisonList As Lst
  Set comparisonList = List.From(somethingEnumerable)
  
  For Each entry In internalCollection
    If Not comparisonList.Contains(entry) Then newList.Add entry
  Next entry
  
  Set Except = newList
End Function

Public Function Intersect(somethingEnumerable As Variant) As Lst
  Dim newList As New Lst
  Dim entry As Variant
  Dim comparisonList As Lst
  Set comparisonList = List.From(somethingEnumerable)
  
  For Each entry In internalCollection
    If comparisonList.Contains(entry) Then newList.Add entry
  Next entry
  
  Set Intersect = newList
End Function

Public Function Distinct() As Lst
  Set internalDistinctCollection = New Collection
  Set Distinct = Filter(Me, "internalDistinctFilter")
End Function

Public Function FirstOrDefault(Optional default As Variant = Empty) As Variant
  If internalCollection.Count < 1 Then
    If IsObject(default) Then
      Set FirstOrDefault = default
    Else
      FirstOrDefault = default
    End If
    Exit Function
  End If
  
  If IsObject(internalCollection(1)) Then
    Set FirstOrDefault = internalCollection(1)
  Else
    FirstOrDefault = internalCollection(1)
  End If
End Function

Public Function internalDistinctFilter(entry As Variant, args() As Variant) As Boolean
  internalDistinctFilter = Not List.From(internalDistinctCollection).Contains(entry)
End Function

Public Function internalContainsFilter(entry As Variant, args() As Variant) As Boolean
  internalContainsFilter = EqualityCheck(entry, args(0))
End Function

Public Function internalConcatenator(concatenated, entry) As Variant
  internalConcatenator = concatenated
  If IsObject(entry) Then
    If entry Is Nothing Then Exit Function
  End If
  If LenB(concatenated) = 0 Then
    concatenated = CStr(entry)
  Else
    concatenated = concatenated + internalConcatenationSeparator + CStr(entry)
  End If
  internalConcatenator = concatenated
End Function

Public Function SelectProperty(nameOfProperty As String) As Lst
  internalPropertyName = nameOfProperty
  Set SelectProperty = Me.Map(Me, "internalPropertySelector")
End Function

Public Function internalPropertySelector(entry) As Variant
  If IsObject(CallByName(entry, internalPropertyName, VbGet)) Then
    Set internalPropertySelector = CallByName(entry, internalPropertyName, VbGet)
  Else
    internalPropertySelector = CallByName(entry, internalPropertyName, VbGet)
  End If
End Function

Private Function EqualityCheck(this As Variant, other As Variant) As Boolean
  EqualityCheck = False
  If TypeName(this) <> TypeName(other) Then Exit Function
  If Not IsObject(this) Then
    EqualityCheck = (this = other)
    Exit Function
  End If
  EqualityCheck = CompareWithObjectEquals(this, other)
End Function

Private Function CompareWithObjectEquals(this As Variant, other As Variant) As Boolean
On Error GoTo failed
  CompareWithObjectEquals = False
  CompareWithObjectEquals = this.Equals(other)
Exit Function
failed:
  CompareWithObjectEquals = (this Is other)
End Function
