VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "map_reduce_tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim localTests As Collection
Dim helper As TestFixtureHelper
Dim Assert As Asserter

Private testColllection As Collection
Private counter As Integer

Const ARBITRARY_ERROR_CODE As Long = vbObjectError

Public Property Get Tests() As Collection
  Set Tests = localTests
End Property

Public Function Create() As map_reduce_tests
  Set Create = New map_reduce_tests
End Function

Private Function Fail() As TestErr
  Set Fail = helper.Fail
End Function

Private Sub Class_initialize()
  Set localTests = New Collection
  Set Assert = New Asserter
  Set helper = New TestFixtureHelper
  
  Set testColllection = CreateTestCollection
  
  localTests.Add "Map_should_call_mapping_function_for_each_element"
  localTests.Add "Map_should_set_each_element_to_return_from_mapping_function"
End Sub

Public Function Map_should_call_mapping_function_for_each_element()
On Error GoTo failed
  counter = 0
  Set testColllection = List.From(testColllection).Map(Me, "MapTester").ToCollection
  
  Call Assert.That(counter).Equals(3)
  
  Set Map_should_call_mapping_function_for_each_element = Nothing
Exit Function

failed:
  Set Map_should_call_mapping_function_for_each_element = Fail
End Function

Public Function Map_should_set_each_element_to_return_from_mapping_function()
On Error GoTo failed
  Dim item As Variant
  Set testColllection = List.From(testColllection).Map(Me, "MapTester").ToCollection
  For Each item In testColllection
    Call Assert.That(item).Equals(-1)
  Next item
  
  Set Map_should_set_each_element_to_return_from_mapping_function = Nothing
Exit Function

failed:
  Set Map_should_set_each_element_to_return_from_mapping_function = Fail
End Function


Public Function MapTester(entry) As Variant
  counter = counter + 1
  MapTester = -1
End Function

Private Function CreateTestCollection() As Collection
  Dim newCol As New Collection
  newCol.Add 1
  newCol.Add 2
  newCol.Add 3
  Set CreateTestCollection = newCol
End Function

Private Sub Class_terminate()
  Set localTests = Nothing
  Set Assert = Nothing
End Sub

