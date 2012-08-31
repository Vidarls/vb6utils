VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "string_format_tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim localTests As Collection
Dim helper As TestFixtureHelper
Dim Assert As Asserter

Const ARBITRARY_ERROR_CODE As Long = vbObjectError

Public Property Get Tests() As Collection
  Set Tests = localTests
End Property

Public Function Create() As string_format_tests
  Set Create = New string_format_tests
End Function

Private Function Fail() As TestErr
  Set Fail = helper.Fail
End Function

Private Sub Class_initialize()
  Set localTests = New Collection
  Set Assert = New Asserter
  Set helper = New TestFixtureHelper
  
End Sub

Private Sub Class_terminate()
  Set localTests = Nothing
  Set Assert = Nothing
End Sub

