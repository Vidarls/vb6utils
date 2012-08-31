Attribute VB_Name = "tests"
Public Sub Main()
  Dim runner As New TestRunner
  Set runner = New TestRunner
  Call runner.AddFixture(New string_format_tests)
  Call runner.AddFixture(New map_reduce_tests)
  Call runner.ShowAndRun(1)
End Sub
