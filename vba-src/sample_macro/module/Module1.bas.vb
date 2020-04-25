Option Explicit

Sub Test()
    ThisWorkbook.Worksheets(1).Range("A1").Value = "Hello world!!"
    MsgBox "Hello world!"
    MsgBox "Test"
End Sub