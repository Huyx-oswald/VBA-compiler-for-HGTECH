Attribute VB_Name = "ThisWorkbook"
Private Sub Workbook_Open()
    On Error Resume Next
    AutoSetupToolbar
    On Error GoTo 0
End Sub
