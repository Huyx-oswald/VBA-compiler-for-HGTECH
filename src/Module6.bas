Attribute VB_Name = "Module6"
Sub btnHideFixedAssetColumns()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Check sheet protection
    If ws.ProtectContents Then
        MsgBox "Sheet is protected! Please unprotect first.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Show current sheet name for diagnostic
    Dim sheetName As String
    sheetName = ws.Name

    Dim hideRanges
    hideRanges = Array( _
        "A:D", _
        "G:G", _
        "I:I", _
        "L:R", _
        "U:W", _
        "Y:Y", _
        "AA:AE", _
        "AH:AJ", _
        "AL:AM", _
        "AO:AP" _
    )

    Dim i As Integer
    Dim hideCount As Integer
    Dim errMsg As String
    hideCount = 0
    errMsg = ""

    For i = LBound(hideRanges) To UBound(hideRanges)
        On Error Resume Next
        ws.Range(hideRanges(i)).EntireColumn.Hidden = True
        If Err.Number = 0 Then
            hideCount = hideCount + 1
        Else
            errMsg = errMsg & "Range " & hideRanges(i) & " ERROR: " & Err.Description & vbCrLf
            Err.Clear
        End If
        On Error GoTo 0
    Next

    Dim msg As String
    msg = "Sheet: " & sheetName & vbCrLf
    msg = msg & "Hidden: " & hideCount & " of " & (UBound(hideRanges) + 1) & " ranges"
    If Len(errMsg) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Errors:" & vbCrLf & errMsg
    End If

    MsgBox msg, vbInformation, "Hide Columns Result"
End Sub
