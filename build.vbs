'==========================================================================
' build_final.vbs - Create proper .xlam using SaveAs format 55
' Previous issue: copying .xlsm->.xlam created invalid ZIP structure
' Fix: use WPS SaveAs with xlOpenXMLAddIn (55) format directly
' VBA persistence: save .xlsm first, then SaveAs .xlam
'==========================================================================
Option Explicit

Dim fso, app, wb, vbp, comp, srcDir, scriptDir, xlamPath, xlsmPath, startupPath
Dim file, count, i
Dim twComp, twCodeMod, ai
Dim rStream, basContent, txtFile, lineText
Dim foundHideSub, foundToolbar, twLines

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
xlamPath = scriptDir & "\FinanceRPA.xlam"
xlsmPath = scriptDir & "\FinanceRPA_temp.xlsm"
srcDir = scriptDir & "\src"

If Not fso.FolderExists(srcDir) Then
    WScript.Echo "ERROR: src/ not found"
    WScript.Quit(1)
End If

WScript.Echo "=== Build Final (SaveAs xlam) ==="

' === Step 1: Create WPS ===
On Error Resume Next
Set app = CreateObject("Ket.Application")
If app Is Nothing Then Set app = CreateObject("ET.Application")
If app Is Nothing Then Set app = CreateObject("Excel.Application")
On Error GoTo 0

If app Is Nothing Then
    WScript.Echo "ERROR: Cannot create WPS - close WPS first"
    WScript.Quit(1)
End If

app.Visible = False
app.DisplayAlerts = False

' === Step 2: Uninstall add-in ===
On Error Resume Next
For Each ai In app.AddIns
    If InStr(ai.Name, "FinanceRPA") > 0 Then
        If ai.Installed Then ai.Installed = False
    End If
Next
On Error GoTo 0

' === Step 3: Create new workbook ===
Set wb = app.Workbooks.Add()
WScript.Echo "Step 1: New workbook created"

' Remove extra sheets (keep only 1)
On Error Resume Next
Do While wb.Sheets.Count > 1
    wb.Sheets(wb.Sheets.Count).Delete
Loop
On Error GoTo 0

' === Step 4: Access VBA ===
On Error Resume Next
Set vbp = wb.VBProject
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot access VBA - check macro security"
    wb.Close False
    app.Quit
    WScript.Quit(1)
End If
On Error GoTo 0
WScript.Echo "Step 2: VBA accessible"

' === Step 5: Remove default sheet module code ===
Dim totalComps, j
totalComps = vbp.VBComponents.Count
Dim compArray()
ReDim compArray(totalComps - 1)
For i = 1 To totalComps
    Set compArray(i - 1) = vbp.VBComponents.Item(i)
Next

For i = totalComps - 1 To 0 Step -1
    If LCase(compArray(i).Name) <> "thisworkbook" Then
        On Error Resume Next
        vbp.VBComponents.Remove compArray(i)
        If Err.Number <> 0 Then
            Dim sheetCodeMod
            Set sheetCodeMod = compArray(i).CodeModule
            If sheetCodeMod.CountOfLines > 0 Then
                sheetCodeMod.DeleteLines 1, sheetCodeMod.CountOfLines
            End If
            Err.Clear
        End If
        On Error GoTo 0
    End If
Next

' === Step 6: Replace ThisWorkbook ===
On Error Resume Next
Set twComp = vbp.VBComponents("ThisWorkbook")
If Err.Number = 0 Then
    Set twCodeMod = twComp.CodeModule
    If twCodeMod.CountOfLines > 0 Then
        twCodeMod.DeleteLines 1, twCodeMod.CountOfLines
    End If
    twCodeMod.AddFromString "Private Sub Workbook_Open()" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    AutoSetupToolbar" & vbCrLf & _
        "    On Error GoTo 0" & vbCrLf & _
        "End Sub"
    WScript.Echo "Step 3: ThisWorkbook = " & twCodeMod.CountOfLines & " lines"
End If
Err.Clear
On Error GoTo 0

' === Step 7: Convert .bas to GBK ===
On Error Resume Next
For Each file In fso.GetFolder(srcDir).Files
    If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
        Set rStream = CreateObject("ADODB.Stream")
        rStream.Type = 2
        rStream.Charset = "utf-8"
        rStream.Open
        rStream.LoadFromFile file.Path
        basContent = rStream.ReadText
        rStream.Close
        Set rStream = Nothing

        If InStr(basContent, ChrW(&HFFFD)) > 0 Then
            Set rStream = CreateObject("ADODB.Stream")
            rStream.Type = 2
            rStream.Charset = "gbk"
            rStream.Open
            rStream.LoadFromFile file.Path
            basContent = rStream.ReadText
            rStream.Close
            Set rStream = Nothing
        Else
            Set txtFile = fso.CreateTextFile(file.Path, True, False)
            txtFile.Write basContent
            txtFile.Close
            Set txtFile = Nothing
        End If
    End If
Next
On Error GoTo 0

' === Step 8: Import .bas files ===
count = 0
On Error Resume Next
For Each file In fso.GetFolder(srcDir).Files
    If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
        Dim importedComp
        Set importedComp = vbp.VBComponents.Import(file.Path)
        If Err.Number = 0 Then
            count = count + 1
            WScript.Echo "  OK: " & file.Name
        Else
            WScript.Echo "  FAIL: " & file.Name
        End If
        Err.Clear
    End If
Next
On Error GoTo 0
WScript.Echo "Step 4: Imported " & count & " modules"

' === Step 9: Save as xlsm first (preserves VBA) ===
If fso.FileExists(xlsmPath) Then fso.DeleteFile xlsmPath, True
On Error Resume Next
wb.SaveAs xlsmPath, 52
If Err.Number <> 0 Then
    WScript.Echo "ERROR: SaveAs xlsm failed: " & Err.Description
    wb.Close False
    app.Quit
    WScript.Quit(1)
End If
On Error GoTo 0
WScript.Echo "Step 5: Saved xlsm (VBA preserved)"

' === Step 10: Set IsAddin=True, then SaveAs as xlam (format 55) ===
' Must set IsAddin=True BEFORE SaveAs so workbook.xml gets isAddin="true"
' VBA is already persisted from the xlsm save above
On Error Resume Next
If fso.FileExists(xlamPath) Then
    fso.DeleteFile xlamPath, True
End If
Err.Clear
On Error GoTo 0

' Set IsAddin = True (VBA already saved, this only affects add-in property)
On Error Resume Next
wb.IsAddin = True
If Err.Number <> 0 Then
    WScript.Echo "  WARNING: Cannot set IsAddin"
End If
Err.Clear
On Error GoTo 0

' Try SaveAs with format 55
Dim saveAsOk
saveAsOk = False
On Error Resume Next
wb.SaveAs xlamPath, 55
If Err.Number = 0 Then
    saveAsOk = True
    WScript.Echo "Step 6: SaveAs xlam (format 55) OK"
Else
    WScript.Echo "  Format 55 failed: " & Err.Description & ", trying 18..."
    Err.Clear
    wb.SaveAs xlamPath, 18
    If Err.Number = 0 Then
        saveAsOk = True
        WScript.Echo "Step 6: SaveAs xlam (format 18) OK"
    Else
        WScript.Echo "  Format 18 also failed: " & Err.Description
    End If
End If
On Error GoTo 0

If Not saveAsOk Then
    WScript.Echo "ERROR: Cannot save as xlam"
    wb.Close False
    app.Quit
    WScript.Quit(1)
End If

wb.Close False
WScript.Sleep 1000

' === Step 11: Verify the xlam ===
On Error Resume Next
Set wb = app.Workbooks.Open(xlamPath)
If Err.Number = 0 Then
    Set vbp = wb.VBProject
    If Err.Number = 0 Then
        foundHideSub = False
        foundToolbar = False
        twLines = 0

        For Each comp In vbp.VBComponents
            Set twCodeMod = comp.CodeModule
            For i = 1 To twCodeMod.CountOfLines
                lineText = twCodeMod.Lines(i, 1)
                If InStr(lineText, "btnHideFixedAssetColumns") > 0 Then foundHideSub = True
                If InStr(lineText, "AutoSetupToolbar") > 0 Then foundToolbar = True
                If LCase(comp.Name) = "thisworkbook" Then twLines = twCodeMod.CountOfLines
            Next
        Next

        WScript.Echo "Step 7: Verification"
        WScript.Echo "  Components: " & vbp.VBComponents.Count
        WScript.Echo "  ThisWorkbook: " & twLines & " lines"
        WScript.Echo "  btnHideFixedAssetColumns: " & IIf(foundHideSub, "FOUND", "MISSING!")
        WScript.Echo "  AutoSetupToolbar: " & IIf(foundToolbar, "FOUND", "MISSING!")
    End If
    wb.Close False
End If
On Error GoTo 0

' === Step 11.5: Fix isAddin in workbook.xml via PowerShell ===
' WPS SaveAs format 55 doesn't write isAddin="true" to workbook.xml
' Must patch the ZIP directly or WPS shows "invalid add-in"
Dim shell, retCode
Set shell = CreateObject("WScript.Shell")
retCode = shell.Run("powershell -ExecutionPolicy Bypass -File """ & scriptDir & "\fix_isaddin.ps1"" """ & xlamPath & """", 0, True)
WScript.Echo "Step 6.5: Fixed isAddin in workbook.xml (exit code: " & retCode & ")"

' === Step 12: Copy to startup ===
On Error Resume Next
For Each ai In app.AddIns
    If InStr(ai.Name, "FinanceRPA") > 0 Then ai.Installed = True
Next
On Error GoTo 0

startupPath = app.StartupPath
app.Quit
WScript.Sleep 2000

If Len(startupPath) > 0 And fso.FolderExists(startupPath) Then
    If fso.FileExists(startupPath & "\FinanceRPA.xlam") Then
        fso.DeleteFile startupPath & "\FinanceRPA.xlam", True
    End If
    fso.CopyFile xlamPath, startupPath & "\FinanceRPA.xlam", True
    WScript.Echo "Copied to startup"
End If

' Clean up temp
If fso.FileExists(xlsmPath) Then fso.DeleteFile xlsmPath, True

WScript.Echo ""
WScript.Echo "=== DONE ==="
If foundHideSub And foundToolbar Then
    WScript.Echo "Status: ALL VERIFIED - ready to test"
Else
    WScript.Echo "Status: WARNING - verification failed"
End If

Function IIf(expr, trueVal, falseVal)
    If expr Then
        IIf = trueVal
    Else
        IIf = falseVal
    End If
End Function
