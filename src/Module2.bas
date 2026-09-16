Attribute VB_Name = "Module2"
Sub 库存价值表数据清洗()
    Dim wb As Workbook
    Dim dataWS As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object
    
    ' 使用 ActiveWorkbook 而不是 ThisWorkbook
    ' 这样会处理当前活动的工作簿，而不是包含VBA代码的工作簿
    Set wb = ActiveWorkbook
    
    ' 确认用户要处理的工作簿
    Dim response As VbMsgBoxResult
    response = MsgBox("即将处理工作簿: " & wb.Name & vbCrLf & _
                     "是否继续？", vbYesNo + vbQuestion, "确认工作簿")
    
    If response = vbNo Then
        MsgBox "操作已取消。"
        Exit Sub
    End If
    
    ' 创建字典存储物料组匹配关系
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 检查物料组模板文件是否存在
    Dim templateFilePath As String
    templateFilePath = "C:\Users\Finance005\Desktop\111华工图像同步文件\4、附表和简报库龄表\库龄表\物料组模板.xlsx"
    
    If Dir(templateFilePath) = "" Then
        MsgBox "错误：找不到物料组模板文件 '物料组模板.xlsx'，请确保文件在当前目录中。"
        Exit Sub
    End If
    
    ' 打开物料组模板文件并创建匹配字典
    Dim templateWB As Workbook
    On Error Resume Next
    Set templateWB = Workbooks.Open(templateFilePath)
    If Err.Number <> 0 Then
        MsgBox "错误：无法打开物料组模板文件，请检查文件是否被占用或损坏。"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim templateWS As Worksheet
    Set templateWS = templateWB.Sheets(1)
    
    Dim templateLastRow As Long
    templateLastRow = templateWS.Cells(templateWS.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To templateLastRow
        If Not IsEmpty(templateWS.Cells(i, 1)) And Not IsEmpty(templateWS.Cells(i, 2)) Then
            dict(CStr(templateWS.Cells(i, 1).value)) = CStr(templateWS.Cells(i, 2).value)
        End If
    Next i
    
    templateWB.Close SaveChanges:=False
    
    ' 检查主数据表是否存在
    On Error Resume Next
    Set dataWS = wb.Sheets("Sheet1")
    On Error GoTo 0
    
    If dataWS Is Nothing Then
        MsgBox "错误：在工作簿 '" & wb.Name & "' 中找不到名为 'Sheet1' 的工作表。" & vbCrLf & _
               "请确认工作表名称是否正确。"
        Exit Sub
    End If
    
    ' 确认要处理的工作表
    response = MsgBox("即将处理工作表: " & dataWS.Name & vbCrLf & _
                     "是否继续？", vbYesNo + vbQuestion, "确认工作表")
    
    If response = vbNo Then
        MsgBox "操作已取消。"
        Exit Sub
    End If
    
    lastRow = dataWS.Cells(dataWS.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "错误：工作表 '" & dataWS.Name & "' 中没有找到足够的数据。"
        Exit Sub
    End If
    
    ' 1. 添加物料组文本列（E列物料组后）
    dataWS.Cells(1, 5).EntireColumn.Insert
    dataWS.Cells(1, 5) = "物料组文本"
    
    For i = 2 To lastRow
        Dim materialGroupCode As String
        materialGroupCode = CStr(dataWS.Cells(i, 4).value)
        
        If dict.Exists(materialGroupCode) Then
            dataWS.Cells(i, 5) = dict(materialGroupCode)
        Else
            Dim foundMatch As Boolean
            foundMatch = False
            Dim key As Variant
            For Each key In dict.keys
                If Left(materialGroupCode, Len(CStr(key))) = CStr(key) Then
                    dataWS.Cells(i, 5) = dict(key)
                    foundMatch = True
                    Exit For
                End If
            Next key
            
            If Not foundMatch Then
                dataWS.Cells(i, 5) = "未找到匹配"
            End If
        End If
    Next i
    
    ' 2. 添加期初库存价值列
    dataWS.Cells(1, 12).EntireColumn.Insert
    dataWS.Cells(1, 12) = "期初库存价值"
    
    ' 检查期初数据文件是否存在
    Dim prevDataFilePath As String
    prevDataFilePath = "C:\Users\Finance005\Desktop\111华工图像同步文件\4、附表和简报库龄表\库龄表\库存价值表25年12月EXPORT.XLSX"
    
    If Dir(prevDataFilePath) = "" Then
        MsgBox "错误：找不到期初数据文件 '库存价值表25年12月EXPORT.XLSX'。"
        Exit Sub
    End If
    
    ' 打开25年12月数据进行匹配
    Dim prevDataWB As Workbook
    On Error Resume Next
    Set prevDataWB = Workbooks.Open(prevDataFilePath)
    If Err.Number <> 0 Then
        MsgBox "错误：无法打开期初数据文件。"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim prevDataWS As Worksheet
    Set prevDataWS = prevDataWB.Sheets(1)
    
    Dim prevLastRow As Long
    prevLastRow = prevDataWS.Cells(prevDataWS.Rows.Count, "B").End(xlUp).Row
    
    ' 创建字典存储上期数据
    Dim prevDataDict As Object
    Set prevDataDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To prevLastRow
        Dim materialNo As String
        materialNo = CStr(prevDataWS.Cells(i, 2).value)
        Dim value As Double
        If IsNumeric(prevDataWS.Cells(i, 11).value) Then
            value = CDbl(prevDataWS.Cells(i, 11).value)
        Else
            value = 0
        End If
        prevDataDict(materialNo) = value
    Next i
    
    prevDataWB.Close SaveChanges:=False
    
    ' 匹配期初库存价值
    For i = 2 To lastRow
        Dim currentMaterialNo As String
        currentMaterialNo = CStr(dataWS.Cells(i, 2).value)
        
        If prevDataDict.Exists(currentMaterialNo) Then
            dataWS.Cells(i, 12) = prevDataDict(currentMaterialNo)
        Else
            dataWS.Cells(i, 12) = 0
        End If
    Next i
    
    ' 3. 在B列物料编码前添加分类列
    dataWS.Cells(1, 2).EntireColumn.Insert
    dataWS.Cells(1, 2) = "分类"
    
    For i = 2 To lastRow
        Dim materialCode As String
        materialCode = CStr(dataWS.Cells(i, 3).value)
        
        If Len(materialCode) > 0 Then
            Select Case Left(materialCode, 1)
                Case "1"
                    dataWS.Cells(i, 2) = "原材料"
                Case "2"
                    dataWS.Cells(i, 2) = "半成品"
                Case "3"
                    dataWS.Cells(i, 2) = "成品"
                Case "4"
                    dataWS.Cells(i, 2) = "辅材"
                Case "5"
                    dataWS.Cells(i, 2) = "备品备件"
                Case "6"
                    dataWS.Cells(i, 2) = "其他"
                Case Else
                    dataWS.Cells(i, 2) = "其他"
            End Select
        Else
            dataWS.Cells(i, 2) = "未知"
        End If
    Next i
    
    ' 4. 创建原材料和成品表
    Dim rawMatWS As Worksheet
    On Error Resume Next
    Set rawMatWS = wb.Sheets("原材料")
    On Error GoTo 0
    
    If rawMatWS Is Nothing Then
        Set rawMatWS = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        rawMatWS.Name = "原材料"
    Else
        rawMatWS.Cells.Clear
    End If
    
    ' 复制标题行
    rawMatWS.Range("A1").Resize(1, 17).value = dataWS.Range("A1").Resize(1, 17).value
    
    Dim rawMatRow As Long
    rawMatRow = 2
    
    ' 筛选原材料、辅材、备品备件的数据
    For i = 2 To lastRow
        Dim materialType As String
        materialType = CStr(dataWS.Cells(i, 2).value)
        
        If InStr(materialType, "原材料") > 0 Or InStr(materialType, "辅材") > 0 Or InStr(materialType, "备品备件") > 0 Then
            rawMatWS.Range("A" & rawMatRow & ":Q" & rawMatRow).value = _
                dataWS.Range("A" & i & ":Q" & i).value
            rawMatRow = rawMatRow + 1
        End If
    Next i
    
    ' 创建成品表
    Dim finishedProdWS As Worksheet
    On Error Resume Next
    Set finishedProdWS = wb.Sheets("成品")
    On Error GoTo 0
    
    If finishedProdWS Is Nothing Then
        Set finishedProdWS = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        finishedProdWS.Name = "成品"
    Else
        finishedProdWS.Cells.Clear
    End If
    
    ' 复制标题行
    finishedProdWS.Range("A1").Resize(1, 17).value = dataWS.Range("A1").Resize(1, 17).value
    
    Dim finishedProdRow As Long
    finishedProdRow = 2
    
    ' 筛选电化铝、纸品、膜品的数据
    For i = 2 To lastRow
        Dim groupText As String
        Dim groupType As String
        groupText = CStr(dataWS.Cells(i, 6).value)
        groupType = CStr(dataWS.Cells(i, 2).value)
        
        If (InStr(groupText, "电化铝") > 0 Or InStr(groupText, "纸品") > 0 Or InStr(groupText, "膜品") > 0) And InStr(groupType, "成品") > 0 Then
            finishedProdWS.Range("A" & finishedProdRow & ":Q" & finishedProdRow).value = _
                dataWS.Range("A" & i & ":Q" & i).value
            finishedProdRow = finishedProdRow + 1
        End If
    Next i
    
    MsgBox "数据清洗完成！" & vbCrLf & _
           "处理的工作簿: " & wb.Name & vbCrLf & _
           "处理的工作表: " & dataWS.Name & vbCrLf & _
           "数据行数: " & lastRow, vbInformation, "完成"
    
End Sub

