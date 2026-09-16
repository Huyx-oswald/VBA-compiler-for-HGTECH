Attribute VB_Name = "Module5"
Sub 采购发票信息数据清洗()
    '=============================================================
    ' 功能说明：
    ' 1. 筛选 H列="H" 且 W列="否" 的行，提取AK列非空值作为特征集合
    ' 2. 删除所有 AK列值在该特征集合中的行
    ' 3. 对剩余数据中 W列="是" 的行，将 Q列数值改为负数
    '
    ' 列对应关系（基于实际文件）：
    '   H列 = D/C       W列 = 退货项目    Q列 = 金额    AK列 = 参考
    '=============================================================
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "没有数据行，请检查！", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '---------------------------------------------------------
    ' 第一步：收集特征集合（H="H" 且 W="否" 时 AK列的非空值）
    '---------------------------------------------------------
    Dim featureDict As Object
    Set featureDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim akVal As String
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, "H").value)) = "H" And _
           Trim(ws.Cells(i, "W").value) = "否" Then
            akVal = Trim(CStr(ws.Cells(i, "AK").value))
            If akVal <> "" And Not featureDict.Exists(akVal) Then
                featureDict.Add akVal, True
            End If
        End If
    Next i
    
    ' 第一步结果记录（不中断后续步骤）
    Dim step1Count As Long
    step1Count = featureDict.Count
    
    '---------------------------------------------------------
    ' 第二步：从下往上删除 AK列值在特征集合中的行
    ' （从下往上删避免行号偏移问题）
    '---------------------------------------------------------
    Dim deleteCount As Long
    deleteCount = 0
    
    For i = lastRow To 2 Step -1
        akVal = Trim(CStr(ws.Cells(i, "AK").value))
        If akVal <> "" And featureDict.Exists(akVal) Then
            ws.Rows(i).Delete
            deleteCount = deleteCount + 1
        End If
    Next i
    
    '---------------------------------------------------------
    ' 第三步：W列="是" 的行，Q列数值取反（改为负数）
    '---------------------------------------------------------
    Dim negCount As Long
    negCount = 0
    
    ' 重新获取删除后的最后一行
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(ws.Cells(i, "W").value) = "是" Then
            If IsNumeric(ws.Cells(i, "Q").value) And ws.Cells(i, "Q").value <> "" Then
                ' 如果已经是负数则跳过，只对正数或零取反
                If CDbl(ws.Cells(i, "Q").value) >= 0 Then
                    ws.Cells(i, "Q").value = -CDbl(ws.Cells(i, "Q").value)
                    negCount = negCount + 1
                End If
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "处理完成！" & vbCrLf & _
           "Step1 收集特征数：" & step1Count & vbCrLf & _
           "Step2 删除行数：" & deleteCount & vbCrLf & _
           "Step3 Q列取反行数：" & negCount, _
           vbInformation, "执行结果"
           
    Set featureDict = Nothing
End Sub

