Attribute VB_Name = "Module4"
Sub 库龄表期初变化分析()
    Dim wb2025 As Workbook
    Dim wb2026 As Workbook
    Dim ws2025 As Worksheet
    Dim ws2026 As Worksheet
    Dim wsResult As Worksheet
    Dim filePath2025 As String

    ' 设置文件路径（请根据实际情况修改）
    filePath2025 = "C:\Users\Finance005\Desktop\111华工图像同步文件\4、附表和简报库龄表\库龄表\库龄表2025.12\2025年12月库龄表.XLSX"

    On Error GoTo ErrorHandler

    ' 关闭屏幕刷新以提高性能
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 步骤1: 2026年数据从当前活动工作表获取，2025年从文件打开
    Set wb2026 = ActiveWorkbook
    Set ws2026 = ActiveSheet

    Set wb2025 = Workbooks.Open(filePath2025)
    Set ws2025 = wb2025.Sheets(1)
    
    ' 步骤2: 在目标库龄表工作簿中创建结果工作表
    Application.DisplayAlerts = False
    On Error Resume Next
    wb2026.Sheets("分析结果").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    ' 在目标工作簿中创建新工作表
    Set wsResult = wb2026.Sheets.Add(After:=wb2026.Sheets(wb2026.Sheets.Count))
    wsResult.Name = "分析结果"
    
    ' 步骤3: 设置标题行
    Call 设置标题行(wsResult)
    
    ' 步骤4: 按物料号汇总2026年5月数据
    Dim dict2026 As Object
    Set dict2026 = CreateObject("Scripting.Dictionary")
    Call 汇总数据(ws2026, dict2026)
    
    ' 步骤5: 按物料号汇总2025年12月数据
    Dim dict2025 As Object
    Set dict2025 = CreateObject("Scripting.Dictionary")
    Call 汇总数据(ws2025, dict2025)
    
    ' 步骤6: 填充结果表
    Call 填充结果表(wsResult, dict2026, dict2025)
    
    ' 步骤7: 按差异金额倒序排序
    Call 按差异金额排序(wsResult)
    
    ' 步骤8: 格式化结果表
    Call 格式化结果表(wsResult)
    
    ' 激活结果工作表
    wsResult.Activate
    
    ' 关闭2025年工作簿（不保存）
    wb2025.Close SaveChanges:=False
    
    ' 保存2026年工作簿（包含新生成的分析结果表）
    wb2026.Save
    
    ' 恢复屏幕刷新
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "分析完成！结果已生成在当前工作簿的【分析结果】工作表中。", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    MsgBox "发生错误: " & Err.Description, vbCritical, "错误"
End Sub

' 设置标题行
Private Sub 设置标题行(ws As Worksheet)
    With ws.Range("A1")
        .value = "物料号"
        .Offset(0, 1).value = "分类"
        .Offset(0, 2).value = "物料描述"
        .Offset(0, 3).value = "物料组描述"
        .Offset(0, 4).value = "单位"
        .Offset(0, 5).value = "当期数量"
        .Offset(0, 6).value = "当期金额"
        .Offset(0, 7).value = "当期呆滞数量"
        .Offset(0, 8).value = "当期呆滞金额"
        .Offset(0, 9).value = "期初数量"
        .Offset(0, 10).value = "期初金额"
        .Offset(0, 11).value = "期初呆滞数量"
        .Offset(0, 12).value = "期初呆滞金额"
        .Offset(0, 13).value = "数量差异"
        .Offset(0, 14).value = "金额差异"
        .Offset(0, 15).value = "数量差异幅度"
        .Offset(0, 16).value = "金额差异幅度"
    End With
End Sub

' 按物料号汇总数据
Private Sub 汇总数据(ws As Worksheet, dict As Object)
    Dim lastRow As Long
    Dim i As Long
    Dim materialCode As Variant
    Dim category As String
    Dim materialDesc As String
    Dim unit As String
    Dim qty As Double
    Dim amt As Double
    Dim groupDesc As String
    Dim stagnantQty As Double
    Dim stagnantAmt As Double
    Dim key As String
    
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    For i = 2 To lastRow
        materialCode = ws.Cells(i, "C").value
        
        ' 跳过空值
        If materialCode <> "" Then
            key = CStr(materialCode)
            category = CStr(ws.Cells(i, "B").value)
            materialDesc = CStr(ws.Cells(i, "F").value)
            unit = CStr(ws.Cells(i, "K").value)
            qty = CDbl(Nz(ws.Cells(i, "J").value, 0))
            amt = CDbl(Nz(ws.Cells(i, "L").value, 0))
            groupDesc = CStr(ws.Cells(i, "I").value)
            stagnantQty = CDbl(Nz(ws.Cells(i, "N").value, 0))
            stagnantAmt = CDbl(Nz(ws.Cells(i, "O").value, 0))
            
            If dict.Exists(key) Then
                ' 累加数量和金额
                Dim arr As Variant
                arr = dict(key)
                arr(3) = arr(3) + qty        ' 数量
                arr(4) = arr(4) + amt        ' 金额
                arr(6) = arr(6) + stagnantQty  ' 呆滞数量
                arr(7) = arr(7) + stagnantAmt  ' 呆滞金额
                dict(key) = arr
            Else
                ' 新增记录
                dict.Add key, Array(category, materialDesc, unit, qty, amt, groupDesc, stagnantQty, stagnantAmt)
            End If
        End If
    Next i
End Sub

' 填充结果表
Private Sub 填充结果表(ws As Worksheet, dict2026 As Object, dict2025 As Object)
    Dim key As Variant
    Dim rowNum As Long
    Dim arr2026 As Variant
    Dim arr2025 As Variant
    Dim qtyDiff As Double
    Dim amtDiff As Double
    Dim qtyDiffRate As Variant
    Dim amtDiffRate As Variant
    
    rowNum = 2
    
    ' 遍历2026年5月数据
    For Each key In dict2026.keys
        arr2026 = dict2026(key)
        
        ' 写入基础信息
        ws.Cells(rowNum, 1).value = key
        ws.Cells(rowNum, 2).value = arr2026(0)  ' 分类
        ws.Cells(rowNum, 3).value = arr2026(1)  ' 物料描述
        ws.Cells(rowNum, 4).value = arr2026(5)  ' 物料组描述
        ws.Cells(rowNum, 5).value = arr2026(2)  ' 单位
        ws.Cells(rowNum, 6).value = arr2026(3)  ' 当期数量
        ws.Cells(rowNum, 7).value = arr2026(4)  ' 当期金额
        ws.Cells(rowNum, 8).value = arr2026(6)  ' 当期呆滞数量
        ws.Cells(rowNum, 9).value = arr2026(7)  ' 当期呆滞金额
        
        ' 匹配2025年12月数据
        If dict2025.Exists(key) Then
            arr2025 = dict2025(key)
            ws.Cells(rowNum, 10).value = arr2025(3)  ' 期初数量
            ws.Cells(rowNum, 11).value = arr2025(4)  ' 期初金额
            ws.Cells(rowNum, 12).value = arr2025(6)  ' 期初呆滞数量
            ws.Cells(rowNum, 13).value = arr2025(7)  ' 期初呆滞金额
        Else
            ws.Cells(rowNum, 10).value = 0
            ws.Cells(rowNum, 11).value = 0
            ws.Cells(rowNum, 12).value = 0
            ws.Cells(rowNum, 13).value = 0
        End If
        
        ' 计算差异
        qtyDiff = arr2026(6) - ws.Cells(rowNum, 12).value
        amtDiff = arr2026(7) - ws.Cells(rowNum, 13).value
        
        ws.Cells(rowNum, 14).value = qtyDiff   ' 数量差异
        ws.Cells(rowNum, 15).value = amtDiff   ' 金额差异
        
        ' 计算差异幅度
        If ws.Cells(rowNum, 12).value <> 0 Then
            qtyDiffRate = qtyDiff / ws.Cells(rowNum, 12).value
        Else
            qtyDiffRate = IIf(qtyDiff = 0, 0, "#DIV/0!")
        End If
        
        If ws.Cells(rowNum, 13).value <> 0 Then
            amtDiffRate = amtDiff / ws.Cells(rowNum, 13).value
        Else
            amtDiffRate = IIf(amtDiff = 0, 0, "#DIV/0!")
        End If
        
        ws.Cells(rowNum, 16).value = qtyDiffRate  ' 数量差异幅度
        ws.Cells(rowNum, 17).value = amtDiffRate  ' 金额差异幅度
        
        rowNum = rowNum + 1
    Next key
    
    ' 处理2025年有但2026年没有的数据（期初有，期末无）
    For Each key In dict2025.keys
        If Not dict2026.Exists(key) Then
            arr2025 = dict2025(key)
            
            ws.Cells(rowNum, 1).value = key
            ws.Cells(rowNum, 2).value = arr2025(0)   ' 分类
            ws.Cells(rowNum, 3).value = arr2025(1)   ' 物料描述
            ws.Cells(rowNum, 4).value = arr2025(5)   ' 物料组描述
            ws.Cells(rowNum, 5).value = arr2025(2)   ' 单位
            ws.Cells(rowNum, 6).value = 0             ' 当期数量
            ws.Cells(rowNum, 7).value = 0             ' 当期金额
            ws.Cells(rowNum, 8).value = 0             ' 当期呆滞数量
            ws.Cells(rowNum, 9).value = 0             ' 当期呆滞金额
            ws.Cells(rowNum, 10).value = arr2025(3)  ' 期初数量
            ws.Cells(rowNum, 11).value = arr2025(4)  ' 期初金额
            ws.Cells(rowNum, 12).value = arr2025(6)  ' 期初呆滞数量
            ws.Cells(rowNum, 13).value = arr2025(7)  ' 期初呆滞金额
            ws.Cells(rowNum, 14).value = -arr2025(6) ' 数量差异
            ws.Cells(rowNum, 15).value = -arr2025(7) ' 金额差异
            ws.Cells(rowNum, 16).value = -1          ' 数量差异幅度
            ws.Cells(rowNum, 17).value = -1          ' 金额差异幅度
            
            rowNum = rowNum + 1
        End If
    Next key
End Sub

' 按差异金额倒序排序
Private Sub 按差异金额排序(ws As Worksheet)
    Dim lastRow As Long
    Dim dataRange As Range
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow > 1 Then
        Set dataRange = ws.Range("A1:Q" & lastRow)
        
        dataRange.Sort Key1:=ws.Range("O1"), Order1:=xlDescending, _
                       Header:=xlYes, MatchCase:=False, Orientation:=xlTopToBottom
    End If
End Sub

' 格式化结果表
Private Sub 格式化结果表(ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 17
    
    ' 设置标题行格式
    With ws.Range("A1:Q1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' 设置数据格式
    If lastRow > 1 Then
        ' 数量列格式（整数）
        ws.Range("F2:F" & lastRow).NumberFormat = "#,##0"   ' 当期数量
        ws.Range("H2:H" & lastRow).NumberFormat = "#,##0"   ' 当期呆滞数量
        ws.Range("J2:J" & lastRow).NumberFormat = "#,##0"   ' 期初数量
        ws.Range("L2:L" & lastRow).NumberFormat = "#,##0"   ' 期初呆滞数量
        ws.Range("N2:N" & lastRow).NumberFormat = "#,##0"   ' 数量差异
        
        ' 金额列格式（货币）
        ws.Range("G2:G" & lastRow).NumberFormat = "#,##0.00" ' 当期金额
        ws.Range("I2:I" & lastRow).NumberFormat = "#,##0.00" ' 当期呆滞金额
        ws.Range("K2:K" & lastRow).NumberFormat = "#,##0.00" ' 期初金额
        ws.Range("M2:M" & lastRow).NumberFormat = "#,##0.00" ' 期初呆滞金额
        ws.Range("O2:O" & lastRow).NumberFormat = "#,##0.00" ' 金额差异
        
        ' 差异幅度格式（百分比）
        ws.Range("P2:Q" & lastRow).NumberFormat = "0.00%"
    End If
    
    ' 自动调整列宽
    For i = 1 To lastCol
        ws.Columns(i).AutoFit
    Next i
    
    ' 冻结首行
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    ' 添加边框
    If lastRow > 1 Then
        With ws.Range("A1:Q" & lastRow).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End If
End Sub

' 辅助函数：处理空值
Private Function Nz(value As Variant, Optional defaultValue As Variant = 0) As Variant
    If IsError(value) Then
        Nz = defaultValue
    ElseIf IsEmpty(value) Or IsNull(value) Then
        Nz = defaultValue
    ElseIf value = "" Then
        Nz = defaultValue
    Else
        Nz = value
    End If
End Function

