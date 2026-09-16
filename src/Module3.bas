Attribute VB_Name = "Module3"
Sub 清洗库龄表数据()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim stage As String

    Set ws = ActiveSheet

    On Error GoTo ErrorHandler

    ' 检查是否已处理过
    If ws.Cells(1, 2).value = "分类" Or ws.Cells(1, 14).value = "呆滞数量" Then
        MsgBox "该工作表已处理过，请使用原始库龄表数据！", vbExclamation, "提示"
        Exit Sub
    End If

    ' 验证列结构：检查关键列标题
    Dim colCount As Long
    colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If colCount < 30 Then
        MsgBox "列数不足！当前 " & colCount & " 列，预期至少 30 列。" & vbCrLf & _
               "请确认是否为标准库龄表数据。", vbCritical, "列结构异常"
        Exit Sub
    End If

    ' 获取最后一行数据
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "未检测到数据行！", vbExclamation, "提示"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    stage = "插入分类列"
    ' 1. 在B列前插入一列用于分类
    ws.Columns("B:B").Insert Shift:=xlToRight

    ' 2. 更新分段标题（插入分类列后，原M列变成N列，依次右移）
    stage = "更新分段标题1"
    ws.Cells(1, 14).value = "0-60数量"
    ws.Cells(1, 15).value = "0-60金额"
    ws.Cells(1, 16).value = "60-90数量"
    ws.Cells(1, 17).value = "60-90金额"
    ws.Cells(1, 18).value = "90-120数量"
    ws.Cells(1, 19).value = "90-120金额"
    ws.Cells(1, 20).value = "120-150数量"
    ws.Cells(1, 21).value = "120-150金额"
    ws.Cells(1, 22).value = "150-180数量"
    ws.Cells(1, 23).value = "150-180金额"
    ws.Cells(1, 24).value = "180-360数量"
    ws.Cells(1, 25).value = "180-360金额"
    ws.Cells(1, 26).value = "360-720数量"
    ws.Cells(1, 27).value = "360-720金额"
    ws.Cells(1, 28).value = "720-1080数量"
    ws.Cells(1, 29).value = "720-1080金额"
    ws.Cells(1, 30).value = "1080以上数量"
    ws.Cells(1, 31).value = "1080以上金额"

    ' 根据物料号首字母分类（原B列现在是C列）
    stage = "填充分类"
    For i = 2 To lastRow
        Dim firstChar As String
        firstChar = Left(CStr(ws.Cells(i, 3).value), 1)

        Select Case firstChar
            Case "1"
                ws.Cells(i, 2).value = "原材料"
            Case "2"
                ws.Cells(i, 2).value = "半成品"
            Case "3"
                ws.Cells(i, 2).value = "成品"
            Case "4"
                ws.Cells(i, 2).value = "辅材"
            Case "5"
                ws.Cells(i, 2).value = "备品备件"
            Case "9"
                ws.Cells(i, 2).value = "单元板"
            Case Else
                ws.Cells(i, 2).value = "其他"
        End Select
    Next i

    ws.Cells(1, 2).value = "分类"

    ' 3. 在N列前插入两列：呆滞数量和呆滞金额
    stage = "插入呆滞列"
    ws.Columns("N:N").Insert Shift:=xlToRight
    ws.Columns("N:N").Insert Shift:=xlToRight

    ws.Cells(1, 14).value = "呆滞数量"
    ws.Cells(1, 15).value = "呆滞金额"

    ' 4. 重新设置分段列标题（又插入两列后右移2位）
    stage = "更新分段标题2"
    ws.Cells(1, 16).value = "0-60数量"
    ws.Cells(1, 17).value = "0-60金额"
    ws.Cells(1, 18).value = "60-90数量"
    ws.Cells(1, 19).value = "60-90金额"
    ws.Cells(1, 20).value = "90-120数量"
    ws.Cells(1, 21).value = "90-120金额"
    ws.Cells(1, 22).value = "120-150数量"
    ws.Cells(1, 23).value = "120-150金额"
    ws.Cells(1, 24).value = "150-180数量"
    ws.Cells(1, 25).value = "150-180金额"
    ws.Cells(1, 26).value = "180-360数量"
    ws.Cells(1, 27).value = "180-360金额"
    ws.Cells(1, 28).value = "360-720数量"
    ws.Cells(1, 29).value = "360-720金额"
    ws.Cells(1, 30).value = "720-1080数量"
    ws.Cells(1, 31).value = "720-1080金额"
    ws.Cells(1, 32).value = "1080以上数量"
    ws.Cells(1, 33).value = "1080以上金额"

    ' 5. 判定呆滞物料并填充数值
    stage = "判定呆滞物料"
    For i = 2 To lastRow
        Dim sGroupDesc As String
        Dim sCategory As String
        Dim seg6Qty As Double, seg6Amt As Double
        Dim seg7Qty As Double, seg7Amt As Double
        Dim seg8Qty As Double, seg8Amt As Double
        Dim seg9Qty As Double, seg9Amt As Double
        Dim totalStagnantQty As Double, totalStagnantAmt As Double
        Dim isStagnant As Boolean

        sGroupDesc = CStr(ws.Cells(i, 9).value)
        sCategory = CStr(ws.Cells(i, 2).value)

        isStagnant = False
        seg6Qty = 0: seg6Amt = 0
        seg7Qty = 0: seg7Amt = 0
        seg8Qty = 0: seg8Amt = 0
        seg9Qty = 0: seg9Amt = 0

        If sCategory = "辅材" Or sCategory = "备品备件" Then
            isStagnant = False
        ElseIf InStr(sGroupDesc, "半成品-镍版") > 0 Then
            seg9Qty = SafeCDbl(ws.Cells(i, 32).value)
            seg9Amt = SafeCDbl(ws.Cells(i, 33).value)
            If seg9Qty > 0 Or seg9Amt > 0 Then
                isStagnant = True
            End If
        Else
            seg6Qty = SafeCDbl(ws.Cells(i, 26).value)
            seg6Amt = SafeCDbl(ws.Cells(i, 27).value)
            seg7Qty = SafeCDbl(ws.Cells(i, 28).value)
            seg7Amt = SafeCDbl(ws.Cells(i, 29).value)
            seg8Qty = SafeCDbl(ws.Cells(i, 30).value)
            seg8Amt = SafeCDbl(ws.Cells(i, 31).value)
            seg9Qty = SafeCDbl(ws.Cells(i, 32).value)
            seg9Amt = SafeCDbl(ws.Cells(i, 33).value)

            If (seg6Qty + seg7Qty + seg8Qty + seg9Qty) > 0 Or (seg6Amt + seg7Amt + seg8Amt + seg9Amt) > 0 Then
                isStagnant = True
            End If
        End If

        If isStagnant Then
            If InStr(sGroupDesc, "半成品-镍版") > 0 Then
                totalStagnantQty = seg9Qty
                totalStagnantAmt = seg9Amt
            Else
                totalStagnantQty = seg6Qty + seg7Qty + seg8Qty + seg9Qty
                totalStagnantAmt = seg6Amt + seg7Amt + seg8Amt + seg9Amt
            End If
            ws.Cells(i, 14).value = totalStagnantQty
            ws.Cells(i, 15).value = totalStagnantAmt
        Else
            ws.Cells(i, 14).value = 0
            ws.Cells(i, 15).value = 0
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "库龄表数据清洗完成！共处理 " & (lastRow - 1) & " 行数据。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清洗失败！阶段: " & stage & vbCrLf & _
           "错误行: " & i & vbCrLf & _
           "错误: " & Err.Description & " (代码 " & Err.Number & ")", _
           vbCritical, "清洗失败"
End Sub

' 安全转换为Double，处理空值、错误值、文本
Function SafeCDbl(value As Variant) As Double
    If IsError(value) Then
        SafeCDbl = 0
    ElseIf IsEmpty(value) Then
        SafeCDbl = 0
    ElseIf IsNumeric(value) Then
        SafeCDbl = CDbl(value)
    Else
        SafeCDbl = 0
    End If
End Function

' 兼容旧代码的Nz函数（已修复IsError顺序）
Function Nz(value As Variant, Optional defaultValue As Variant = 0) As Variant
    If IsError(value) Then
        Nz = defaultValue
    ElseIf IsEmpty(value) Then
        Nz = defaultValue
    ElseIf value = "" Then
        Nz = defaultValue
    Else
        Nz = value
    End If
End Function
