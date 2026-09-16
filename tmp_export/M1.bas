Attribute VB_Name = "RibbonCallbacks"
'==========================================================================
' Ribbon 选项卡回调模块
'
' 维护方式: 直接编辑本文件，运行 build.vbs 即可更新到 .xlsm
' 新增按钮:
'   1. 在 customUI.xml 中添加 <button> 标签
'   2. 在此模块添加对应的 Sub，名称与 onAction 一致
'   3. 运行 build.vbs 更新到 .xlsm
'==========================================================================

Public g_Ribbon As IRibbonUI

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set g_Ribbon = ribbon
End Sub

'--- 库存价值表数据清洗 ---
Public Sub OnBtnInvValue(control As IRibbonControl)
    Call 库存价值表数据清洗
End Sub

'--- 库龄表数据清洗 ---
Public Sub OnBtnInvAging(control As IRibbonControl)
    Call 清洗库龄表数据
End Sub

'--- 库龄表期初变化分析 ---
Public Sub OnBtnAgingAnalysis(control As IRibbonControl)
    Call 库龄表期初变化分析
End Sub

'--- 采购发票信息数据清洗 ---
Public Sub OnBtnInvoiceClean(control As IRibbonControl)
    Call 采购发票信息数据清洗
End Sub

'--- 关于 ---
Public Sub OnBtnAbout(control As IRibbonControl)
    MsgBox "财务RPA工作台 v1.0" & vbCrLf & _
           "更新日期: 2026-08-20" & vbCrLf & _
           "WPS VBA 7.1" & vbCrLf & vbCrLf & _
           "模块清单:" & vbCrLf & _
           "  - 库存价值表数据清洗" & vbCrLf & _
           "  - 清洗库龄表数据" & vbCrLf & _
           "  - 库龄表期初变化分析" & vbCrLf & _
           "  - 采购发票信息数据清洗", vbInformation, "关于"
End Sub
