'==========================================================================
' pack.vbs - 一键生成安装包（纯 VBScript，无需 Node.js）
'
' 用法: 双击运行
' 功能: 将 FinanceRPA.xlam + install.vbs + 说明打包到 FinanceRPA安装包/
' 前提: 先运行 build 生成最新的 FinanceRPA.xlam
'==========================================================================
Option Explicit

Dim fso, scriptDir, xlamPath, packDir
Dim resp, file, folder

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
xlamPath = scriptDir & "\FinanceRPA.xlam"
packDir = scriptDir & "\FinanceRPA安装包"

' 检查 xlam
If Not fso.FileExists(xlamPath) Then
    MsgBox "FinanceRPA.xlam not found!" & vbCrLf & _
           "Please run build first.", vbCritical, "Error"
    WScript.Quit(1)
End If

resp = MsgBox("Generate installation package?" & vbCrLf & vbCrLf & _
              "Source: FinanceRPA.xlam" & vbCrLf & _
              "Output: FinanceRPA安装包\" & vbCrLf & vbCrLf & _
              "Click Yes to continue.", vbYesNo + vbQuestion, "Pack")
If resp = vbNo Then WScript.Quit(0)

' 清理旧安装包
If fso.FolderExists(packDir) Then
    fso.DeleteFolder packDir, True
End If

' 创建目录
fso.CreateFolder packDir

' 复制 xlam
fso.CopyFile xlamPath, packDir & "\FinanceRPA.xlam", True

' 生成 install.vbs
Dim stream, installContent
installContent = "'==========================================================================" & vbCrLf & _
"' FinanceRPA 加载宏 — 一键安装脚本" & vbCrLf & _
"'" & vbCrLf & _
"' 用法: 双击运行" & vbCrLf & _
"' 功能: 自动注册 WPS 加载宏，配置信任设置" & vbCrLf & _
"'==========================================================================" & vbCrLf & _
"Option Explicit" & vbCrLf & _
"" & vbCrLf & _
"Dim fso, app, wshShell" & vbCrLf & _
"Dim srcDir, xlamPath, startupPath, regKey" & vbCrLf & _
"Dim resp" & vbCrLf & _
"" & vbCrLf & _
"Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
"srcDir = fso.GetParentFolderName(WScript.ScriptFullName)" & vbCrLf & _
"xlamPath = srcDir & ""\\FinanceRPA.xlam""" & vbCrLf & _
"" & vbCrLf & _
"If Not fso.FileExists(xlamPath) Then" & vbCrLf & _
"    MsgBox ""FinanceRPA.xlam not found!"" & vbCrLf & _" & vbCrLf & _
"           ""Path: "" & xlamPath, vbCritical, ""Error""" & vbCrLf & _
"    WScript.Quit(1)" & vbCrLf & _
"End If" & vbCrLf & _
"" & vbCrLf & _
"resp = MsgBox(""FinanceRPA 加载宏安装"" & vbCrLf & vbCrLf & _" & vbCrLf & _
"              ""将执行以下操作:"" & vbCrLf & _" & vbCrLf & _
"              ""  1. 配置 WPS 信任设置"" & vbCrLf & _" & vbCrLf & _
"              ""  2. 复制加载宏到启动目录"" & vbCrLf & _" & vbCrLf & _
"              ""  3. 注册加载宏"" & vbCrLf & vbCrLf & _" & vbCrLf & _
"              ""点击 Yes 开始安装"", vbYesNo + vbQuestion, ""Install"")" & vbCrLf & _
"If resp = vbNo Then WScript.Quit(0)" & vbCrLf & _
"" & vbCrLf & _
"' === Step 1: 设置信任 VBA 工程 ===" & vbCrLf & _
"Set wshShell = CreateObject(""WScript.Shell"")" & vbCrLf & _
"regKey = ""HKCU\Software\Kingsoft\Office\ET\Security\AccessVBOM""" & vbCrLf & _
"On Error Resume Next" & vbCrLf & _
"wshShell.RegWrite regKey, 1, ""REG_DWORD""" & vbCrLf & _
"If Err.Number <> 0 Then" & vbCrLf & _
"    WScript.Echo ""Warning: Cannot set registry key""" & vbCrLf & _
"    Err.Clear" & vbCrLf & _
"End If" & vbCrLf & _
"On Error GoTo 0" & vbCrLf & _
"" & vbCrLf & _
"' === Step 2: 创建 WPS COM 对象 ===" & vbCrLf & _
"On Error Resume Next" & vbCrLf & _
"Set app = CreateObject(""Ket.Application"")" & vbCrLf & _
"If app Is Nothing Then Set app = CreateObject(""ET.Application"")" & vbCrLf & _
"If app Is Nothing Then Set app = CreateObject(""Excel.Application"")" & vbCrLf & _
"On Error GoTo 0" & vbCrLf & _
"" & vbCrLf & _
"If app Is Nothing Then" & vbCrLf & _
"    MsgBox ""Cannot create WPS Application."", vbCritical, ""Error""" & vbCrLf & _
"    WScript.Quit(1)" & vbCrLf & _
"End If" & vbCrLf & _
"" & vbCrLf & _
"app.Visible = False" & vbCrLf & _
"app.DisplayAlerts = False" & vbCrLf & _
"" & vbCrLf & _
"' === Step 3: 复制到启动文件夹 ===" & vbCrLf & _
"startupPath = app.StartupPath" & vbCrLf & _
"" & vbCrLf & _
"If Len(startupPath) > 0 And fso.FolderExists(startupPath) Then" & vbCrLf & _
"    If fso.FileExists(startupPath & ""\\FinanceRPA.xlam"") Then" & vbCrLf & _
"        fso.DeleteFile startupPath & ""\\FinanceRPA.xlam"", True" & vbCrLf & _
"    End If" & vbCrLf & _
"    fso.CopyFile xlamPath, startupPath & ""\\FinanceRPA.xlam"", True" & vbCrLf & _
"End If" & vbCrLf & _
"" & vbCrLf & _
"' === Step 4: 注册加载宏 ===" & vbCrLf & _
"Dim installed, ai" & vbCrLf & _
"installed = False" & vbCrLf & _
"On Error Resume Next" & vbCrLf & _
"For Each ai In app.AddIns" & vbCrLf & _
"    If InStr(LCase(ai.Name), ""financerpa"") > 0 Then" & vbCrLf & _
"        ai.Installed = True" & vbCrLf & _
"        installed = True" & vbCrLf & _
"        Exit For" & vbCrLf & _
"    End If" & vbCrLf & _
"Next" & vbCrLf & _
"" & vbCrLf & _
"If Not installed Then" & vbCrLf & _
"    app.AddIns.Add xlamPath, True" & vbCrLf & _
"    For Each ai In app.AddIns" & vbCrLf & _
"        If InStr(LCase(ai.Name), ""financerpa"") > 0 Then" & vbCrLf & _
"            ai.Installed = True" & vbCrLf & _
"            installed = True" & vbCrLf & _
"            Exit For" & vbCrLf & _
"        End If" & vbCrLf & _
"    Next" & vbCrLf & _
"End If" & vbCrLf & _
"On Error GoTo 0" & vbCrLf & _
"" & vbCrLf & _
"app.Quit" & vbCrLf & _
"" & vbCrLf & _
"MsgBox ""Installation complete!"" & vbCrLf & vbCrLf & _" & vbCrLf & _
"       ""Please open WPS Spreadsheet to verify."" & vbCrLf & _" & vbCrLf & _
"       ""Look for 'Add-Ins' (加载项) tab in the ribbon."", vbInformation, ""Install Complete"""

Set stream = CreateObject("ADODB.Stream")
stream.Type = 2
stream.Charset = "gbk"
stream.Open
stream.WriteText installContent
stream.SaveToFile packDir & "\install.vbs", 2
stream.Close

' 生成使用说明.txt
Dim readmeContent
readmeContent = "FinanceRPA 加载宏安装说明" & vbCrLf & _
"==========================" & vbCrLf & _
"" & vbCrLf & _
"安装步骤" & vbCrLf & _
"--------" & vbCrLf & _
"1. 确认已安装 WPS 表格（2019 及以上版本）" & vbCrLf & _
"2. 关闭所有 WPS 窗口" & vbCrLf & _
"3. 双击 install.vbs" & vbCrLf & _
"4. 点击 Yes 开始安装" & vbCrLf & _
"5. 安装完成后打开 WPS 表格" & vbCrLf & _
"" & vbCrLf & _
"验证安装" & vbCrLf & _
"--------" & vbCrLf & _
"1. 打开 WPS 表格" & vbCrLf & _
"2. 查看功能区是否有「加载项」选项卡" & vbCrLf & _
"3. 点击「加载项」应看到「财务RPA」工具组" & vbCrLf & _
"4. 点击任意按钮测试功能" & vbCrLf & _
"" & vbCrLf & _
"如果看不到按钮" & vbCrLf & _
"--------------" & vbCrLf & _
"1. 右键功能区 → 自定义功能区" & vbCrLf & _
"2. 勾选「加载项」选项卡" & vbCrLf & _
"3. 确定后重新打开 WPS" & vbCrLf & _
"" & vbCrLf & _
"卸载方法" & vbCrLf & _
"--------" & vbCrLf & _
"1. 关闭 WPS" & vbCrLf & _
"2. 删除启动文件夹中的 FinanceRPA.xlam" & vbCrLf & _
"3. 重新打开 WPS，选项卡消失" & vbCrLf & _
"" & vbCrLf & _
"系统要求" & vbCrLf & _
"--------" & vbCrLf & _
"- WPS 表格 2019 及以上" & vbCrLf & _
"- Windows 10/11" & vbCrLf & _
"- 无需安装 Node.js 或其他依赖"

Set stream = CreateObject("ADODB.Stream")
stream.Type = 2
stream.Charset = "utf-8"
stream.Open
stream.WriteText readmeContent
stream.SaveToFile packDir & "\使用说明.txt", 2
stream.Close

' 完成
Dim fileList
fileList = ""
For Each file In fso.GetFolder(packDir).Files
    fileList = fileList & "  " & file.Name & " (" & file.Size & " bytes)" & vbCrLf
Next

MsgBox "Package generated!" & vbCrLf & vbCrLf & _
       "Location: " & packDir & vbCrLf & vbCrLf & _
       "Files:" & vbCrLf & fileList & vbCrLf & _
       "Copy the entire folder to another computer," & vbCrLf & _
       "then double-click install.vbs", vbInformation, "Done"
