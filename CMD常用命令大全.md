# Windows CMD 常用命令大全

> 本文档整理自 GitHub、CSDN、Microsoft Learn 等公开资料，按使用场景分类汇总 Windows CMD 命令，便于查阅。

## 目录

- [一、目录与文件管理](#一目录与文件管理)
- [二、批量操作（重点）](#二批量操作重点)
- [三、文件内容查看与搜索](#三文件内容查看与搜索)
- [四、系统信息与管理](#四系统信息与管理)
- [五、进程与服务管理](#五进程与服务管理)
- [六、网络命令](#六网络命令)
- [七、环境变量](#七环境变量)
- [八、重定向与管道](#八重定向与管道)
- [九、批处理脚本语法](#九批处理脚本语法)
- [十、常用实用技巧](#十常用实用技巧)
- [参考资料](#参考资料)

---

## 一、目录与文件管理

### 1. 导航类

| 命令 | 说明 |
|---|---|
| `cd` | 显示当前目录 |
| `cd <dir>` | 进入指定子目录 |
| `cd ..` | 返回上一级目录 |
| `cd \` | 跳转到当前盘根目录 |
| `cd /d D:\path` | 同时切换盘符和目录 |
| `D:` | 切换到 D 盘 |
| `pushd <dir>` | 保存当前目录并跳转 |
| `popd` | 返回到 pushd 保存的目录 |

### 2. 列出与查看

| 命令 | 说明 |
|---|---|
| `dir` | 列出当前目录文件和文件夹 |
| `dir /a` | 包含隐藏文件 |
| `dir /s` | 递归列出子目录内容 |
| `dir /b` | 仅显示文件名 |
| `dir /o:d` | 按日期排序 |
| `dir /o:s` | 按大小排序 |
| `dir /w` | 宽列表显示 |
| `dir /p` | 分页查看 |
| `dir *.txt` | 按扩展名过滤 |
| `type <file>` | 打印文件内容 |
| `more <file>` | 分页查看文件 |
| `cls` | 清屏 |

### 3. 文件与文件夹操作

| 命令 | 说明 |
|---|---|
| `md <dir>` / `mkdir <dir>` | 新建文件夹 |
| `rd <dir>` / `rmdir <dir>` | 删除空文件夹 |
| `rmdir /s /q <dir>` | 递归且静默删除文件夹 |
| `copy <src> <dst>` | 复制文件 |
| `copy /y <src> <dst>` | 复制不提示覆盖 |
| `xcopy <src> <dst> /e /i` | 复制目录树（含空目录） |
| `xcopy /s /e /y <src> <dst>` | 递归复制、不提示覆盖 |
| `robocopy <src> <dst> /e` | 高级复制（含子目录） |
| `robocopy <src> <dst> /mir` | 镜像同步目录 |
| `move <src> <dst>` | 移动或重命名 |
| `ren <old> <new>` | 重命名 |
| `del <file>` | 删除文件 |
| `del /f /s /q <file>` | 强制递归静默删除 |
| `attrib +h <file>` | 设置隐藏属性 |
| `attrib -h <file>` | 取消隐藏属性 |

---

## 二、批量操作（重点）

### 1. 批量创建文件夹（`md` 命令）

#### a) 一次创建多个文件夹

```cmd
md folder1 folder2 folder3
```

也可使用逗号、分号、等号分隔：

```cmd
md test1,test2,test3
md test1;test2;test3
md test1=test2=test3
```

含空格的文件夹名需加引号：

```cmd
md "My Folder 1" "My Folder 2"
```

#### b) 创建多级目录

```cmd
md 111\222\333
```

#### c) 同时创建多个目录层

```cmd
md test1\111\222\333,test2\111\222\333
```

#### d) 用 for 循环批量创建序号文件夹

创建 `folder1` 到 `folder10`：

```cmd
for /l %i in (1,1,10) do md folder%i
```

> 说明：`/l` 表示数值循环，`(1,1,10)` 表示起始值 1、步长 1、终值 10。在 `.bat` 批处理文件中需用 `%%i` 而非 `%i`。

#### e) 带前导零命名（如 `folder01` ~ `folder10`）

```cmd
for /l %i in (1,1,10) do (
  if %i lss 10 (
    md "folder0%i"
  ) else (
    md "folder%i"
  )
)
```

#### f) 从 names.txt 读取名称批量创建

```cmd
for /F "usebackq delims=" %i in ("names.txt") do md "%i"
```

#### g) 用 PowerShell 嵌入日期/序号

```powershell
1..10 | ForEach-Object {
  $date = Get-Date -Format "yyyy-MM-dd"
  New-Item -ItemType Directory -Name "员工_$_$date"
}
```

### 2. 一键式批处理脚本示例

为每位员工创建主文件夹及"文档/合同/考勤"子目录：

```cmd
@echo off
setlocal enabledelayedexpansion
for /f "delims=" %%a in (employees.txt) do (
  md "%%a"
  md "%%a\文档"
  md "%%a\合同"
  md "%%a\考勤"
)
```

> 保存为 `CreateFolders.bat`，与 `employees.txt` 同目录，双击运行即可。

### 3. 批量复制/备份

```cmd
xcopy "C:\MyProject" "D:\Backup\MyProject" /s /e /y
robocopy C:\MyFolder D:\Backup\MyFolder /mir /MT:8
```

### 4. 批量删除

```cmd
del /f /s /q C:\Logs\*.log
rmdir /s /q D:\OldDemo
```

---

## 三、文件内容查看与搜索

| 命令 | 说明 |
|---|---|
| `find "text" <file>` | 在文件中查找字符串 |
| `find /i "text" <file>` | 不区分大小写 |
| `find /c "text" <file>` | 统计匹配次数 |
| `findstr "pattern" <file>` | 正则搜索 |
| `findstr /s "text" *.*` | 递归搜索当前目录 |
| `findstr /i /s "text" *.txt` | 递归、忽略大小写搜索 txt 文件 |
| `where <cmd>` | 查找可执行文件位置 |
| `dir /s /p report.docx` | 递归搜索文件 |

### 通配符

| 通配符 | 含义 | 示例 |
|---|---|---|
| `*` | 匹配多个字符 | `dir *.txt` 列出所有 txt |
| `?` | 匹配单个字符 | `dir file?.txt` 匹配 file1.txt 不匹配 file10.txt |

---

## 四、系统信息与管理

| 命令 | 说明 |
|---|---|
| `systeminfo` | 显示 OS 与硬件详细信息 |
| `hostname` | 显示计算机名 |
| `whoami` | 显示当前用户 |
| `ver` | 显示 Windows 版本 |
| `sfc /scannow` | 扫描并修复系统文件 |
| `chkdsk /f` | 检查并修复磁盘错误 |
| `diskpart` | 磁盘分区管理 |
| `wmic` | 显示系统信息（WMI） |
| `powercfg /batteryreport` | 生成电池报告 |
| `shutdown /s /t 0` | 立即关机 |
| `shutdown /r /t 0` | 立即重启 |
| `shutdown /a` | 取消已计划的关机 |
| `shutdown /s /t 60` | 60 秒后关机 |

---

## 五、进程与服务管理

| 命令 | 说明 |
|---|---|
| `tasklist` | 列出所有运行进程 |
| `tasklist /fi "imagename eq app.exe"` | 按名称过滤进程 |
| `taskkill /im <name>.exe /f` | 按名称强制结束进程 |
| `taskkill /pid <PID> /f` | 按 PID 强制结束 |
| `taskkill /pid <PID> /f /t` | 结束进程及其子进程 |
| `start <program>` | 启动程序 |
| `sc query` | 查询服务状态 |
| `sc start <service>` | 启动服务 |
| `sc stop <service>` | 停止服务 |

---

## 六、网络命令

| 命令 | 说明 |
|---|---|
| `ipconfig` | 显示 IP 配置 |
| `ipconfig /all` | 显示完整网络信息 |
| `ipconfig /flushdns` | 清除 DNS 缓存 |
| `ping <host>` | 测试连通性 |
| `ping -t <host>` | 持续 ping |
| `tracert <host>` | 路由跟踪 |
| `nslookup <host>` | DNS 查询 |
| `netstat -an` | 显示所有连接与监听端口 |
| `netstat -ano` | 显示连接及对应 PID |
| `netstat -ano \| findstr :8080` | 查找占用 8080 端口的进程 |
| `arp -a` | 显示本地 ARP 表 |
| `netsh interface ip show config` | 显示接口配置 |

---

## 七、环境变量

| 命令/变量 | 说明 |
|---|---|
| `set` | 列出所有变量 |
| `set VAR=value` | 为当前会话设置变量 |
| `echo %VAR%` | 输出变量值 |
| `setx VAR "value"` | 永久设置用户变量 |
| `setx VAR "value" /m` | 永久设置系统变量（需管理员） |
| `%USERPROFILE%` | 用户主目录 |
| `%TEMP%` | 临时目录 |
| `%PATH%` | 可执行文件搜索路径 |
| `%COMPUTERNAME%` | 计算机名 |
| `%USERNAME%` | 当前用户名 |
| `%DATE%` | 当前日期 |
| `%TIME%` | 当前时间 |
| `%CD%` | 当前目录 |

---

## 八、重定向与管道

| 语法 | 说明 |
|---|---|
| `cmd > file` | 标准输出覆盖写入文件 |
| `cmd >> file` | 标准输出追加到文件 |
| `cmd 2> err.txt` | 错误输出到文件 |
| `cmd > out.txt 2>&1` | 标准与错误输出到同一文件 |
| `cmd1 \| cmd2` | 管道：将前命令输出作为后命令输入 |
| `cmd > nul` | 丢弃标准输出 |
| `cmd > nul 2>&1` | 丢弃所有输出 |
| `cmd && echo ok` | 前命令成功时执行后命令 |
| `cmd \|\| echo failed` | 前命令失败时执行后命令 |

---

## 九、批处理脚本语法

| 语法 | 说明 |
|---|---|
| `@echo off` | 关闭命令回显 |
| `:: comment` 或 `rem comment` | 注释 |
| `set VAR=val` | 设置变量 |
| `set /a VAR=1+1` | 算术运算 |
| `set /p VAR=请输入:` | 读取用户输入 |
| `%VAR%` | 使用变量 |
| `!VAR!` | 延迟展开变量（循环内） |
| `%1 %2 …` | 脚本参数 |
| `if "%VAR%"=="x" ...` | 字符串比较 |
| `if %N% GTR 0 ...` | 数值比较（GTR 大于） |
| `if exist <file> ...` | 判断文件是否存在 |
| `for %%i in (a b) do ...` | 遍历列表 |
| `for /l %%i in (1,1,10) do ...` | 数值循环 |
| `for /f "tokens=*" %%l in (f) do ...` | 逐行读取文件 |
| `goto :label` | 跳转到标签 |
| `call :label arg` | 调用子例程 |
| `exit /b 0` | 退出脚本并返回码 |

### 简单批处理示例

```cmd
@echo off
echo Hello, World!
pause
```

将上述内容保存为 `.bat` 文件双击即可运行。

---

## 十、常用实用技巧

| 命令 | 说明 |
|---|---|
| `explorer .` | 在资源管理器中打开当前目录 |
| `cd \| clip` | 将当前路径复制到剪贴板 |
| `title MyWindow` | 修改 CMD 窗口标题 |
| `color 0A` | 黑底绿字（0=黑，A=绿） |
| `color /?` | 查看所有颜色代码 |
| `mode con:cols=100 lines=30` | 设置窗口尺寸 |
| `prompt $P$G` | 自定义提示符（默认显示路径加 >） |
| `doskey alias=command` | 临时别名 |

### CMD 使用提示

- **以管理员运行**：`Win + R` 输入 `cmd`，再按 `Ctrl + Shift + Enter`。
- **复制粘贴**：Windows 10+ 支持 `Ctrl + C` / `Ctrl + V`。
- **历史命令**：按 `F7` 弹出历史命令列表。
- **自动补全**：按 `Tab` 补全文件/目录名。
- **取消命令**：按 `Ctrl + C` 终止当前命令。

---

## 参考资料

- [Microsoft Learn — mkdir 命令](https://learn.microsoft.com/zh-cn/windows-server/administration/windows-commands/mkdir)
- [Microsoft Learn — Windows 命令](https://learn.microsoft.com/zh-cn/windows-server/administration/windows-commands/windows-commands/)
- [CSDN — windows操作系统 cmd 如何批量创建文件夹](https://blog.csdn.net/weixin_46339668/article/details/139945498)
- [CSDN — DOS cmd mkdir 命令快速新建目录层、目录](https://lishuoboy.blog.csdn.net/article/details/84580454)
- [CSDN — 常用的 Windows CMD（命令提示符）指令合集](https://blog.csdn.net/2401_83912923/article/details/145149418)
- [GitHub — CestMoiRoma/small-scrips cmd cheatsheet](https://github.com/CestMoiRoma/small-scrips/wiki/cmd-cheatsheet)
- [GitHub — leotech-projects/cheat-sheet](https://github.com/leotech-projects/cheat-sheet/blob/main/prompt_de_comando.md)
- [PHP 中文 — Windows 命令如何批量新建文件夹](https://www.php.cn/faq/2063055.html)
- [TechPCTips — CMD Commands Cheat Sheet](https://techpctips.com/cmd-command-prompt-commands-cheat-sheet/)
