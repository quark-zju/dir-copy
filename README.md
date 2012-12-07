DirCopy
=======

功能
----

复制 DirCopy 自身所在的目录下的所有子目录到目标目录，覆盖所有存在的文件。
之后执行自身所在目录下的所有批处理 (bat,vbs) 文件。

目标目录 = (命令行参数中包含 ':') ? (命令行参数) : (最后一个非移动本地磁盘的根目录 + 命令行参数)
目标目录不存在时会被自动建立。

批处理文件执行时可以有两个环境变量可用:
* `%SRC%`  程序自身所在目录
* `%DEST%` 实际的目标目录

故事背景
--------

* 在学校机房举办比赛需要批量对所有机器安装一些软件
* 这些软件多是绿色软件，复制到任何地方都可以运行
* 机房机器使用 Windows 系统，配置了硬件还原卡，硬盘内容更改会在重启时被还原，除了最后一个盘符
* 希望把软件装在不会被还原的地方，这样比赛过程中即便遭遇死机重启，软件也不用重装
* 机房机器可能有三个盘符（C， D， E），也可能只有两个（C，D）


脚本样例
--------

在桌面上创建快捷方式的样例 vbs 脚本:

    Set oShell   = CreateObject("WScript.Shell")
    Set oFs      = CreateObject("Scripting.FileSystemObject")
    
    sDesktopPath = oShell.SpecialFolders("AllUsersDesktop")
    sTargetDir   = oShell.ExpandEnvironmentStrings("%DEST%\FOLDER_NAME")
    sTargetPath  = sTargetDir & "\EXE_NAME.exe"
    sArguments   = ""
    sIconPath    = sTargetPath & ",0"
    sName        = "SHORTCUT_NAME"
    
    If Not oFs.FileExists(sTargetPath) Then WScript.Quit
    
    Set oShortCut = oShell.CreateShortcut(sDesktopPath & "\" & sName & ".lnk")
    oShortCut.TargetPath = sTargetPath
    oShortCut.Arguments = sArguments
    oShortCut.IconLocation = sIconPath
    oShortCut.WorkingDirectory = sTargetDir
    oShortCut.Save
