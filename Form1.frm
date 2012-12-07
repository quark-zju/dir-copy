VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirCopy"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4830
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   240
   End
   Begin VB.Timer tmrEnd 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3840
      Top             =   240
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WAIT_INFINITE = -1&
Private Const WAIT_TIMEOUT = &H102&

Private Const SYNCHRONIZE = &H100000
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
  
Dim destPath As String
Dim fso As New FileSystemObject
Dim src As Folder
Dim showingReadme As Boolean

Private Sub LogAppend(ByRef msg As String)
    Dim s As String
    s = Me.txtLog.Text
    s = s & msg
    Me.txtLog.Text = s
    Me.txtLog.SelStart = Len(s)
    Me.txtLog.Refresh
    If Not Me.Visible Then Me.Show
    DoEvents
End Sub

Private Sub Log(ByRef msg As String)
    If Len(Me.txtLog.Text) = 0 Then LogAppend (msg) Else LogAppend (vbCrLf & msg)
End Sub

Sub mkdir_p(ByRef path As String)
    If fso.FolderExists(path) Then Exit Sub
    
    mkdir_p fso.GetParentFolderName(path)
    
    Log "MKDIR " & path & " ..."
    fso.CreateFolder path
    LogAppend " ok"
End Sub

Sub ShowReadme()
    If App.PrevInstance Then End
    
    With txtLog
        .Text = _
            "DirCopy 说明" & vbCrLf & _
            "------------------" & vbCrLf & _
            "复制程序自身所在的目录下的所有子目录到目标目录，覆盖所有存在的文件。之后执行自身所在目录下的所有批处理 (bat,vbs) 文件。" & vbCrLf & _
            "" & vbCrLf & _
            "目标目录 = (命令行参数中包含 ':') ? (命令行参数) : (最后一个非移动本地磁盘的根目录 + 命令行参数)" & vbCrLf & _
            "目标目录不存在时会被自动建立。" & vbCrLf & _
            "" & vbCrLf & _
            "批处理文件执行时可以有两个环境变量可用:" & vbCrLf & _
            "%SRC%  程序自身所在目录" & vbCrLf & _
            "%DEST% 实际的目标目录" & vbCrLf & _
            "" & vbCrLf & _
            "此说明在未检测到有目录需要复制时显示。" & vbCrLf & _
            "按空格或回车或 ESC 键关闭。"
    
        Dim heightDelta As Long
        heightDelta = Me.Height - .Height
        .Height = .Height * 2.5
        Me.Height = .Height + heightDelta
    End With
    
    showingReadme = True
    Me.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Not showingReadme Then Exit Sub
    If KeyAscii = 27 Or KeyAscii = 13 Or KeyAscii = 10 Or KeyAscii = 32 Then End
End Sub

Sub Main()
    On Error GoTo ErrHandle
    
    ' get dest path
    If InStr(Command, ":") > 0 Then
        ' Use command as full dest path
        destPath = Command
    Else
        Dim drv As Drive
        Dim destDrive As String
        
        For Each drv In fso.Drives
            If drv.DriveType = Fixed And drv.DriveLetter > destDrive Then destDrive = drv.DriveLetter
        Next
        
        destPath = destDrive & ":\" & Command
    End If
    
    If Right(destPath, 1) <> "\" Then destPath = destPath & "\"
    
    mkdir_p destPath
    Log "DEST = " & destPath
    
    ' copy dirs
    Dim cur As Folder
    
    For Each cur In src.SubFolders
        Log "COPY " & cur.Name & " ..."
        cur.Copy destPath, True
        LogAppend " ok"
    Next
    
    ' set env
    SetEnvironmentVariable "DEST", Left(destPath, Len(destPath) - 1)
    SetEnvironmentVariable "SRC", src.ShortPath
    
    ' execute batchs
    ChDir src.path
    Dim bat As File
    For Each bat In src.Files
        Dim ext As String, cmd As String
        ext = UCase(Right(bat.Name, 4))
        
        Select Case ext
        Case ".BAT"
            cmd = bat.ShortPath
        Case ".VBS"
            cmd = "cscript """ & bat.ShortPath & """"
        Case Else
            cmd = ""
        End Select
        
        If cmd <> "" Then
            Log "EXEC " & bat.Name & " .."
            Dim taskId As Long
            
            taskId = Shell(cmd, vbMinimizedNoFocus)
            
            ' wait
            Dim hProcess As Long
            hProcess = OpenProcess(SYNCHRONIZE, True, taskId)
            If hProcess <> 0 Then
                Dim stat As Long
                Dim tick As Long
                tick = 0
                
                Do
                    stat = WaitForSingleObject(hProcess, 100)
                    tick = tick + 1
                    If (tick Mod 24) = 0 Then LogAppend (".")
                    DoEvents
                    If (stat <> WAIT_TIMEOUT) Then Exit Do
                Loop
                CloseHandle hProcess
            End If
            LogAppend " ok"
        End If
    Next
    
    Set src = Nothing
    
    Log "DONE."
    GoTo Out
    
ErrHandle:
    Log "Error: " & Err.Description
Out:
    tmrEnd.Enabled = True
End Sub

Private Sub Form_Load()
    Set src = fso.GetFolder(App.path)
    
    If src.SubFolders.Count = 0 Then
        ShowReadme
        Exit Sub
    Else
        showingReadme = False
        tmrMain.Enabled = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not showingReadme Then Cancel = 1
End Sub

Private Sub Form_Terminate()
    Set src = Nothing
End Sub

Private Sub tmrEnd_Timer()
    End
End Sub

Private Sub tmrMain_Timer()
    Main
    tmrMain.Enabled = False
End Sub
