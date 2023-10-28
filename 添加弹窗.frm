VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "添加弹窗"
   ClientHeight    =   1680
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   8052
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   8052
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "删除文件"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "确定"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "检测进程"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "检查路径"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "取消"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6960
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "浏览"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   0
      Width           =   4815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "弹窗进程名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "弹窗文件路径："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FileN As String
Private Function TestFileName(FileN As String)
If FileN = "explorer.exe" Then
    Call MsgBox("不支持拦截Windows资源管理器！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "csrss.exe" Then
    Call MsgBox("不支持拦截csrss.exe！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "Client Server Runtime Process.exe" Then
    Call MsgBox("不支持拦截Client Server Runtime Process.exe！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "smss.exe" Then
    Call MsgBox("不支持拦截smss.exe！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "system.exe" Then
    Call MsgBox("不支持拦截system.exe！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "cmd.exe" Then
    Call MsgBox("不支持拦截Windows命令提示符！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "lsass.exe" Then
    Call MsgBox("不支持拦截lsass.exe！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "winlogon.exe" Then
    Call MsgBox("不支持拦截Windows用户登录程序！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "spoolsv.exe" Then
    Call MsgBox("不支持拦截spoolsv.exe！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "spoolst.exe" Then
    Call MsgBox("不支持拦截spoolst.exe！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "svchost.exe" Then
    Call MsgBox("不支持拦截svchost.exe！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "ljxpopupkiller.exe" Then
    Call MsgBox("不支持拦截LJXPopupKiller.exe！", vbOKOnly + vbCritical, "不支持拦截的程序")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If



End Function
Private Sub Command1_Click()
On Error GoTo Errt
Dim OpenUrl As String
CD1.Filter = "可执行文件(*.exe)|*.EXE|动态链接库(*.dll)|*.DLL|所有文件(*.*)|*.*"
CD1.FilterIndex = 1
CD1.ShowOpen
If CD1.FileName = "" Then
    Exit Sub
End If
OpenUrl = CD1.FileName
Text1.Text = OpenUrl
'检测文件名
Dim t As Long
Dim X As Long
Dim tmp As String
t = Len(OpenUrl)
For X = 1 To t
    tmp = Mid(OpenUrl, (t - X), 1)
    If tmp = "\" Then
        Exit For
    End If
Next
Dim FileN As String
FileN = Mid(OpenUrl, (t - X + 1), X)
Text2.Text = FileN
'判断文件是否符合要求
TestFileName (LCase(FileN))
Exit Sub
Errt:
Call MsgBox("文件浏览错误：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "错误")

End Sub

Private Sub Command2_Click()
If AddPopup = True Then
    End
End If
Unload Me
End Sub

Private Sub Command3_Click()
Dim OpenUrl As String
Dim t

OpenUrl = Text1.Text
t = Len(OpenUrl)
For X = 1 To t
    tmp = Mid(OpenUrl, (t - X), 1)
    If tmp = "\" Then
        Exit For
    End If
Next
If Dir(OpenUrl) = "" Then
    Call MsgBox("路径错误：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "错误")
    Exit Sub
End If
t = Mid(OpenUrl, (t - X + 1), (t - X))
TestFileName (t)
End Sub

Private Sub Command4_Click()
On Error GoTo Errt
Dim a As Boolean
Dim Stri As String
Dim q

Stri = Text2.Text
a = CheckExeIsRun(Stri)

If a = True Then
    Call MsgBox("进程" & Stri & "已经被启动。" & vbCrLf & "检测到的进程数：" & GetExeNumber(Stri), vbOKOnly + vbInformation, "进程" & Stri & "已启动")
Else
    Call MsgBox("没有找到指定的进程：" & Stri, vbOKOnly + vbCritical, "找不到进程" & Stri)
End If
Exit Sub
Errt:
Call MsgBox("检测进程时遇到错误：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "错误")
End Sub

Private Sub Command5_Click()
'On Error GoTo Errt
Dim OpenUrl As String
Dim t As String

OpenUrl = Text1.Text
t = Len(OpenUrl)
For X = 1 To t
    tmp = Mid(OpenUrl, (t - X + 1), 1)
    If tmp = "\" Then
        Exit For
    End If
Next

t = Mid(OpenUrl, (t - X + 1), (t - X))
TestFileName (t)
TestFileName (FileN)
Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & Module1.MaxPops & ".ltx") For Output As #1
Print #1, Text1.Text
Module1.MaxPops = Module1.MaxPops + 1
Close #1

If AddPopup = True Then
    End
End If

Call LoadAll
Unload Form2
On Error Resume Next
Form2.Show
Unload Me
Exit Sub
Errt:
Close #1
Call MsgBox("错误：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "   ")
Unload Me
End Sub

Private Sub Command6_Click()
On Error GoTo Errt
Dim t
t = MsgBox("所选文件将被永久删除，确定要这样做吗？", vbOKCancel + vbExclamation, "删除确定")
If t = vbOK Then
    Kill (Text1.Text)
    If Dir(Text1.Text) = "" Then
        MsgBox ("所选文件已经成功删除！")
    Else
        Call MsgBox("文件删除失败！", vbOKOnly + vbCritical, "文件删除失败")
    End If
End If
Exit Sub
Errt:
Call MsgBox("文件删除错误：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Form_Load()
Form2.Enabled = False
Command5.Enabled = False

If HavAdminPer = True Then
    Me.Caption = " [Admin] " & Me.Caption
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Enabled = True
End Sub

Private Sub Text1_Change()
On Error Resume Next
OpenUrl = Text1.Text
t = Len(OpenUrl)
For X = 1 To t
    tmp = Mid(OpenUrl, (t - X), 1)
    If tmp = "\" Then
        Exit For
    End If
Next
Dim FileN As String
FileN = Mid(OpenUrl, (t - X + 1), X)
Text2.Text = FileN
If Text1.Text = "" Then
    Command5.Enabled = False
Else
    Command5.Enabled = True
End If
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then
    Command5.Enabled = False
Else
    Command5.Enabled = True
End If
End Sub
