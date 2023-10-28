VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "被拦截的弹窗"
   ClientHeight    =   7740
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   11748
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   11748
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6480
      Top             =   7320
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6672
      ItemData        =   "被拦截的弹窗.frx":0000
      Left            =   0
      List            =   "被拦截的弹窗.frx":0007
      TabIndex        =   7
      Top             =   600
      Width           =   4248
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6600
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   7335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "更改项目"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "删除项目"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "添加项目"
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "关闭窗口"
      Height          =   375
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "关闭窗口并保存更改"
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "如果移除按钮不见了就点我"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   7440
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "这里是被设置为“拦截”的弹窗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
Form3.Show
End Sub

Private Sub Command3_Click()
On Error GoTo Errt
Form2.SetFocus
Dim a
Dim b
a = List1.ListIndex
If a = -1 Then
    Call MsgBox(Languages(langNumber)(109), vbOKOnly + vbExclamation, Languages(langNumber)(110))
    Exit Sub
End If
If a <> -1 Then
    b = MsgBox(Languages(langNumber)(86) & a & Languages(langNumber)(87), vbOKCancel, Languages(langNumber)(88))
    Form2.SetFocus
    If b = vbOK Then
        Call DelFile(a)
    End If
    Call LoadAll
    Unload Form2
    Load Form2
    Call Form2.Refresh
    On Error Resume Next
    Form2.Show
End If
Exit Sub
Errt:
Call MsgBox("F2_C3_Cli：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub
Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
On Error GoTo Errt
Dim a
Dim b
Dim c
a = List1.ListIndex
If a = -1 Then
    Call MsgBox(Languages(langNumber)(109), vbOKOnly + vbExclamation, Languages(langNumber)(110))
    Exit Sub
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & a & ".ltx") = "" Then
    MsgBox (Languages(langNumber)(90))
    Exit Sub
End If
On Error Resume Next
Form5.Show
Form5.Text1.Text = Module1.Pops(List1.ListIndex)
Exit Sub
Errt:
Call MsgBox("F2_C5_Cli：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Form_Load()
On Error GoTo Errt
Form1.Hide
'Form1.Enabled = False
Text1.Text = ""
List1.Top = 600
List1.Width = 11685
Text1.Top = 600
Text1.Height = List1.Height - 10
Form1.Enabled = False
For X = 0 To 1023
    If Module1.Pops(X) <> "" Then
        If List1.List(0) = Languages(langNumber)(71) Then
            List1.List(0) = PopsP(0)
        Else
            List1.AddItem (PopsP(X))
        End If
    End If
Next

Timer1.Enabled = True
Call Form_Loadlanguage

Exit Sub
Errt:
Call MsgBox("Sta_F2_Loa：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
On Error Resume Next
Form1.Show
End Sub
Private Sub Form_Loadlanguage()
If HasOtherLanguage = True Then
    Me.Caption = Languages(langNumber)(63)
    Command1.Caption = Languages(langNumber)(64)
    Command1.ToolTipText = Languages(langNumber)(83)
    Command2.Caption = Languages(langNumber)(65)
    Command3.Caption = Languages(langNumber)(66)
    Command5.Caption = Languages(langNumber)(67)
    Label1.Caption = Languages(langNumber)(68)
    Label2.Caption = Languages(langNumber)(70)
    Label3.Caption = Languages(langNumber)(69)
    If List1.List(0) = "" Then
        List1.List(0) = Languages(langNumber)(71)
    End If
End If

If HavAdminPer = True Then
    Me.Caption = " [Admin] " & Me.Caption
End If
End Sub

Private Sub Label2_Click()
Command3.Value = True
End Sub

Private Sub Label3_Click()
Unload Form2
LoadAll
On Error Resume Next
Form2.Show
End Sub

Private Sub Timer1_Timer()
Form2.Refresh
Command3.Refresh
Command3.Visible = True
End Sub

Private Sub List1_Click()
Text1.Text = "加载中......"
Dim X
For X = 1 To 125
    If List1.Width <= 4250 Then
        List1.Width = 4250
        Exit For
    End If
    List1.Width = List1.Width - 52.4
Next
List1.Width = 4250

Dim fState
Dim fProN
If Dir(Module1.Pops(List1.ListIndex)) = "" Then
    fState = "文件已不存在"
Else
    If CheckExeIsRun(List1.List(List1.ListIndex)) = True Then
        fState = "正在运行"
        fProN = GetExeNumber(List1.List(List1.ListIndex))
    Else
        fState = "未运行"
    End If
End If


Text1.Text = ""
If List1.List(List1.ListIndex) = Languages(langNumber)(71) Then
    Text1.Text = "弹窗拦截列表是空的，现在点击右下角的“添加项目”按钮来添加一个弹窗吧。"
    Exit Sub
End If
Text1.Text = Text1.Text & "编  号：" & List1.ListIndex & vbCrLf
Text1.Text = Text1.Text & "文件名：" & List1.List(List1.ListIndex) & vbCrLf
Text1.Text = Text1.Text & "进程名：" & List1.List(List1.ListIndex) & vbCrLf
Text1.Text = Text1.Text & "现状态：" & fState & vbCrLf
If fState = "正在运行" Then
    Text1.Text = Text1.Text & "进程数：" & fProN & vbCrLf
End If
Text1.Text = Text1.Text & "路  径：" & Module1.Pops(List1.ListIndex)
End Sub

