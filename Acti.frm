VERSION 5.00
Begin VB.Form Acti 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "激活 LJX弹窗杀手"
   ClientHeight    =   1368
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5568
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1368
   ScaleWidth      =   5568
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "显示密钥"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   1005
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "取消"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "确定"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "请输入LJX提供的激活密钥"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Acti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Open ("C:\Users\admin\AppData\Roaming\LJXPopupKiller\Inf\Act.prove") For Output As #1
'    Print #1, (Mid(Key, 1, 7) & Mid(Key(11, 5)))
'    Close #1
'    Call SetAttr("C:\Users\admin\AppData\Roaming\LJXPopupKiller\Inf\Act.prove", vbSystem + vbHidden + vbReadOnly)

Private Sub Command1_Click()
On Error GoTo Errt
Dim Key As String
Dim Numbers(10)
For X = 0 To 9
    Numbers(X) = Str(X)
Next
For X = 2 To Len(Text1.Text)
    If EleInArray(Str(Mid(Text1.Text, X, 1)), Numbers) = False Then
        GoTo ErrProveU
    End If
Next
Key = ""
For X = 2 To 28
    Key = Key & Str(Mid(Text1.Text, X, 1))
Next
Key = RemoveStrInStr(" ", Key)
'Key = Key & Mid(Str(Mid(Text1.Text, 2, 5)), 2, 5)
'Key = Key & Mid(Str(Mid(Text1.Text, 7, 5)), 2, 5)
'Key = Key & Mid(Str(Mid(Text1.Text, 12, 5)), 2, 5)
'Key = Key & Mid(Str(Mid(Text1.Text, 17, 5)), 2, 5)
'Key = Key & Mid(Str(Mid(Text1.Text, 22, 5)), 2, 5)
'Key = Key & Mid(Str(Mid(Text1.Text, 27, 2)), 2, 2)
Dim NowTime
NowTime = Mid(Str(Format(Now, "yyyymmddhhmmss")), 2, 14)
If Len(Key) <> 27 Then
    GoTo ErrProve
Else
    Dim tArray(0 To 13) As String
    Dim rKey As String
    rKey = ""
    For X = 0 To 13
        tArray(X) = Mid(NowTime, X + 1, 1)
    Next
    For X = 0 To 13
        If X = 0 Then
            rKey = tArray(X)
        Else
            rKey = rKey & tArray(X) + tArray(X - 1)
        End If
    Next
    Dim fx
    Dim mx
    Dim lx
    For X = 1 To 27
        If X Mod 2 = 1 Then
            fx = Mid(Key, X, 1)
            If X <> 27 Then
                lx = Mid(Key, X + 2, 1)
                mx = Val(fx) + Val(lx)
                If Len(mx) = 2 Then
                    mx = Mid(mx, 2, 1)
                End If
                If mx <> Val(Mid(Key, X + 1, 1)) Then
                    GoTo ErrProve
                End If
            End If
        End If
    Next
    If rKey <> Key Then
        Dim tKey As String
        tKey = ""
        For X = 1 To 35
            If X Mod 2 = 1 Then
                tKey = tKey & Mid(Key, X, 1)
            End If
        Next
        Dim nYear As String, kYear As String
        Dim nMonth As String, kMonth As String
        Dim nDay As String, kDay As String
        Dim nHour As String, kHour As String
        Dim nMinute As String, kMinute As String
        Dim nSecond As String, kSecond As String
        nYear = Str(Mid(NowTime, 1, 4)): kYear = Str(Mid(tKey, 1, 4))
        nMonth = Str(Mid(NowTime, 5, 2)): kMonth = Str(Mid(tKey, 5, 2))
        nDay = Str(Mid(NowTime, 7, 2)): kDay = Str(Mid(tKey, 7, 2))
        nHour = Str(Mid(NowTime, 9, 2)): kHour = Str(Mid(tKey, 9, 2))
        nMinute = Str(Mid(NowTime, 11, 2)): kMinute = Str(Mid(tKey, 11, 2))
        nSecond = Str(Mid(NowTime, 13, 2)): kSecond = Str(Mid(tKey, 13, 2))
        If Val(nYear) > Val(kYear) Then
            GoTo OldKey
        ElseIf Val(nYear) < Val(kYear) Then
            GoTo FinishAction
        ElseIf Val(nYear) = Val(kYear) Then
            If Val(nMonth) > Val(kMonth) Then
                GoTo OldKey
            ElseIf Val(nMonth) < Val(kMonth) Then
                GoTo FinishAction
            ElseIf Val(nMonth) = Val(kMonth) Then
                If Val(nDay) > Val(kDay) Then
                    GoTo OldKey
                ElseIf Val(nDay) < Val(kDay) Then
                    GoTo FinishAction
                ElseIf Val(nDay) = Val(kDay) Then
                    If Val(nHour) > Val(kHour) Then
                        GoTo OldKey
                    ElseIf Val(nHour) < Val(kHour) Then
                        GoTo FinishAction
                    ElseIf Val(nHour) = Val(kHour) Then
                        If Val(nMinute) > Val(kMinute) Then
                            GoTo OldKey
                        ElseIf Val(nMinute) < Val(kMinute) Then
                            GoTo FinishAction
                        ElseIf Val(nMinute) = Val(kMinute) Then
                            If Val(nSecond) >= Val(kSecond) Then
                                GoTo OldKey
                            ElseIf Val(nSecond) < Val(kSecond) Then
                                GoTo FinishAction
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        GoTo ErrProve
    End If
End If

GoTo ErrProve
Exit Sub
ErrProve:
    Call MsgBox("“U" & Key & "”" & "是一个无效的激活密钥！", vbCritical, "无效的密钥")
Exit Sub
ErrProveU:
    Call MsgBox("“" & Text1.Text & "”" & "是一个不合规的激活密钥！", vbCritical, "不合规的密钥")
Exit Sub
OldKey:
    Call MsgBox("“U" & Key & "”" & "是一个已经过期的激活密钥！", vbCritical, "过期的密钥")
    
Exit Sub
FinishAction:
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Act.prove") For Output As #1
    Print #1, "免激活版本" & vbCrLf & _
              "激活密钥有效期至2026年10月9日23:59:59"
    'Print #1, (Mid(Key, 1, 7) & Mid(Key, 11, 7))
    Close #1
    Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Act.prove", vbSystem + vbHidden + vbReadOnly)
    yn = MsgBox("激活成功！要现在开始使用软件吗？", vbYesNo + vbInformation, "激活成功！")
    If yn = vbYes Then
        Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp")
        Shell (App.Path & "/LJXPopupKiller.exe")
        Call EndProgram
        Timer1.Enabled = True
    Else
        Call EndProgram
    End If
Exit Sub
Errt:
    If Err.Number = 13 Then
        GoTo ErrProveU
    End If
    Call MsgBox("Ac_C1_Cil：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "激活时错误")
End Sub

Private Sub Command2_Click()
a = MsgBox("你确定要取消这次激活吗？", vbOKCancel + vbExclamation, "取消激活")
If a = vbOK Then
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp") <> "" Then
        Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp")
        Call EndProgram
    End If
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "显示密钥" Then
    Text1.PasswordChar = ""
    Command3.Caption = "隐藏密钥"
ElseIf Command3.Caption = "隐藏密钥" Then
    Text1.PasswordChar = "*"
    Command3.Caption = "显示密钥"
End If
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
If HavAdminPer = True Then
    Me.Caption = " [Admin] " & Me.Caption
End If

Call MsgBox("这是一个免激活版本的LJX弹窗杀手，可以直接点击“激活”按钮激活软件", vbOKOnly + vbInformation, "免激活")

Text1.Text = "U220228671100099125385494549"

Command3.Enabled = False
Text1.Enabled = False

End Sub

Private Sub Text1_Change()
On Error Resume Next
Dim X
If Len(Text1.Text) > 27 Then
    Text1.Text = Mid(Text1.Text, 1, 28)
    e = 28
    Text1.SelStart = 28
End If
If Len(Text1.Text) = 28 Then
    Label2.ForeColor = &HFF00&
    Label2.Caption = "激活密钥最大为28位！你已经输入了28位。"
Else
    Label2.ForeColor = &HFF&
    Label2.Caption = ""
End If
If Mid(Text1.Text, 1, 1) <> "U" Then
    Text1.Text = "U" & Mid(Text1.Text, 1, Len(Text1.Text) - 1)
    Text1.SelStart = Len(Text1.Text)
    Label2.ForeColor = &HFF&
    Label2.Caption = "激活密钥必须以U开头！"
ElseIf Len(Text1.Text) <> 28 Then
    Label2.ForeColor = &HFF&
    Label2.Caption = ""
End If
If Label2.Caption = "" Then
    Acti.Height = 1815
Else
    Acti.Height = 2070
End If
Dim hStr As Boolean
hStr = False
For X = 1 To Len(Text1.Text)
    If X > 1 Then
        If Mid(Text1.Text, X, 1) <> "0" And Mid(Text1.Text, X, 1) <> "1" And Mid(Text1.Text, X, 1) <> "2" And Mid(Text1.Text, X, 1) <> "3" And Mid(Text1.Text, X, 1) <> "4" And Mid(Text1.Text, X, 1) <> "5" And Mid(Text1.Text, X, 1) <> "6" And Mid(Text1.Text, X, 1) <> "7" And Mid(Text1.Text, X, 1) <> "8" And Mid(Text1.Text, X, 1) <> "9" Then
            hStr = True
        End If
    End If
Next
If hStr = True Then
    Acti.Height = 2070
    Label2.ForeColor = &HFF&
    Label2.Caption = "激活密钥中除了前缀U之外的字符必须都是数字！"
ElseIf hStr = False And Label2.Caption = "激活密钥中除了前缀U之外的字符必须都是数字！" Then
    Acti.Height = 1815
    Label2.Caption = ""
    Label2.ForeColor = &HFF&
End If
End Sub

Private Sub Timer1_Timer()
Call EndProgram
End Sub
