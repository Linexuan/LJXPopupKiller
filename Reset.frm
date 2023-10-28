VERSION 5.00
Begin VB.Form Reset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "重置LJX弹窗杀手"
   ClientHeight    =   5760
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   4368
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.8
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   4368
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Line Line5 
      X1              =   4200
      X2              =   4200
      Y1              =   1560
      Y2              =   1920
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4200
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   120
      Y1              =   1560
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4200
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      Caption         =   "50% "
      BeginProperty Font 
         Name            =   "思源黑体 CN Regular"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "<这有着一定的风险>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "单击“开始”以开始重置LJX弹窗杀手"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4200
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "Reset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ResetStart As Boolean
Private ReNum As Variant
Private Sub Form_Loadlanguage()
If HasOtherLanguage = True Then
    Me.Caption = Languages(langNumber)(113)
    'Check1.Caption = Languages(langNumber)(114)
    'Check2.Caption = Languages(langNumber)(116)
    Command1.Caption = Languages(langNumber)(129)
    Command2.Caption = Languages(langNumber)(130)
    Command3.Caption = Languages(langNumber)(128)
    Label3.Caption = Languages(langNumber)(134)
    Label4.Caption = Languages(langNumber)(135)
End If
If HavAdminPer = True Then
    Me.Caption = " [Admin] " & Me.Caption
End If
End Sub
Private Function St(Texts, State As Boolean)
If State = True Then
    Text1.Text = Text1.Text & Texts & vbCrLf
Else
    Text1.Text = Text1.Text & Texts
End If
End Function
Private Function Ad(Num As Long)
ReNum = ReNum + Num
Label1.Left = 120
Dim t
t = 4095 * (Num / 100)
Dim X ', u
'u = 0
'For x = 1 To 5000000
'    u = u + x
'Next
Label1.Width = Label1.Width + t
Label1.Caption = ReNum & "%"
End Function

Private Function proReset()
On Error GoTo Errt

'开始进行重置计划

Ad (0)
Call St("Initialization......", False)
Ad (1)
If Dir("C:\windows\System32\cmd.exe") = "" Then
    Call MsgBox(Languages(langNumber)(136), vbOKOnly + vbCritical, Languages(langNumber)(133))
    Call St("Fail", True)
    Call St("File not found:cmd.exe", False)
    Exit Function
End If
Ad (1)
If Dir("C:\windows\System32\reg.exe") = "" Then
    Call MsgBox(Languages(langNumber)(136), vbOKOnly + vbCritical, Languages(langNumber)(133))
    Call St("Fail", True)
    Call St("File not found:reg.exe", False)
    Exit Function
End If
Ad (1)
If Dir("C:\Windows\System32\mshta.exe") = "" Then
    Call MsgBox(Languages(langNumber)(136), vbOKOnly + vbCritical, Languages(langNumber)(133))
    Call St("Fail", True)
    Call St("File not found:mshta.exe", False)
    Exit Function
End If
Ad (1)

Call St("", True)
Ad (1)
Call St("Success.", True)
Ad (1)
Ad (1)
Call St("Verify file......", False)
Ad (1)
Shell ("cmd.exe /c mshta Verify file finish!")
Ad (1)
Call St("", True)
Ad (1)
Call St("┌Reset work.", True)
Ad (1)
Call St("├Removing files...", True)
Ad (1)
Call St("│├Removing root files...", False)
Ad (1)
Call Shell("cmd /c del /A /F /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\*.*", vbHide)
Ad (1)
Call Shell("cmd /c del /F /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\*.*", vbHide)
Ad (1)
Call St("Finish", True)
Ad (1)
Call St("│││├Removing inf files...", False)
Ad (1)
Call Shell("cmd /c del /F /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\Inf\*.*", vbHide)
Ad (1)
Call Shell("cmd /c del /F /A /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\Inf\*.*", vbHide)
Ad (1)
Call St("Finish", True)
Ad (1)
Call St("││││├Removing activation files...", False)
Ad (1)
Call Shell("cmd /c del /F /A /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\Inf\*.*", vbHide)
Ad (1)
Call St("Finish", True)
Ad (1)
Call St("││││├Removing log files...", False)
Ad (1)
Call Shell("cmd /c del /F /A /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\Log\*.*", vbHide)
Ad (1)
Call Shell("cmd /c del /F  /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\Log\*.*", vbHide)
Ad (1)
Call St("Finish", True)
Ad (1)
Call St("││││├Removing number files...", False)
Ad (1)
Call Shell("cmd /c del /F /A /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\Number\*.*", vbHide)
Ad (1)
Call Shell("cmd /c del /F  /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\Number\*.*", vbHide)
Ad (1)
Call St("Finish", True)
Ad (1)
Call St("││││├Removing popups files...", False)
Ad (1)
Call Shell("cmd /c del /F /A /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\Pops\*.*", vbHide)
Ad (1)
Call Shell("cmd /c del /F  /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\Pops\*.*", vbHide)
Ad (1)
Call St("Finish", True)
Ad (1)
Call St("││└└└Removing temp files...", False)
Ad (1)
Call Shell("cmd /c del /F /A /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\Temp\*.*", vbHide)
Ad (1)
Call Shell("cmd /c del /F  /C C:\Users\" & MyName & "admin\AppData\Roaming\LJXPopupKiller\Temp\*.*", vbHide)
Ad (1)
Call St("Finish", True)
Ad (1)
Call St("│├Removing files in App.Path...", True)
Ad (1)
Call St("│└└Removing Setting.ini...", False)
Ad (1)
If Dir(App.Path & "\Setting.ini") <> "" Then
    Kill (App.Path & "\Setting.ini")
End If
Ad (1)
Call St("Finish", True)
Ad (1)




Exit Function
Errt:
Call MsgBox(Languages(langNumber)(131) & Err.Number & " , " & Err.Description & vbCrLf & Languages(langNumber)(132), vbOKOnly + vbCritical, Languages(langNumber)(133))
Call St("", True)
Call St("", True)
Call St("Reset Fail", True)
Call St("Error Number:" & Err.Number, True)
Call St("Description:" & Err.Description, True)
Call St("Please click the cancel button to back to the LJXPopupKiller Main.", False)
Command2.Enabled = True
End Function
Private Sub Check1_Click()
End Sub

Private Sub Check2_Click()

End Sub

Private Sub Command1_Click()
Dim X
Dim q
q = MsgBox(Languages(langNumber)(112), vbOKCancel + vbExclamation, Languages(langNumber)(111))
If q = vbOK Then
    ResetStart = True
    For X = 1 To 50
        Me.Height = Me.Height + 90
        Me.Top = Me.Top - 45
        If Me.Height >= 6200 Then
            Exit For
        End If
    Next
    Me.Height = 6200
    
    Label1.Width = 0
    Label1.Left = 120
    
    'Check1.Enabled = False
    'Check2.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    ReNum = 0
    Call proReset
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
ResetStart = False
Me.Height = 1700

Call Form_Loadlanguage
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ResetStart = True Then
    Cancel = -1
Else
    Setting.Enabled = True
End If
End Sub

