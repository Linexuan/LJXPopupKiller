VERSION 5.00
Begin VB.Form Setting 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LJX弹窗杀手-设置"
   ClientHeight    =   4584
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   8220
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4584
   ScaleWidth      =   8220
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "显示“添加到LJX弹窗杀手拦截列表”右键菜单"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   480
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "重置文件"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "不保存并关闭"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "同步运行"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2532
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "保存并应用"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   276
      ItemData        =   "Setting.frx":0000
      Left            =   1440
      List            =   "Setting.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3612
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "语言："
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Loadlanguage()
If HasOtherLanguage = True Then
    Me.Caption = Languages(langNumber)(72)
    
    Label1.Caption = Languages(langNumber)(73)
    Label1.ToolTipText = Languages(langNumber)(104)
    Check1.Caption = Languages(langNumber)(102)
    Check1.ToolTipText = Languages(langNumber)(103)
    Command1.Caption = Languages(langNumber)(105)
    Command2.Caption = Languages(langNumber)(82)
    Command3.Caption = Languages(langNumber)(111)
    Combo1.Text = langNames(langNumber)
End If
If HavAdminPer = True Then
    Me.Caption = " [Admin] " & Me.Caption
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim i
i = MsgBox(Languages(langNumber)(106), vbOKCancel + vbQuestion, Languages(langNumber)(59))
If i = vbOK Then
    Settings(0) = Val(Combo1.ListIndex)
    If Check1.Value = 0 Then
        Settings(1) = False
    Else
        Settings(1) = True
    End If
    If Check2.Value = 0 Then
        Settings(2) = False
    Else
        Settings(2) = True
    End If
    
    
    Call SaveSetting
    Call UnloadAll
    Call InitPrograms
    Call UnloadAll
    Call InitPrograms
    'Command1.Value = True
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Call MsgBox("此功能正在开发中。", vbOKOnly + vbCritical, "开发中")
Exit Sub
Reset.Show
Me.Enabled = False
End Sub

Private Sub Form_Load()

Dim X
For Each X In langNames
    Combo1.AddItem (X)
Next

If Settings(1) = True Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If
If Settings(2) = True Then
    Check2.Value = 1
Else
    Check2.Value = 0
End If

Call Form_Loadlanguage
'Form1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub

