VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LJX����ɱ��-�������"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9036
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6672
   ScaleWidth      =   9036
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFFF00&
      Caption         =   "���ã�Setting��"
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "��ʾ��������LJX����ɱ�ֵ��������ݣ�����Ը�������"
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   6480
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   6000
      Top             =   0
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "���������Ϣ"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "�鿴LJX����ɱ�ֵĵ�������ⱨ��"
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "ǿ������LJX����ɱ��"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "ǿ�������������е�LJX����ɱ��"
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "ǿ�ƽ�������"
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "ǿ��ֹͣ�������е�LJX����ɱ��"
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   5520
      Top             =   0
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0080FFFF&
      Caption         =   "����LJX����ɱ��"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "��������LJX����ɱ��"
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FF00FF&
      Caption         =   "����������"
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "���LJX����ɱ�ֵ��������"
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FF00&
      Caption         =   "����LJX����ɱ��"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "����LJX����ɱ��"
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF00FF&
      Caption         =   "��ӵ���������"
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "��LJX����ɱ����ӵ�����������"
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������������(ѡ����ʵ�ģʽ����Ӧ���Ե�����)��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   9015
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ɾ����ֻɾ�������ĳ���û���κ�Ӱ�죩"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3960
         Width           =   8775
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����ģʽ  (û���κ�Ӱ�죬�����������ӳ�)  "
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   8775
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "������ģʽ  (�����ܼ���û��Ӱ�죬���������ӳ�)  "
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   8775
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ͨģʽ  (�����ܻ���΢С��Ӱ�죬���������ӳ�)  "
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   8775
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ǿ��ģʽ  (���ܻ��������һ��Ӱ�죬����΢С���ӳ�)  "
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   8775
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����ģʽ  (���ܻ�������н϶�Ӱ�죬����û���ӳ�)  "
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   8775
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����ģʽ  (����Ӱ��϶࣬��ɾ�������ĳ���)  "
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   8775
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����LJX����ɱ��"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "��ʾ����LJX����ɱ�ֵ���Ϣ"
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "����Ҫ���صĵ���"
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "�鿴��ǰ������Ϊ�����ء��ĵ���"
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "ֹͣLJX����ɱ�ֵ�����"
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "ʹ�������е�LXJ����ɱ��ֹͣ����"
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�رմ˿������"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "�����趨���رտ������"
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   9000
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��ǰ�У�������������������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2
Private ExeNumber
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Timer1.Enabled = True
Else
    Timer1.Enabled = False
End If
End Sub

Private Sub Command11_Click()
On Error Resume Next

Call MsgBox("LJX����ɱ��" & vbCrLf & _
            "�汾v" & App.Major & "." & App.Minor & "." & App.Revision & "[�⼤��]" & vbCrLf & _
            "����״̬������ѳɹ�������Ϊ�⼤��汾��" & vbCrLf & _
            "������Կ��Ч����2026��10��9�� 23:59:59", vbOKOnly + vbInformation, "���������Ϣ")

Exit Sub
'Form4.Show
End Sub

Private Sub Command12_Click()
On Error Resume Next
Setting.Show
End Sub

Private Sub Command2_Click()
On Error GoTo Errt
KillerStart = False
Dim p
p = MsgBox(Languages(langNumber)(95), vbOKCancel + vbExclamation, Languages(langNumber)(59))
If p = vbOK Then
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp") <> "" Then
        Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp")
    End If
    
    End
    
End If
Exit Sub
Errt:
MsgBox ("F1_C2_Cli" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command3_Click()
On Error GoTo Errt
Dim a
a = MsgBox(Languages(langNumber)(52), vbOKCancel + vbExclamation, Languages(langNumber)(59))
If a = vbOK Then
    KillerStart = False
    Unload KillerMain
    KillerStart = True
    Load KillerMain
    Call MsgBox(Languages(langNumber)(107), vbOKOnly + vbInformation, Languages(langNumber)(59))
End If
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("F1_C3_Cli" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Errt
If UnloadMode = 0 Then
    Cancel = True
Else
    Exit Sub
End If
If KillerStart = True Then
    i = MsgBox(Languages(langNumber)(53), vbOKCancel + vbQuestion, Languages(langNumber)(59))
    If i = vbOK Then
        Me.Hide
    End If
Else
    i = MsgBox(Languages(langNumber)(96), vbOKCancel + vbQuestion, Languages(langNumber)(59))
    If i = vbOK Then
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp")
        End If
        End
    End If
End If
Exit Sub
Errt:
MsgBox ("F1_C1_Cli" & Err.Number & vbCrLf & Err.Description)
End
End Sub

Private Sub Timer1_Timer()
On Error GoTo Errt
KillerStart = True
Form1.Enabled = True
Form1.Caption = Languages(langNumber)(0)
Timer1.Enabled = False
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("F1_Ti1_Tim" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command1_Click()
On Error GoTo Errt
If KillerStart = True Then
    i = MsgBox(Languages(langNumber)(53), vbOKCancel + vbQuestion, Languages(langNumber)(59))
    If i = vbOK Then
        Me.Hide
    End If
Else
    i = MsgBox(Languages(langNumber)(96), vbOKCancel + vbQuestion, Languages(langNumber)(59))
    If i = vbOK Then
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp")
        End If
        Call EndProgram
    End If
End If
Exit Sub
Errt:
MsgBox ("F1_C1_Cli" & Err.Number & vbCrLf & Err.Description)
End
End Sub

Private Sub Command10_Click()
On Error GoTo Errt
i = MsgBox(Languages(langNumber)(47), vbOKCancel, Languages(langNumber)(59))
If i = vbOK Then
    KillerStart = False
    Form1.Enabled = False
    Form1.Caption = Languages(langNumber)(48)
    Timer1.Interval = 50
    Timer1.Enabled = True
End If
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("F1_C10_Cli" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command4_Click()
On Error GoTo Errt

KillerStart = False
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("F1_C4_Cli" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command5_Click()
LoadAll
On Error Resume Next
Form2.Show
End Sub

Private Sub Command6_Click()
On Error Resume Next
frmAbout.Show
End Sub

Private Sub Command7_Click()
On Error GoTo Errt

Dim a

Dim i
i = MsgBox(Languages(langNumber)(40), vbOKCancel + vbExclamation, Languages(langNumber)(59))
If i = vbCancel Then
    Exit Sub
End If
a = Dir(App.Path & "\LJXPopupKiller.exe")
If a = "" Then
    f = MsgBox(Languages(langNumber)(42), vbOKOnly, Languages(langNumber)(43))
    Exit Sub
End If
If i = vbOK Then
    Set w = CreateObject("wscript.shell")
    w.regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & "LJXPopupKiller", App.Path & "\LJXPopupKiller.exe -1"
End If
MsgBox (Languages(langNumber)(41))
Exit Sub
Errt:
MsgBox ("F1_C7_Cli" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command8_Click()
On Error GoTo Errt
If KillerStart = True Then
    Call MsgBox("�����Ѿ�������", vbOKOnly + vbExclamation, "�����Ѿ�����")
End If
KillerStart = True
Load KillerMain
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("����:" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command9_Click()
Call GetProgramRunState

If KillerStart = True Then
    Call MsgBox(Languages(langNumber)(36), vbInformation, Languages(langNumber)(37))
Else
    Call MsgBox(Languages(langNumber)(38), vbInformation, Languages(langNumber)(39))
End If
End Sub
Private Sub Form_Loadlanguage()
On Error Resume Next
'''''
'langNumber = 0
If HasOtherLanguage = True Then
    Me.Caption = Languages(langNumber)(0)
    Frame1.Caption = Languages(langNumber)(4)
    Dim X
    For X = 1 To 7
        Me("Option" & X).Caption = Languages(langNumber)(X + 4)
    Next
    Dim nL
    nL = 12
    For X = 1 To 12
        Me("Command" & X).Caption = Languages(langNumber)(nL)
        Me("Command" & X).ToolTipText = Languages(langNumber)(nL + 1)
        nL = nL + 2
    Next
End If

If HavAdminPer = True Then
    Me.Caption = " [Admin] " & Me.Caption
End If

' Languages(langNumber)(0) & Languages(langNumber)(1)
End Sub
Private Sub Form_Load()
'On Error GoTo Errt
Timer1.Enabled = False
WindowState = vbNormal



Call Form_Loadlanguage
'Call LoadModes



Call GetProgramRunState

If KillerStart = True Then
    Load KillerMain
    Call GetProgramRunState
End If
    
'
Exit Sub
Errt:
Call MsgBox("Sta_F1_Loa" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
If UnloadMode <> 0 Then
    Cancel = -1
End If
End Sub

Private Sub Option1_Click()
On Error GoTo Errt:
If Option1.Value = True And Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode1.ltmp") = "" Then
    a = MsgBox(Languages(langNumber)(57), vbOKCancel + vbExclamation, Languages(langNumber)(58))
    If a = vbOK Then
        If Option1.Value = True Then
            Dim X
            For X = 1 To 7
                If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp") <> "" Then
                    Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp")
                End If
            Next
            Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode1.ltmp") For Output As #1
            Close #1
        End If
    Else
         Option1.Value = False
         Call LoadModes
    End If
End If
Exit Sub
Errt:
Call MsgBox("F1_Op1_Cli" & Languages(langNumber)(62) & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option2_Click()
On Error GoTo Errt
If Option2.Value = True Then
    Dim X
    For X = 1 To 7
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp")
        End If
    Next
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode2.ltmp") For Output As #1
    Close #1
End If
Exit Sub
Errt:
Call MsgBox("F1_Op2_Cli" & Languages(langNumber)(62) & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option3_Click()
On Error GoTo Errt
If Option3.Value = True Then
    Dim X
    For X = 1 To 7
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp")
        End If
    Next
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode3.ltmp") For Output As #1
    Close #1
End If
Exit Sub
Errt:
Call MsgBox("F1_Op3_Cli" & Languages(langNumber)(62) & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option4_Click()
On Error GoTo Errt
If Option4.Value = True Then
    Dim X
    For X = 1 To 7
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp")
        End If
    Next
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode4.ltmp") For Output As #1
    Close #1
End If
Exit Sub
Errt:
Call MsgBox("F1_Op4_Cli" & Languages(langNumber)(62) & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option5_Click()
On Error GoTo Errt
If Option5.Value = True Then
    Dim X
    For X = 1 To 7
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp")
        End If
    Next
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode5.ltmp") For Output As #1
    Close #1
End If
Exit Sub
Errt:
Call MsgBox("F1_Op5_Cli" & Languages(langNumber)(62) & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option6_Click()
On Error GoTo Errt
If Option6.Value = True Then
    Dim X
    For X = 1 To 7
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp")
        End If
    Next
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode6.ltmp") For Output As #1
    Close #1
End If
Exit Sub
Errt:
Call MsgBox("F1_Op6_Cli" & Languages(langNumber)(62) & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option7_Click()
On Error GoTo Errt
If Option7.Value = True And Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode7.ltmp") = "" Then
    a = MsgBox(Languages(langNumber)(56), vbOKCancel + vbExclamation, Languages(langNumber)(58))
    If a = vbOK Then
        If Option7.Value = True Then
            Dim X
            For X = 1 To 73
                If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp") <> "" Then
                    Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode" & X & ".ltmp")
                End If
            Next
            Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode7.ltmp") For Output As #1
            Close #1
        End If
    Else
        Option7.Value = False
        Call LoadModes
    End If
End If
Exit Sub
Errt:
Call MsgBox("F1_Op7_Cli" & Languages(langNumber)(62) & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp") = "" Then
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp") For Output As #1
    Close #1
End If
Exit Sub
End Sub

Private Sub Timer3_Timer()
Call GetProgramRunState
End Sub
