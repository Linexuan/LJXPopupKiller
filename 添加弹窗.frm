VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ӵ���"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ɾ���ļ�"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "������"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "���·��"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ȡ��"
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
      Caption         =   "���"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "������������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����ļ�·����"
      BeginProperty Font 
         Name            =   "����"
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
    Call MsgBox("��֧������Windows��Դ��������", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "csrss.exe" Then
    Call MsgBox("��֧������csrss.exe��", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "Client Server Runtime Process.exe" Then
    Call MsgBox("��֧������Client Server Runtime Process.exe��", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "smss.exe" Then
    Call MsgBox("��֧������smss.exe��", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "system.exe" Then
    Call MsgBox("��֧������system.exe��", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "cmd.exe" Then
    Call MsgBox("��֧������Windows������ʾ����", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "lsass.exe" Then
    Call MsgBox("��֧������lsass.exe��", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "winlogon.exe" Then
    Call MsgBox("��֧������Windows�û���¼����", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "spoolsv.exe" Then
    Call MsgBox("��֧������spoolsv.exe��", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "spoolst.exe" Then
    Call MsgBox("��֧������spoolst.exe��", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "svchost.exe" Then
    Call MsgBox("��֧������svchost.exe��", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If

If FileN = "ljxpopupkiller.exe" Then
    Call MsgBox("��֧������LJXPopupKiller.exe��", vbOKOnly + vbCritical, "��֧�����صĳ���")
    Text1.Text = ""
    Text2.Text = ""
    FileN = ""
End If



End Function
Private Sub Command1_Click()
On Error GoTo Errt
Dim OpenUrl As String
CD1.Filter = "��ִ���ļ�(*.exe)|*.EXE|��̬���ӿ�(*.dll)|*.DLL|�����ļ�(*.*)|*.*"
CD1.FilterIndex = 1
CD1.ShowOpen
If CD1.FileName = "" Then
    Exit Sub
End If
OpenUrl = CD1.FileName
Text1.Text = OpenUrl
'����ļ���
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
'�ж��ļ��Ƿ����Ҫ��
TestFileName (LCase(FileN))
Exit Sub
Errt:
Call MsgBox("�ļ��������" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "����")

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
    Call MsgBox("·������" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "����")
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
    Call MsgBox("����" & Stri & "�Ѿ���������" & vbCrLf & "��⵽�Ľ�������" & GetExeNumber(Stri), vbOKOnly + vbInformation, "����" & Stri & "������")
Else
    Call MsgBox("û���ҵ�ָ���Ľ��̣�" & Stri, vbOKOnly + vbCritical, "�Ҳ�������" & Stri)
End If
Exit Sub
Errt:
Call MsgBox("������ʱ��������" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "����")
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
Call MsgBox("����" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "   ")
Unload Me
End Sub

Private Sub Command6_Click()
On Error GoTo Errt
Dim t
t = MsgBox("��ѡ�ļ���������ɾ����ȷ��Ҫ��������", vbOKCancel + vbExclamation, "ɾ��ȷ��")
If t = vbOK Then
    Kill (Text1.Text)
    If Dir(Text1.Text) = "" Then
        MsgBox ("��ѡ�ļ��Ѿ��ɹ�ɾ����")
    Else
        Call MsgBox("�ļ�ɾ��ʧ�ܣ�", vbOKOnly + vbCritical, "�ļ�ɾ��ʧ��")
    End If
End If
Exit Sub
Errt:
Call MsgBox("�ļ�ɾ������" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
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
