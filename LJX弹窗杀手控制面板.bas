Attribute VB_Name = "Module1"
Public Declare Function IsUserAnAdmin Lib "shell32" () As Long


Public MyName As String                 '�û���
Public Pops(0 To 1023) As String        '�����صĵ���·��
Public PopsP(0 To 1023) As String       '�����صĵ�������
Public MaxPops As Long                  'Ŀǰ���Ŀ�д�뵯��
Public StartMe As Boolean
Public IsOne As Boolean
Public Act As Boolean                   '����Ѿ����
Public KillerStart As Boolean           '����Ƿ����У�
Public AutoRun As Boolean               '�Ƿ��Զ����У�ͬ�����У�
Public HideMain As Boolean              '�Ƿ�����������
Public AddPopup As Boolean

Public Languages(0 To 10)               '�Ѷ�ȡ�Ľ�������
Public langNames(0 To 10)               '�������Ʊ�
Public langNumber                       '��ǰ���Ա��
Public HasOtherLanguage As Boolean      '�Ƿ�ӵ����������
Public Settings(0 To 100)               '������
Public ShouldRefLoadPopups As Boolean   'KillerMain�Ƿ�Ӧ��ˢ���ļ��б�

Public ExeNumber
Public HavAdminPer As Boolean           '�Ƿ�ӵ�й���ԱȨ��
Public ldP

Function UnloadAll()
Unload Acti
Unload Form1
Unload Form2
Unload Form3
'Unload Form4
Unload Form5
Unload frmAbout
Unload frmSplash
Unload InfoFrm
Unload KillerMain
Unload Reset
Unload Setting
End Function
Function GetExeNumber(ExeName) As Variant
Dim objWMIService As Object
Dim colProcessList As Object
Dim objProcess As Object
Dim objProType As Object
Dim strResult As String
Dim strTmp As String
Set objWMIService = GetObject("winmgmts:" & "{impersonationlevel=impersonate}!\\.\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='" & ExeName & "'")
GetExeNumber = colProcessList.Count
End Function
Public Function CheckExeIsRun(ExeName As String) As Boolean
On Error GoTo Errt
Dim WMI
Dim Obj
Dim Objs
CheckExeIsRun = False
Set WMI = GetObject("WinMgmts:")
Set Objs = WMI.InstancesOf("Win32_Process")
Do
    DoEvents
    For Each Obj In Objs
        If (InStr(UCase(ExeName), UCase(Obj.Description)) <> 0) Then
            CheckExeIsRun = True
            ExeNumber = InStr(UCase(ExeName), UCase(Obj.Description))
            If Not Objs Is Nothing Then Set Objs = Nothing
            If Not WMI Is Nothing Then Set WMI = Nothing
            Exit Function
        End If
    Next
    Exit Do
Loop
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
Exit Function
Errt:
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
End Function
Function LoadModes()
Dim u As Boolean
u = False
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode1.ltmp") <> "" Then
    Form1.Option1.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode2.ltmp") <> "" Then
    Form1.Option2.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode3.ltmp") <> "" Then
    Form1.Option3.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode4.ltmp") <> "" Then
    Form1.Option4.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode5.ltmp") <> "" Then
    Form1.Option5.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode6.ltmp") <> "" Then
    Form1.Option6.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode7.ltmp") <> "" Then
    Form1.Option7.Value = True
    u = True
End If
If u = False Then
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode4.ltmp") For Output As #1
    Close #1
    u = True
    Form1.Option4.Value = True
End If
End Function
Sub Main()
On Error GoTo Errt
KillerStart = False
If IsOne = False Then
    IsOne = True
    WindowState = vbNormal
End If
If Dir(Command) <> "" And Command <> "" Then
    '�Ƿ�Ϊ���ļ���ӵ�LJX����ɱ�������б�
    AddPopup = True
End If

If Command = "-1" Then
    '����������
    AutoRun = True
    HideMain = True
ElseIf Command = "" Or Command = "0" Then
    '�������������
    AutoRun = False
    HideMain = False
ElseIf Command = "1" Then
    '����������岢���г���
    AutoRun = True
    HideMain = False
End If
If HideMain = False Then
    '�������ش���
    frmSplash.Show
    frmSplash.SetFocus
Else
    Load frmSplash
End If
Call InitPrograms
Load Form1
Exit Sub
Errt:
Call MsgBox("Main()��ʼ������" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub
Public Function InitPrograms()
On Error GoTo Errt
Call RefrmSplash(1)
Call verAdminPer
If HavAdminPer = False Then
    Call MsgBox("���û�й���ԱȨ�ޣ���ʹ�ù���ԱȨ�����г���", vbOKOnly + vbCritical, "��Ҫ����Ȩ��")
    Call EndProgram
End If
Call LoadSetting
Call LoadLanguages
Call LoadAll
Call LoadModes
Call GetProgramRunState
Exit Function
Errt:
Call MsgBox("InitPrograms()���д���" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Function
Public Function LoadAll()
On Error GoTo Errt
Call RefrmSplash(5)
Module1.MaxPops = 0
MyName = Environ("USERNAME")
'�������
For X = 0 To 1023
    Module1.Pops(X) = ""
    PopsP(X) = ""
Next
'----------
Dim a
Dim b
Dim c
Dim d
Dim e
Dim f
Dim g

a = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller", vbDirectory + vbSystem)
If a = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller", vbSystem)

b = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf", vbDirectory + vbSystem)
If b = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf", vbSystem)

c = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Number", vbDirectory + vbSystem)
If c = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Number")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Number", vbSystem)

d = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Log", vbDirectory + vbSystem)
If d = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Log")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Log", vbSystem)

e = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops", vbDirectory + vbSystem)
If e = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops", vbSystem)

f = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp", vbDirectory + vbSystem)
If f = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp", vbSystem)

g = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\���ڸ��ļ��е�˵��.txt", vbSystem + vbReadOnly)
If g = "" Then
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\���ڸ��ļ��е�˵��.txt") For Output As #1
    Print #1, "����LJX����ɱ�ֵ���Ҫ�ļ��У��벻Ҫ����������κ��ļ���"
    Close #1
    Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\���ڸ��ļ��е�˵��.txt", vbSystem + vbReadOnly)
End If

Dim q
Dim Stri As String
Dim r As Long
Dim L
r = 0
Call RefrmSplash(6)
For X = 0 To 1023
    Stri = ""
    q = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx", vbSystem)
    If q <> "" Then
        If FileLen("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx") = 0 Then
            Close #1
            DelFile (X)
        Else
            Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx") For Input As #1
            Line Input #1, Stri
            Module1.Pops(r) = Stri
            Module1.MaxPops = Module1.MaxPops + 1
            t = Len(Stri)
            If Stri <> "" Then                  '�������ļ��Է�ֹ����
                For L = 1 To t - 1
                    tmp = Mid(Stri, (t - L), 1)
                    If tmp = "\" Then
                        Exit For
                    End If
                Next
                Dim FileN As String
                FileN = Mid(Stri, (t - L + 1), L)
                PopsP(r) = FileN
                r = r + 1
            End If
            Close #1
        End If
    End If
Next


Call RefrmSplash(10)
'ɾ�����ļ�
Dim Y As String
For X = 0 To 1023
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx") <> "" Then
        Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx") For Input As #1
        Line Input #1, Y
        If Y = "" Then
            Close #1
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx")
        End If
        Close #1
    End If
Next

Call RefrmSplash(11)
'����Ƿ����һ���������
If GetExeNumber("LJXPopupKiller.exe") > 1 And AddPopup = False Then
    Call MsgBox(Languages(langNumber)(97) & vbCrLf & Languages(langNumber)(98), vbOKOnly + vbExclamation, Languages(langNumber)(59))
    End
    Exit Function
End If
Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp") For Output As #1
Close #1
StartMe = True

If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Act.prove", vbSystem + vbHidden + vbReadOnly) = "" Then
    Act = False
    bo = MsgBox(Languages(langNumber)(99), vbOKCancel + vbExclamation, Languages(langNumber)(100))
    If bo = vbOK Then
        Acti.Show
        Exit Function
    Else
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\ControlRun.ltmp")
            Call EndProgram
        End If
    End If
End If


'����������
Call LoadSetting
Call RefrmSplash(7)
'��ʼ������ɱ������
'Load KillerMain
Refform1

If AddPopup = True Then
    Form3.Show
    Form3.Text1.Text = Command
    Exit Function
End If

'������ģʽ������������
If HideMain = False Then
    RefrmSplash (9)
    Form1.Show
End If
If AutoRun = True Or Settings(1) = True Then
    KillerStart = True
    Load KillerMain
End If
If HideMain = True And AutoRun = False Then
    MsgBox ("Load_All�������󣺴����������ʽ��")
    Form1.Show
End If
Call RefrmSplash(8)
Set w = CreateObject("wscript.shell")
If Settings(2) = True Then
    Call w.regwrite("HKEY_CLASSES_ROOT\*\shell\��ӵ�LJX����ɱ�������б�\", "")
    Call w.regwrite("HKEY_CLASSES_ROOT\*\shell\��ӵ�LJX����ɱ�������б�\Command\", App.Path & "\LJXPopupKiller.exe %0")
ElseIf Setting(2) = False Then
    On Error Resume Next
    Call w.regdelete("HKEY_CLASSES_ROOT\*\shell\��ӵ�LJX����ɱ�������б�\command\")
    Call w.regdelete("HKEY_CLASSES_ROOT\*\shell\��ӵ�LJX����ɱ�������б�\")
End If

Call GetProgramRunState
Exit Function

Errt:
Call MsgBox("Load_All��������" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End

End Function

Public Function DelFile(FileNumber)
On Error GoTo Errt
Form2.Hide
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & FileNumber & ".ltx") = "" Then
    Call MsgBox(Languages(langNumber)(101), vbOKOnly + vbExclamation)
End If
'�ȱ���Ҫ�����ŵ��ļ���
Dim X As Long
Dim r(0 To 1023) As String
Dim Maxr As Long
Maxr = 0
For X = (FileNumber + 1) To 1023
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx") <> "" Then
        Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx") For Input As #1
        Line Input #1, r(X)
        Maxr = Maxr + 1
        Close #1
    Else
        Exit For
    End If
Next
Close #1
'ɾ���ļ�
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & FileNumber & ".ltx") <> "" Then
    Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & FileNumber & ".ltx")
    Module1.MaxPops = Module1.MaxPops - 1
End If
'ɾ�����ļ�
For X = (FileNumber + 1) To 1023
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx") <> "" Then
        Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx")
    End If
Next
'�������ļ�
Dim rfn
rfn = FileNumber
For X = 1 To Maxr
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & rfn & ".ltx") For Output As #1
    Print #1, r(rfn + 1)
    rfn = rfn + 1
    Close #1
Next
Exit Function
Call LoadAll

Unload All
Errt:
Call MsgBox("DelFile()����" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Function
Function GetProgramRunState()
On Error Resume Next
If KillerStart = False Then
    Form1.Label1.ForeColor = &HFF&
    Form1.Label1.Caption = Languages(langNumber)(3)
Else
    Form1.Label1.ForeColor = &H0&
    Form1.Label1.Caption = Languages(langNumber)(1) & Module1.MaxPops & Languages(langNumber)(2)
End If
End Function

Public Function Refform1()
If Module1.Pops(0) = "" Then
    Module1.MaxPops = 0
End If
Call GetProgramRunState
End Function

Function LoadLanguages()
Call RefrmSplash(4)
Dim aN
aN = 0

Dim tmpTxt As String
Dim n

If Dir(App.Path & "/����.ini") <> "" Then
    Open (App.Path & "/����.ini") For Input As #10
    n = 0
    ReDim allTxt((LOF(10) - 1)) As String
    Do While Not EOF(10)
        Line Input #10, tmpTxt
        allTxt(n) = tmpTxt
        n = n + 1
        DoEvents
    Loop
    Languages(aN) = allTxt
    aN = aN + 1
    langNames(aN - 1) = "����"
    Close #10
End If
If Dir(App.Path & "/English.ini") <> "" Then
    Open (App.Path & "/English.ini") For Input As #10
    n = 0
    ReDim allTxt((LOF(10) - 1)) As String
    Do While Not EOF(10)
        Line Input #10, tmpTxt
        allTxt(n) = tmpTxt
        n = n + 1
        DoEvents
    Loop
    Languages(aN) = allTxt
    aN = aN + 1
    langNames(aN - 1) = "English"
    Close #10
End If
If Dir(App.Path & "/�����ܧڧ� ��٧��.ini") <> "" Then
    Open (App.Path & "/�����ܧڧ� ��٧��.ini") For Input As #10
    n = 0
    ReDim allTxt((LOF(10) - 1)) As String
    Do While Not EOF(10)
        Line Input #10, tmpTxt
        allTxt(n) = tmpTxt
        n = n + 1
        DoEvents
    Loop
    Languages(aN) = allTxt
    aN = aN + 1
    langNames(aN - 1) = "�����ܧڧ� ��٧��"
    Close #10
End If
If Dir(App.Path & "/�ձ��Z.ini") <> "" Then
    Open (App.Path & "/�ձ��Z.ini") For Input As #10
    n = 0
    ReDim allTxt((LOF(10) - 1)) As String
    Do While Not EOF(10)
        Line Input #10, tmpTxt
        allTxt(n) = tmpTxt
        n = n + 1
        DoEvents
    Loop
    Languages(aN) = allTxt
    aN = aN + 1
    langNames(aN - 1) = "�ձ��Z"
    Close #10
End If

langNumber = Settings(0)

If aN = 0 Then
    HasOtherLanguage = False
Else
    HasOtherLanguage = True
End If
End Function
Function LoadSetting()
Call RefrmSplash(3)
If Dir(App.Path & "/Setting.ini") <> "" Then
    Open (App.Path & "/Setting.ini") For Input As #11
    Dim t
    Dim n
    n = 0
    Do While Not EOF(11)
        Line Input #11, t
        Settings(n) = t
        n = n + 1
        DoEvents
    Loop
    Close #11
Else
    Open (App.Path & "/Setting.ini") For Output As #11
    Print #11, 0 & vbCrLf
    Close #11
End If

End Function
Function SaveSetting()
If Dir(App.Path & "/Setting.ini") <> "" Then
    Open (App.Path & "/Setting.ini") For Output As #11
    Dim X
    Dim s
    s = ""
    For Each X In Settings
        If X <> "" Then
            s = s & X & vbCrLf
        End If
    Next
    Print #11, s
    Close #11
Else
    Open (App.Path & "/Setting.ini") For Output As #11
    Print #11, 0 & vbCrLf
    Close #11
End If
End Function
Function ShowInfo(Caption)
InfoFrm.Show
InfoFrm.Caption = Caption
End Function
Function EleInArray(Element As Variant, Arr As Variant) As Boolean
Dim X
For Each X In Arr:
    If X = Element Then
        EleInArray = True
        Exit Function
    End If
Next
EleInArray = False
End Function
Function RemoveStrInStr(Str1 As String, Str2 As String) As String
Dim X
Dim eStr
eStr = ""
For X = 1 To Len(Str2)
    If Str1 <> Mid(Str2, X, 1) Then
        eStr = eStr & Mid(Str2, X, 1)
    End If
Next
RemoveStrInStr = eStr
End Function
Function verAdminPer() As Boolean
'����Ƿ���й���ԱȨ��
RefrmSplash (2)
On Error GoTo Errt

If IsUserAnAdmin() Then
    HavAdminPer = True
Else
    HavAdminPer = False
    'Call MsgBox("����ԱȨ�޻�ȡʧ�ܣ���ʹ�ù���ԱȨ�����г���", vbOKOnly + vbCritical, "��Ҫ����Ȩ��")
End If

Exit Function
Errt:
Call MsgBox("F_verAdminPer()����" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "����")

End Function
Function RefrmSplash(n)
'label5.width.max=6840
Dim p
p = ldP
If n = 0 Then
    frmSplash.Label6.Caption = "׼�����س���"
    p = p + 0.1
ElseIf n = 1 Then
    frmSplash.Label6.Caption = "��ʼ������"
     p = p + 0.05
ElseIf n = 2 Then
    frmSplash.Label6.Caption = "������ԱȨ��"
    p = p + 0.1
ElseIf n = 3 Then
    frmSplash.Label6.Caption = "���������ļ�"
    p = p + 0.05
ElseIf n = 4 Then
    frmSplash.Label6.Caption = "���������ļ�"
    p = p + 0.2
ElseIf n = 5 Then
    frmSplash.Label6.Caption = "���ڼ��س��������ļ�"
    p = p + 0.05
ElseIf n = 6 Then
    frmSplash.Label6.Caption = "���ڶ�ȡ�����������ļ�"
    p = p + 0.15
ElseIf n = 7 Then
    frmSplash.Label6.Caption = "��ʼ��������������"
    p = p + 0.05
ElseIf n = 8 Then
    frmSplash.Label6.Caption = "У��ע���ֵ"
    p = p + 0.05
ElseIf n = 9 Then
    frmSplash.Label6.Caption = "����׼���û�����"
    p = 1
End If
ldP = p
frmSplash.Label5.Width = 6840 * p
If p >= 1 Then
    Unload frmSplash
End If
End Function

Public Function EndProgram()
'ֹͣ������
KillerStart = False
'�������
Call KillerMain.CI
'��������
End
End Function

