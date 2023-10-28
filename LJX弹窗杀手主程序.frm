VERSION 5.00
Begin VB.Form KillerMain 
   Caption         =   "LJXµ¯´°É±ÊÖ-Ö÷³ÌÐò"
   ClientHeight    =   2604
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   4416
   LinkTopic       =   "KillerMain"
   ScaleHeight     =   2604
   ScaleWidth      =   4416
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "KillerMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objWMIService

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
 Const NIM_ADD = &H0
 Const NIM_DELETE = &H2
 Const NIF_ICON = &H2
 Const NIF_MESSAGE = &H1
 Const NIF_TIP = &H4
 Const WM_MOUSEMOVE = &H200
 Const WM_LBUTTONDBLCLK = &H203
Private Type NOTIFYICONDATA
cbSize As Long
hWnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Dim tray As NOTIFYICONDATA
Function Icon_Del() As Long
Dim IconVa As NOTIFYICONDATA
Dim L As Long
With IconVa
.hWnd = iHwnd
.uId = lIndex
.cbSize = Len(IconVa)
End With
Icon_Del = Shell_NotifyIcon(NIM_DELETE, IconVa)
End Function

Private Function LoadAll()
On Error GoTo Errt
Dim q
Dim Stri As String
Dim r As Long
Dim X As Long
Dim Y As Long
r = 0

For X = 0 To 1023
    KillerMainModble.PopsURL(X) = ""
    KillerMainModble.Pops(X) = ""
Next

For X = 0 To 1023
    Stri = ""
    q = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx")
    If q <> "" Then
        Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Pops\" & X & ".ltx") For Input As #1
        Line Input #1, Stri
        If Stri <> "" Then
            KillerMainModble.PopsURL(r) = Stri
            KillerMainModble.MaxPops = KillerMainModble.MaxPops + 1
        End If
        Close #1
        Dim t As Long
        Dim tmp As String
        t = Len(PopsURL(r))
        For Y = 1 To t
            tmp = Mid(PopsURL(r), (t - Y), 1)
            If tmp = "\" Then
                Exit For
            End If
        Next
        KillerMainModble.Pops(r) = Mid(PopsURL(r), (t - Y + 1), Y)
        r = r + 1
    End If
Next

If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode1.ltmp") <> "" Then
    Timer1.Interval = 1
    Mode = 1
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode2.ltmp") <> "" Then
    Timer1.Interval = 1
    Mode = 2
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode3.ltmp") <> "" Then
    Timer1.Interval = 50
    Mode = 3
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode4.ltmp") <> "" Then
    Timer1.Interval = 150
    Mode = 4
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode5.ltmp") <> "" Then
    Timer1.Interval = 500
    Mode = 5
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode6.ltmp") <> "" Then
    Timer1.Interval = 1000
    Mode = 6
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode7.ltmp") <> "" Then
    Timer1.Interval = 5000
    Mode = 7
End If
Exit Function
Errt:
MsgBox ("Sta_F1_Loa£º" & Err.Number & vbCrLf & Err.Description)
Call EndProgram
End Function

Private Sub Form_Load()
On Error GoTo Errt
'Call loadLang
'Timer1.Enabled = False
If KillerStart = True Then
    langNumber = 0
    Me.Hide
    tray.cbSize = Len(tray)
    tray.uId = vbNull
    tray.hWnd = Me.hWnd
    tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
    tray.uCallBackMessage = WM_MOUSEMOVE
    tray.hIcon = Me.Icon
    tray.szTip = Languages(langNumber)(79) & vbCrLf & Languages(langNumber)(80) & vbNullChar
    Shell_NotifyIcon NIM_ADD, tray
    
    
    If Dir("C:\Users\" & Environ("USERNAME") & "\AppData\Roaming\LJXPopupKiller\Inf\Act.prove", vbSystem + vbHidden + vbReadOnly) = "" Then
        Call MsgBox(Languages(langNumber)(76), vbCritical, Languages(langNumber)(81))
        Shell_NotifyIcon NIM_DELETE, tray
        Call EndProgram
    End If
    KillerMainModble.MaxPops = 0
    Me.Hide
    MyName = Environ("USERNAME")
    Call LoadAll
    
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode1.ltmp") <> "" Then
        Timer1.Interval = 1
        Mode = 1
    End If
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode2.ltmp") <> "" Then
        Timer1.Interval = 1
        Mode = 2
    End If
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode3.ltmp") <> "" Then
        Timer1.Interval = 50
        Mode = 3
    End If
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode4.ltmp") <> "" Then
        Timer1.Interval = 150
        Mode = 4
    End If
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode5.ltmp") <> "" Then
        Timer1.Interval = 500
        Mode = 5
    End If
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode6.ltmp") <> "" Then
        Timer1.Interval = 1000
        Mode = 6
    End If
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Mode7.ltmp") <> "" Then
        Timer1.Interval = 1500
        Mode = 7
    End If
    If Dir("C:\Users\" & Environ("USERNAME") & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Start.ltmp") <> "" Then
        Kill ("C:\Users\" & Environ("USERNAME") & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\Start.ltmp")
    End If
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Timer1.Enabled = True
Else
    Timer1.Enabled = False
End If
Exit Sub
Errt:
Shell_NotifyIcon NIM_DELETE, tray
Call MsgBox("Sta_F1_Loa£º" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
'End
End Sub
Private Sub Form_Unload(Cancel As Integer)
KillerStart = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If KillerStart = False Then
    'GoTo CleanIcon
    Exit Sub
End If
LoadAll
If Mode = 7 Then
    X = 0
    Do While X <= 1023
        If Dir(PopsURL(X)) <> "" Then
            Kill (PopsURL(X))
        End If
        X = X + 1
        DoEvents
    Loop
Else
    L = 0
    X = 0
    Do While X <= 1023
        r = KillerMainModble.Pops(X)
        If r <> "" Then
            L = L + 1
        End If
        X = X + 1
        DoEvents
    Loop
    X = 0
    Dim p
    Do While X <= L
        s = KillerMainModble.Pops(X)
        Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='" & s & "'")
        For Each p In colProcessList
            p.Terminate
        Next
            
        DoEvents
        X = X + 1
    Loop
    If Mode = 1 Then
        X = 0
        Do While X <= 1023
            If Dir(PopsURL(X)) <> "" Then
                Kill (PopsURL(X))
            End If
            X = X + 1
            DoEvents
        Loop
    End If
End If

Exit Sub
CleanIcon:
If Dir("C:\Users\" & Environ("USERNAME") & "\AppData\Roaming\LJXPopupKiller\Inf\Temp\EndAll.ltmp") <> "" Then
    Shell_NotifyIcon NIM_DELETE, tray
    'Call EndProgram
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
msg = X / 15
'msgÈ¡Öµ
'410Êó±êÔÚÍÐÅÌÉÏ
'411×ó¼üµ¥»÷ÍÐÅÌ
'413ÓÒ¼üµã»÷ÍÐÅÌ
'415ÖÐ¼üµã»÷ÍÐÅÌ
If msg = 411 Or msg = 514 Then
    On Error Resume Next
    Form1.Show
    
End If
End Sub
Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, tray
Call EndProgram
End Sub

Public Function CI()
Shell_NotifyIcon NIM_DELETE, tray
End Function
