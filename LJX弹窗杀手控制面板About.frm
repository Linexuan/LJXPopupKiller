VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于LJX弹窗杀手"
   ClientHeight    =   4740
   ClientLeft      =   2340
   ClientTop       =   1932
   ClientWidth     =   5664
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3271.632
   ScaleMode       =   0  'User
   ScaleWidth      =   5313.225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   432
      Left            =   240
      Picture         =   "LJX弹窗杀手控制面板About.frx":0000
      ScaleHeight     =   263.118
      ScaleMode       =   0  'User
      ScaleWidth      =   263.118
      TabIndex        =   1
      Top             =   120
      Width           =   432
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFC0&
      Cancel          =   -1  'True
      Caption         =   "关闭"
      Default         =   -1  'True
      Height          =   345
      Left            =   4125
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      BackColor       =   &H00FFFFC0&
      Caption         =   "系统信息(&S)..."
      Height          =   345
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1485
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "本次主要更新内容："
      Height          =   1332
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   3972
   End
   Begin VB.Line Line2 
      X1              =   112.568
      X2              =   5183.771
      Y1              =   1987.827
      Y2              =   1987.827
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "赖靖轩（LJX）版权所有，保留所有权利"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5338.553
      Y1              =   1573.696
      Y2              =   1573.696
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"LJX弹窗杀手控制面板About.frx":030A
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1176
      Left            =   1080
      TabIndex        =   3
      Top             =   1008
      Width           =   3888
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "应用程序标题"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "版本"
      Height          =   225
      Left            =   1080
      TabIndex        =   6
      Top             =   660
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "警告:LJX弹窗杀手可能会使一些应用无法运行或导致应用运行时错误。当你使用LJX弹窗杀手时，所导致的损失由使用者自负。"
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   5415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字 ROOT 类型...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' 独立的空的终结字符串
Const REG_DWORD = 4                      ' 32位数字

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "关于 " & App.Title
    lblVersion.Caption = "软件版本v" & App.Major & "." & App.Minor & "." & App.Revision & "[免激活]"
    lblTitle.Caption = App.Title
    Form1.Enabled = False
    '注意修改frmSplash上的版本数据！
    Label2.Caption = "[2023.10.21]v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
                    "主要更新内容：" & vbCrLf & _
                    "1、改善了管理员权限的检测逻辑" & vbCrLf & _
                    "2、修复部分激活漏洞" & vbCrLf & _
                    "3、修复了启动画面的显示漏洞" & vbCrLf & _
                    "4、修复了部分情况下从托盘呼出主界面时无法显示的问题"
                    
    If HavAdminPer = True Then
        Me.Caption = " [Admin] " & Me.Caption
    End If
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' 试图从注册表中获得系统信息程序的路径及名称...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' 试图仅从注册表中获得系统信息程序的路径...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' 已知32位文件版本的有效位置
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' 错误 - 文件不能被找到...
        Else
            GoTo SysInfoErr
        End If
    ' 错误 - 注册表相应条目不能被找到...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    Call TryReOpen
End Sub
Private Function TryReOpen()
On Error GoTo Errt
Shell ("MSINFO32.EXE")
Exit Function
Errt:
    Call MsgBox("系统信息调用失败：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Function

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' 循环计数器
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 打开的注册表关键字句柄
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' 注册表关键字数据类型
    Dim tmpVal As String                                    ' 注册表关键字值的临时存储器
    Dim KeyValSize As Long                                  ' 注册表关键自变量的尺寸
    '------------------------------------------------------------
    ' 打开 {HKEY_LOCAL_MACHINE...} 下的 RegKey
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表关键字
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量尺寸
    
    '------------------------------------------------------------
    ' 检索注册表关键字的值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' 获得/创建关键字值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 外接程序空终结字符串...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null 被找到,从字符串中分离出来
    Else                                                    ' WinNT 没有空终结字符串...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null 没有被找到, 分离字符串
    End If
    '------------------------------------------------------------
    ' 决定转换的关键字的值类型...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' 搜索数据类型...
    Case REG_SZ                                             ' 字符串注册关键字数据类型
        KeyVal = tmpVal                                     ' 复制字符串的值
    Case REG_DWORD                                          ' 四字节的注册表关键字数据类型
        For i = Len(tmpVal) To 1 Step -1                    ' 将每位进行转换
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 生成值字符。 By Char。
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' 转换四字节的字符为字符串
    End Select
    
    GetKeyValue = True                                      ' 返回成功
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
    Exit Function                                           ' 退出
    
GetKeyError:      ' 错误发生后将其清除...
    KeyVal = ""                                             ' 设置返回值到空字符串
    GetKeyValue = False                                     ' 返回失败
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
End Function

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub

