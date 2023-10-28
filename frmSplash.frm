VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2664
   ClientLeft      =   252
   ClientTop       =   1416
   ClientWidth     =   7356
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2664
   ScaleWidth      =   7356
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2472
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "[2023.10.21]v2.0.7[免激活]"
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   3012
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "赖靖轩 版权所有"
         Height          =   252
         Left            =   5520
         TabIndex        =   3
         Top             =   240
         Width           =   1452
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   6960
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line3 
         X1              =   6960
         X2              =   6960
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   120
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6960
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "加载中。。。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   16.2
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   6852
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LJX弹窗杀手"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6852
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Height          =   372
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   3420
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "正在加载主界面"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   6852
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Label4.Caption = "[2023.10.21]v" & App.Major & "." & App.Minor & "." & App.Revision & "[免激活]"
Label5.Width = 20
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

