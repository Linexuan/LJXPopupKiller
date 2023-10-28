VERSION 5.00
Begin VB.Form InfoFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   3132
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3132
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "InfoFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If HavAdminPer = True Then
    Me.Caption = " [Admin] " & Me.Caption
End If
End Sub
