Attribute VB_Name = "KillerMainModble"

Public Mode As Byte             '��ǰģʽ
Public PopsURL(0 To 1023) As String     '�����ļ�·��
Public Pops(0 To 1023) As String        '�����ļ�����
Public MaxPops As Long      '���ĵ����ļ�
Function loadLang()
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


If aN = 0 Then
    HasOtherLanguage = False
Else
    HasOtherLanguage = True
End If
End Function
Function LoadSetting()

End Function
