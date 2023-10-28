Attribute VB_Name = "KillerMainModble"

Public Mode As Byte             '当前模式
Public PopsURL(0 To 1023) As String     '弹窗文件路径
Public Pops(0 To 1023) As String        '弹窗文件进程
Public MaxPops As Long      '最多的弹窗文件
Function loadLang()
Dim aN
aN = 0
Dim tmpTxt As String
Dim n

If Dir(App.Path & "/中文.ini") <> "" Then
    Open (App.Path & "/中文.ini") For Input As #10
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
    langNames(aN - 1) = "中文"
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
If Dir(App.Path & "/русский язык.ini") <> "" Then
    Open (App.Path & "/русский язык.ini") For Input As #10
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
    langNames(aN - 1) = "русский язык"
    Close #10
End If
If Dir(App.Path & "/日本語.ini") <> "" Then
    Open (App.Path & "/日本語.ini") For Input As #10
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
    langNames(aN - 1) = "日本語"
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
