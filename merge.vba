Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal wClassName As Any, ByVal wWindowName As String) As Long


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
ByVal lpsz2 As String) As Long

Public BM_CLICK


Sub Start()
Dim wd
BM_CLICK = &HF5&

 For Each wd In CreateObject("Shell.Application").Windows
            'MsgBox wd.LocationName
            If InStr(CStr(wd.LocationName), "PestPac") > 0 Then
                'MsgBox wd.document.Title + "lol"
                'SendKeys "%(s)"
                If Not wd.document.getElementsByName("butUpdate")(0) Is Nothing Then
                    AppActivate wd
                    SendKeys "%(u)"
                    Sleep 1000
                    SendKeys "y"
                    MsgBox wd.document.Title + "lol"
                    'AppActivate wd
                    'SendKeys "y"
                    'SendKeys "y"
                    'wd.document.getElementsByName("butUpdate")(0).Click
                    'MsgBox wd.document.Title + "lol"
                    Dim hWND As Long, childHWND As Long
                   
                   While wd.Busy Or wd.readyState <> READYSTATE_COMPLETE
            DoEvents
            If wd.Busy Then
                DoEvents
                hWND = FindWindow(vbNullString, "POPPP!!")
                If hWND <> 0 Then childHWND = FindWindowEx(hWND, ByVal 0&, "Button", "OK")
                If childHWND <> 0 Then SendMessage childHWND, BM_CLICK, 0, 0
                MsgBox wd.document.Title + "lol"
            End If
    Wend


                End If
            End If
            Next wd
End Sub




