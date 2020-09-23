Attribute VB_Name = "Module2"
'=======================================================================================
                                 'Modual 1
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwnd1 As Long, ByVal hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long

Public Function GetCaption(ByVal lhWnd As Long) As String
    Dim sA As String, lLen As Long
    
    lLen& = GetWindowTextLength(lhWnd&)
    sA$ = String(lLen&, 0&)
    Call GetWindowText(lhWnd&, sA$, lLen& + 1)
    GetCaption$ = sA$
End Function


Public Function FindAnyWindow(frm As Form, ByVal WinTitle As String, Optional ByVal CaseSensitive As Boolean = False) As Long
    Dim lhWnd As Long, sA As String
    lhWnd& = frm.hwnd


    Do Until lhWnd& = 0


        DoEvents
            
            sA$ = GetCaption(lhWnd&)
            If InStr(IIf(CaseSensitive = False, LCase$(sA$), sA$), IIf(CaseSensitive = False, LCase$(WinTitle$), WinTitle$)) Then FindAnyWindow& = lhWnd&: Exit Do Else FindAnyWindow& = 0
            
            lhWnd& = GetNextWindow(lhWnd&, 2)
        Loop
    End Function

