Attribute VB_Name = "Module1"

'=================================== Modual 2 ================================================'

Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwnd1 As Long, ByVal hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub PushKeys Lib "KeyPush" (ByVal Keystrokes As String)
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWindow As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_TAB = &H9
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11
Public Const VK_RBUTTON = &H2
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const CB_SELECTSTRING = &H14D
Public Const CB_GETCOUNT = &H146
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_SPACE = &H20
Public Const WM_DESTROY = &H2
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const BM_SETCHECK = &HF1
Public Const BST_CHECKED = &O1
Public Const BST_INDETERMINATE = &O2
Public Const BST_UNCHECKED = &O0
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const HWND_TOPMOST = -1
Public nid As NOTIFYICONDATA

Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type

Public Sub Save_ComboBox(Path As String, Combo As ComboBox)
'Ex: Call Save_ComboBox("c:\windows\desktop\combo.cmb", combo1)

    Dim Saves As Long
    On Error Resume Next

    Open Path$ For Output As #1
    For Saves& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(Saves&)
    Next Saves&
    Close #1
End Sub

Public Sub Load_ComboBox(Path As String, Combo As ComboBox)
'Call Load_ComboBox("c:\windows\desktop\combo.cmb", Combo1)

    Dim What As String
    On Error Resume Next
    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, What$
        DoEvents
        Combo.AddItem What$
    Wend
    Close #1
End Sub
Public Sub Save_Text(Txt As TextBox, FilePath As String)
'Ex: Call Save_Text(list1,"c:\windows\desktop\text.txt")
    
    Open FilePath$ For Output As #1
        Print #1, Txt
    Close 1
End Sub

Public Sub Load_Text(Txt As TextBox, FilePath As String)
'Ex: Call load_Text(list1,"c:\windows\desktop\text.txt")

    Dim mystr As String, textz As String, a As String
    
    Open FilePath$ For Input As #1
    Do While Not EOF(1)
    Line Input #1, a$
        textz$ = textz$ + a$ + Chr$(13) + Chr$(10)
        Loop
        Txt = textz$
    Close #1
End Sub
'======================================================End of Modual 2=====================

