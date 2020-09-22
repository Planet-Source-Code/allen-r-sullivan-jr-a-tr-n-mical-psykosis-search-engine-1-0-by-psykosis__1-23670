Attribute VB_Name = "Module1"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HWND_TOPMOST = -1
Public Const WM_MOVE = &HF012
Public Const WM_SYSCOMMAND = &H112
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE





Public Sub FormOnTop(frm As Form)
    Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
End Sub



Public Sub FormDrag(frm As Form)
    'self explanatory
    Call ReleaseCapture
    Call SendMessage(frm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
