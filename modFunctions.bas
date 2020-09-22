Attribute VB_Name = "modFunctions"
Option Explicit
'_________________________________________________________________________________________________'
'@ H @ T @ T @ P @ : @ / @ / @ W @ W @ W @ . @ C @ H @ A @ D @ W @ O @ R @ K @ Z @ . @ C @ O @ M @'
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'    ..######..##.....##....###....########..##......##..#######..########..##....##.######## TM  '
'    .##....##.##.....##...##.##...##.....##.##..##..##.##.....##.##.....##.##...##.......##.     '
'    .##.......##.....##..##...##..##.....##.##..##..##.##.....##.##.....##.##..##.......##..     '
'    .##.......#########.##.....##.##.....##.##..##..##.##.....##.########..#####.......##...     '
'    .##.......##.....##.#########.##.....##.##..##..##.##.....##.##...##...##..##.....##....     '
'    .##....##.##.....##.##.....##.##.....##.##..##..##.##.....##.##....##..##...##...##.....     '
'    ..######..##.....##.##.....##.########...###..###...#######..##.....##.##....##.########.COM '
'_________________________________________________________________________________________________'
'@ H @ T @ T @ P @ : @ / @ / @ W @ W @ W @ . @ C @ H @ A @ D @ W @ O @ R @ K @ Z @ . @ C @ O @ M @'
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'   This example is part of my Chadworkz™ Example Series - http://www.chadworkz.com/vb/examples   '
'_________________________________________________________________________________________________'

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const WM_MOVE = &HF012
Public Const WM_SYSCOMMAND = &H112
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Function CursorX() As Long

    Dim pt As POINTAPI
    
    Call GetCursorPos(pt)
    CursorX& = pt.X
End Function

Public Function CursorY() As Long

    Dim pt As POINTAPI
    
    Call GetCursorPos(pt)
    CursorY& = pt.Y
End Function

Public Sub FormOnTop(frmForm As Form, blnOnTop As Boolean)
    
    If blnOnTop = True Then
        Call SetWindowPos(frmForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
    Else
        Call SetWindowPos(frmForm.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
    End If
End Sub

Public Sub FormDrag(frmForm As Form, blnXP As Boolean)
    
    Call ReleaseCapture
    If blnXP = True Then
        Call SendMessage(frmForm.hwnd, &HA1, 2, 0&)
    Else
        Call SendMessage(frmForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
    End If
End Sub

Public Sub CenterForm(frmForm As Form)
    
    frmForm.Left = Screen.Width / 2 - frmForm.Width / 2
    frmForm.Top = Screen.Height / 2 - frmForm.Height / 2
End Sub

Public Sub Pause(Length As Double)

    Dim starttime
    
    starttime = Timer
    Do While Timer - starttime <= Length
        DoEvents
    Loop
End Sub

Public Sub PlayWav(strPath As String)
    
    Dim intFlags As Integer
    
    intFlags% = SND_ASYNC Or SND_NODEFAULT
    Call sndPlaySound(strPath$, intFlags%)
End Sub

Public Function ReadData(strSection As String, strKey As String, strFile As String) As String
    
    Dim strBuffer As String
    
    strBuffer$ = String(750&, Chr(0))
    strKey$ = LCase$(strKey$)
    ReadData$ = Left(strBuffer$, GetPrivateProfileString(strSection$, ByVal strKey$, "", strBuffer$, Len(strBuffer$), strFile$))
End Function

Public Function WriteData(strSection As String, strKey As String, strValue As String, strFile As String) As String
    
    WriteData$ = strSection$ & ":" & strKey$ & ":" & strValue$ & ":" & strFile$
    Call WritePrivateProfileString(strSection$, UCase$(strKey$), strValue$, strFile$)
End Function

Public Function FileExists(strFile As String) As Boolean

    If Len(Dir$(strFile$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Public Function DirExists(ByVal strDirectory As String) As Boolean
    
    Dim strContents As String, lngResult As Long
    
    If Mid(strDirectory$, Len(strDirectory$) - 1, 1) <> "\" Then
        strDirectory$ = strDirectory$ & "\"
    End If
    
    strContents$ = Dir(strDirectory$ & "*.*", vbDirectory)
    lngResult& = Not (strContents$ = vbNullString)
    
    If lngResult& = -1 Then
        DirExists = True
    Else
        DirExists = False
    End If
End Function
