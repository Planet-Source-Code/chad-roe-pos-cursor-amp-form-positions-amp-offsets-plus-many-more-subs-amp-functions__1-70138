Attribute VB_Name = "modTransparency"
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

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const RGN_DIFF = 4
Public Const RGN_COPY = 5
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const WS_EX_LAYERED = &H80000

Public Function TransparentByColor(ByRef frmForm As Form, ByVal lngColor As Long) As Long
    
    Dim lngScale As Long, lngRGB As Long
    Dim lngWidth As Long, lngHeight As Long, lngX As Long, lngY As Long
    Dim rgnMain As Long, rgnPixel As Long, bmpMain As Long, dcMain As Long
    
    lngScale& = frmForm.ScaleMode
    frmForm.ScaleMode = 3
    frmForm.BorderStyle = 0
    lngWidth& = frmForm.ScaleX(frmForm.Picture.Width, vbHimetric, vbPixels)
    lngHeight& = frmForm.ScaleY(frmForm.Picture.Height, vbHimetric, vbPixels)
    frmForm.Width = lngWidth& * Screen.TwipsPerPixelX
    frmForm.Height = lngHeight& * Screen.TwipsPerPixelY
    rgnMain& = CreateRectRgn(0&, 0&, lngWidth&, lngHeight&)
    dcMain& = CreateCompatibleDC(frmForm.hDC)
    bmpMain& = SelectObject(dcMain&, frmForm.Picture.handle)
    
    For lngY& = 0& To lngHeight&
        For lngX& = 0& To lngWidth&
            lngRGB& = GetPixel(dcMain&, lngX&, lngY&)
            If lngRGB& = lngColor& Then
                rgnPixel& = CreateRectRgn(lngX&, lngY&, lngX& + 1&, lngY& + 1&)
                Call CombineRgn(rgnMain&, rgnMain&, rgnPixel&, RGN_XOR)
                Call DeleteObject(rgnPixel&)
            End If
        Next lngX&
    Next lngY&

    Call SelectObject(dcMain&, bmpMain&)
    Call DeleteDC(dcMain&)
    Call DeleteObject(bmpMain&)

    If rgnMain& <> 0& Then
        Call SetWindowRgn(frmForm.hwnd, rgnMain&, True)
        TransparentByColor& = rgnMain&
    End If

    frmForm.ScaleMode = lngScale&
End Function

Public Function TransparentByPercent(ByVal hwndWindow As Long, lngPercent As Long) As Long
    
    Dim lngMessage As Long
    
    On Error Resume Next
    If lngPercent& < 0& Or lngPercent& > 100& Then
        TransparentByPercent& = 0&
    Else
        lngMessage& = GetWindowLong(hwndWindow&, GWL_EXSTYLE)
        lngMessage& = lngMessage& Or WS_EX_LAYERED
        Call SetWindowLong(hwndWindow&, GWL_EXSTYLE, lngMessage&)
        Call SetLayeredWindowAttributes(hwndWindow&, 0&, lngPercent& * 2.55, LWA_ALPHA)
        TransparentByPercent& = lngPercent&
    End If
    
    If Err Then TransparentByPercent& = 0&
End Function
