VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "POS: by Chadworkz.com"
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmMain"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   2790
   ScaleWidth      =   4590
   Begin VB.Timer tmrPOS 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2880
      Top             =   1260
   End
   Begin VB.Label lblWebsite 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WWW.CHADWORKZ.COM"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003C3C3C&
      Height          =   120
      Left            =   1140
      TabIndex        =   7
      ToolTipText     =   " Click to visit: www.chadworkz.com "
      Top             =   2625
      Width           =   1275
   End
   Begin VB.Shape shapeLight 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006E6E6E&
      Height          =   90
      Left            =   3450
      Shape           =   5  'Rounded Square
      Top             =   2535
      Width           =   90
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H0012129E&
      Height          =   165
      Left            =   4350
      TabIndex        =   6
      Top             =   345
      Width           =   105
   End
   Begin VB.Image btnEnable 
      Height          =   240
      Left            =   3630
      Picture         =   "frmMain.frx":29CB2
      Top             =   2460
      Width           =   840
   End
   Begin VB.Image imgDisableDown 
      Height          =   240
      Left            =   840
      Picture         =   "frmMain.frx":2A774
      Top             =   3270
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgDisableHover 
      Height          =   240
      Left            =   840
      Picture         =   "frmMain.frx":2B236
      Top             =   3030
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgDisableUp 
      Height          =   240
      Left            =   840
      Picture         =   "frmMain.frx":2BCF8
      Top             =   2790
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgEnableDown 
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":2C7BA
      Top             =   3270
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgEnableHover 
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":2D27C
      Top             =   3030
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgEnableUp 
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":2DD3E
      Top             =   2790
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblOffsetY 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4005
      TabIndex        =   5
      Top             =   1875
      Width           =   105
   End
   Begin VB.Label lblOffsetX 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4005
      TabIndex        =   4
      Top             =   1665
      Width           =   105
   End
   Begin VB.Label lblFormY 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4005
      TabIndex        =   3
      Top             =   1455
      Width           =   105
   End
   Begin VB.Label lblFormX 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4005
      TabIndex        =   2
      Top             =   1245
      Width           =   105
   End
   Begin VB.Label lblCursorY 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4005
      TabIndex        =   1
      Top             =   1035
      Width           =   105
   End
   Begin VB.Label lblCursorX 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4005
      TabIndex        =   0
      Top             =   825
      Width           =   105
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim blnEnabled As Boolean, blnHovering As Boolean

Private Sub btnEnable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If blnEnabled = True Then
        btnEnable.Picture = imgDisableDown.Picture
    Else
        btnEnable.Picture = imgEnableDown.Picture
    End If
    
    Call PlayWav(App.path & "\sounds\click.wav")
End Sub

Private Sub btnEnable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    blnHovering = True
    
    If blnEnabled = True Then
        btnEnable.Picture = imgDisableHover.Picture
    Else
        btnEnable.Picture = imgEnableHover.Picture
    End If
End Sub

Private Sub btnEnable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If blnEnabled = True Then
        If blnHovering = True Then
            btnEnable.Picture = imgDisableHover.Picture
        Else
            btnEnable.Picture = imgDisableUp.Picture
        End If
        blnEnabled = False
        tmrPOS.Enabled = False
        shapeLight.BackColor = &HC0C0C0
        lblCursorX.Caption = "0"
        lblCursorY.Caption = "0"
        lblFormX.Caption = "0"
        lblFormY.Caption = "0"
        lblOffsetX.Caption = "0"
        lblOffsetY.Caption = "0"
    Else
        If blnHovering = True Then
            btnEnable.Picture = imgEnableHover.Picture
        Else
            btnEnable.Picture = imgEnableUp.Picture
        End If
        blnEnabled = True
        tmrPOS.Enabled = True
        shapeLight.BackColor = &H5858C6
    End If
End Sub

Private Sub Form_Load()
    
    Dim lngLoads As Long

    Call CenterForm(Me)
    Call FormOnTop(Me, True)
    Call TransparentByColor(Me, 65280)
    Call TransparentByPercent(Me.hwnd, 80&)
    
    blnEnabled = False
    blnHovering = False
    
    If DirExists(App.path & "\data") = False Then Call MkDir(App.path & "\data")
    If DirExists(App.path & "\sounds") = False Then
        Call MkDir(App.path & "\sounds")
        Name App.path & "\click.wav" As App.path & "\sounds\click.wav"
    End If
    
    If FileExists(App.path & "\data\pos.txt") = True Then
        lngLoads& = CLng(ReadData("POS - Total Times Loaded ", "Total Loads: ", App.path & "\data\pos.txt"))
        lngLoads& = lngLoads& + 1&
        Call WriteData("POS - Total Times Loaded ", "Total Loads: ", CStr(lngLoads&), App.path & "\data\pos.txt")
        Call MsgBox("Number of times loaded [" & lngLoads& & "].", vbInformation + vbOKOnly, "Loads")
    Else
        lngLoads& = 0&
        Call WriteData("POS - Date of First Load ", "Loaded Date: ", Date, App.path & "\data\pos.txt")
        Call WriteData("POS - Time of First Load ", "Loaded Time: ", Time, App.path & "\data\pos.txt")
        Call WriteData("POS - Total Times Loaded ", "Total Loads: ", CStr(lngLoads&), App.path & "\data\pos.txt")
        Call WriteData("POS - Last Time Unloaded ", "Unloaded Time: ", Now, App.path & "\data\pos.txt")
        Call MsgBox("Hello! Welcome to another Chadworkz™ Example!" & vbCrLf & vbCrLf & _
        "Please visit [ http://www.chadworkz.com ] for more examples.", vbInformation + vbOKOnly, "First Load")
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call FormDrag(Me, True)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    blnHovering = False
    
    If blnEnabled = True Then
        btnEnable.Picture = imgDisableUp.Picture
    Else
        btnEnable.Picture = imgEnableUp.Picture
    End If
    
    lblWebsite.ForeColor = &H3C3C3C
    lblWebsite.FontUnderline = False
    lblExit.ForeColor = &H12129E
    lblExit.FontBold = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call Pause(0.15)
    End
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblExit.Top = 360
    lblExit.Left = 4365
    
    Call PlayWav(App.path & "\sounds\click.wav")
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblExit.ForeColor = &H0
    lblExit.FontBold = True
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblExit.Top = 345
    lblExit.Left = 4350
    
    Call Unload(Me)
End Sub

Private Sub lblWebsite_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblWebsite.Top = 2640
    lblWebsite.Left = 1155
    
    Call PlayWav(App.path & "\sounds\click.wav")
End Sub

Private Sub lblWebsite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblWebsite.ForeColor = &H0&
    'lblWebsite.ForeColor = &HFFFFFF
    lblWebsite.FontUnderline = True
End Sub

Private Sub lblWebsite_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblWebsite.Top = 2625
    lblWebsite.Left = 1140
    
    Call ShellExecute(0&, vbNullString, "http://www.chadworkz.com", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub tmrPOS_Timer()

    Dim lngCursorTop As Long, lngFormTop As Long, lngOffsetTop As Long
    Dim lngCursorLeft As Long, lngFormLeft As Long, lngOffsetLeft As Long
    
    lngCursorTop& = CursorY&
    lngCursorLeft& = CursorX&
    
    lngFormTop& = Me.Top / 15&
    lngFormLeft& = Me.Left / 15&
    
    lngOffsetTop& = (lngCursorTop& - lngFormTop&) * 15
    lngOffsetLeft& = (lngCursorLeft& - lngFormLeft&) * 15
    
    lblCursorY.Caption = Round(lngCursorTop&, 0&)
    lblCursorX.Caption = Round(lngCursorLeft&, 0&)
    
    lblFormY.Caption = Round(lngFormTop&, 0&)
    lblFormX.Caption = Round(lngFormLeft&, 0&)
    
    lblOffsetY.Caption = Round(lngOffsetTop&, 0&)
    lblOffsetX.Caption = Round(lngOffsetLeft&, 0&)
End Sub
