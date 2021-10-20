VERSION 5.00
Begin VB.UserControl MaatoohBrowser 
   AutoRedraw      =   -1  'True
   ClientHeight    =   7080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   ScaleHeight     =   7080
   ScaleWidth      =   9495
   Begin VB.PictureBox Browser 
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7035
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "MaatoohBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objDer As PictureBox
Private hndNotepad As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function SetParent Lib "user32" _
    (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Enum eShowWindow
    HIDE_eSW = 0&
    SHOWNORMAL_eSW = 1&
    NORMAL_eSW = 1&
    SHOWMINIMIZED_eSW = 2&
    SHOWMAXIMIZED_eSW = 3&
    MAXIMIZE_eSW = 3&
    SHOWNOACTIVATE_eSW = 4&
    SHOW_eSW = 5&
    MINIMIZE_eSW = 6&
    SHOWMINNOACTIVE_eSW = 7&
    SHOWNA_eSW = 8&
    RESTORE_eSW = 9&
    SHOWDEFAULT_eSW = 10&
    MAX_eSW = 10&
End Enum

Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal nCmdShow As eShowWindow) As Long

Private Declare Function MoveWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function IsChild Lib "user32" _
    (ByVal hWndParent As Long, ByVal hwnd As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECTAPI
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type WINDOWPLACEMENT
    Length As Long
    Flags As Long
    ShowCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECTAPI
End Type

Private Declare Function GetWindowPlacement Lib "user32" _
    (ByVal hwnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long


Private Sub dockForm(ByVal formhWnd As Long, _
                     ByVal picDock As PictureBox, _
                     Optional ByVal ajustar As Boolean = True)

    Call SetParent(formhWnd, picDock.hwnd)
    posDockForm formhWnd, picDock, ajustar
    Call ShowWindow(formhWnd, NORMAL_eSW)
End Sub


Private Sub posDockForm(ByVal formhWnd As Long, _
                        ByVal picDock As PictureBox, _
                        Optional ByVal ajustar As Boolean = True)

    Dim nWidth As Long, nHeight As Long
    Dim wndPl As WINDOWPLACEMENT
    '
    If ajustar Then
        nWidth = picDock.ScaleWidth \ Screen.TwipsPerPixelX
        nHeight = picDock.ScaleHeight \ Screen.TwipsPerPixelY
    Else

        Call GetWindowPlacement(formhWnd, wndPl)
        With wndPl.rcNormalPosition
            nWidth = .Right - .Left
            nHeight = .Bottom - .Top
        End With
    End If
    Call MoveWindow(formhWnd, -8, -120, nWidth + 16, nHeight + 128, True)
End Sub

Private Sub UserControl_Initialize()
Set objDer = Browser

End Sub

Public Sub DockHwnd(HwndPath)
hndNotepad = HwndPath
dockForm hndNotepad, objDer, True
End Sub

Private Sub UserControl_Resize()
Browser.Width = UserControl.Width
Browser.Height = UserControl.Height
End Sub
