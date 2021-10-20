VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test Selenium Browser"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Regist tlb"
      Height          =   255
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "https://www.google.com/"
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Edge"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Chrome"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin Proyecto1.MaatoohBrowser MaatoohBrowser1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9495
      _extentx        =   16748
      _extenty        =   12515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MBrowser As New WebDriver
Private Sub Command1_Click()
MBrowser.Quit
MBrowser.Start ("Chrome")
'Dock Window (Optional)
MaatoohBrowser1.DockHwnd (GetHwnd.GetHwndMaatoohBrowser("data:,"))
'---------------------
MBrowser.Get (Text1)
End Sub

Private Sub Command2_Click()
MBrowser.Quit
MBrowser.Start ("Edge")
'Dock Window (Optional)
MaatoohBrowser1.DockHwnd (GetHwnd.GetHwndMaatoohBrowser("data:,"))
'---------------------
MBrowser.Get (Text1)
End Sub

Private Sub RegisterLibrary()

Shell "cmd.exe /k cd " & App.Path & "&&" & _
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase MaatoohBrws.dll /tlb:MaatoohBrws64.tlb" & "&&" & _
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe /codebase MaatoohBrws.dll /tlb:MaatoohBrws32.tlb" & "&&" & _
"PAUSE" & "&&" & "CLS"
End Sub

Private Sub Command3_Click()
MsgBox "OPEN WITH ADMINISTRATOR"
Call RegisterLibrary
End Sub

Private Sub Form_Resize()
If Not Form1.WindowState = vbMinimized Then
MaatoohBrowser1.Height = Form1.Height * 0.98
MaatoohBrowser1.Width = Form1.Width * 0.98
MaatoohBrowser1.DockHwnd (GetHwnd.GHWND)
DoEvents
End If
End Sub


