VERSION 5.00
Begin VB.Form BSP2ProjectForm1 
   BorderStyle     =   0  'None
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "BSP2ProjectForm1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "BSP2ProjectForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const WS_EX_TRANSPARENT As Long = &H20&

Private Sub Form_Activate()
Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
Label1.Width = Me.Width
Label1.Left = Me.Width
Timer1.Interval = 20
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
SetWindowLong Me.hwnd, -20, GetWindowLong(Me.hwnd, -20) Or &H80000
SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
SetLayeredWindowAttributes Me.hwnd, vbRed, 0, 1
Me.BackColor = vbRed
Label1.BackColor = vbRed
On Error GoTo error
Dim cmd
Dim a As String, b() As String
cmd = Command
If Command = "" Then
Label1.ForeColor = RGB(102, 204, 255)
Label1.Caption = "这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。这只是一个测试。"
GoTo error
End If
a = Command
b = Split(a, "=")
On Error GoTo clrn
Label1.Top = b(0) * Label1.Height
Label1.Caption = b(1)
If b(2) = 255 And b(3) = 0 And b(4) = 0 Then
b(2) = 254
End If
Label1.ForeColor = RGB(b(2), b(3), b(4))
Exit Sub
Label1.Caption = b(1)
Exit Sub
clrn:
Label1.Caption = a
error:
End Sub
Private Sub Timer1_Timer()
Label1.Left = Label1.Left - 50
If Label1.Left + Label1.Width < 0 Then
End
End If
End Sub
