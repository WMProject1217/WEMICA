VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form"
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9840
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   9840
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   233
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":0CCA
      Top             =   720
      Width           =   9375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "∥"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "一"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "↖"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   8400
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Form"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type BLENDFUNCTION
BlendOp As Byte
BlendFlags As Byte
SourceConstantAlpha As Byte
AlphaFormat As Byte
End Type
Private Const AC_SRC_OVER = &H0
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public ndis
Private Sub Form_Load()
Label1.BackColor = RGB(102, 204, 255)
Label1.ForeColor = RGB(255, 255, 255)
Label2.BackColor = RGB(237, 0, 0)
Label2.ForeColor = RGB(255, 255, 255)
Label3.BackColor = RGB(102, 204, 255)
Label4.BackColor = RGB(102, 204, 255)
Label4.ForeColor = RGB(255, 255, 255)
Label5.BackColor = RGB(102, 204, 255)
Label5.ForeColor = RGB(255, 255, 255)
Text1.BackColor = RGB(0, 0, 0)
Text1.ForeColor = RGB(255, 255, 255)
Me.BackColor = RGB(0, 0, 0)
On Error GoTo dmd
Dim a As String, b() As String
a = Command
b = Split(a, "=")
Label1.Caption = b(0)
Me.Caption = b(0)
Text1.Text = b(1)
Exit Sub
dmd:
Me.Caption = "Error"
Label1.Caption = "Error"
Text1.Text = "404 Not Found"
End Sub

Private Sub Form_Resize()
Label1.Width = Me.Width
Label1.Left = 0
Label1.Top = 0
Label1.Height = 500
Label2.Width = 735
Label2.Left = Me.Width - Label2.Width '- 200
Label2.Top = 0
Label2.Height = 500
Label3.Width = 735
Label3.Left = Label2.Left - Label3.Width - 30
Label3.Top = 0
Label3.Height = 500
Label4.Width = 735
Label4.Left = Label3.Left - Label4.Width - 30
Label4.Top = 0
Label4.Height = 500
Label5.Width = 735
Label5.Left = Label4.Left - Label5.Width - 30
Label5.Top = 0
Label5.Height = 500
End Sub



Private Sub label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Button
End Sub
Sub MoveForm(hButton As Integer)
If hButton = vbLeftButton Then
ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0&
End If
End Sub

Private Sub Label2_Click()
End
End Sub



Private Sub Label4_Click()
Me.WindowState = 1
End Sub

Private Sub Label5_Click()
If ndis = 1 Then
SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Label5.Caption = "∥"
ndis = 0
Else
SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Label5.Caption = "⊥"
ndis = 1
End If
End Sub

Private Sub Label6_Click()
End
End Sub

