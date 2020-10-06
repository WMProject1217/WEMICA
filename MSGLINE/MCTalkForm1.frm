VERSION 5.00
Begin VB.Form MCTalkForm1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13635
   Icon            =   "MCTalkForm1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   13635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Left            =   12960
      Top             =   6720
   End
   Begin VB.Timer Timer1 
      Left            =   12960
      Top             =   7200
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
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
      Left            =   720
      TabIndex        =   7
      Top             =   6960
      Width           =   3975
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
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
      Left            =   720
      TabIndex        =   6
      Top             =   6360
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
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
      Left            =   720
      TabIndex        =   5
      Top             =   5760
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
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
      Left            =   720
      TabIndex        =   4
      Top             =   5160
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
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
      Left            =   720
      TabIndex        =   3
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
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
      Left            =   720
      TabIndex        =   2
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Left            =   720
      TabIndex        =   1
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
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
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   3975
   End
End
Attribute VB_Name = "MCTalkForm1"
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
Public ad001
Public ad002
Public ad003
Public ad004
Public ad005
Public ad006
Public ad007
Public ad008
Public ad009
Public ad010
Public ad011
Public ad012
Public ad013
Public ad014
Public ad015
Public ad016
Public adchk
Public adlst
Public amkl
Public GTX
Public Sub AMLT()
On Error GoTo error
If ad016 <> "" Then
amkl = 16
GoTo runamltn
End If
If ad015 <> "" Then
amkl = 15
GoTo runamltn
End If
If ad014 <> "" Then
amkl = 14
GoTo runamltn
End If
If ad013 <> "" Then
amkl = 13
GoTo runamltn
End If
If ad012 <> "" Then
amkl = 12
GoTo runamltn
End If
If ad011 <> "" Then
amkl = 11
GoTo runamltn
End If
If ad010 <> "" Then
amkl = 10
GoTo runamltn
End If
If ad009 <> "" Then
amkl = 9
GoTo runamltn
End If
If ad008 <> "" Then
amkl = 8
GoTo runamltn
End If
If ad007 <> "" Then
amkl = 7
GoTo runamltn
End If
If ad006 <> "" Then
amkl = 6
GoTo runamltn
End If
If ad005 <> "" Then
amkl = 5
GoTo runamltn
End If
If ad004 <> "" Then
amkl = 4
GoTo runamltn
End If
If ad003 <> "" Then
amkl = 3
GoTo runamltn
End If
If ad002 <> "" Then
amkl = 2
GoTo runamltn
End If
If ad001 <> "" Then
amkl = 1
GoTo runamltn
End If

runamltn:
If GTX = 7 Then
If amkl = 16 Then
Label8.Caption = ad016
ad016 = ""
GoTo error
End If
If amkl = 15 Then
Label8.Caption = ad015
ad015 = ""
GoTo error
End If
If amkl = 14 Then
Label8.Caption = ad014
ad014 = ""
GoTo error
End If
If amkl = 13 Then
Label8.Caption = ad013
ad013 = ""
GoTo error
End If
If amkl = 12 Then
Label8.Caption = ad012
ad012 = ""
GoTo error
End If
If amkl = 11 Then
Label8.Caption = ad011
ad011 = ""
GoTo error
End If
If amkl = 10 Then
Label8.Caption = ad010
ad010 = ""
GoTo error
End If
If amkl = 9 Then
Label8.Caption = ad009
ad009 = ""
GoTo error
End If
If amkl = 8 Then
Label8.Caption = ad008
ad008 = ""
GoTo error
End If
If amkl = 7 Then
Label8.Caption = ad007
ad007 = ""
GoTo error
End If
If amkl = 6 Then
Label8.Caption = ad006
ad006 = ""
GoTo error
End If
If amkl = 5 Then
Label8.Caption = ad005
ad005 = ""
GoTo error
End If
If amkl = 4 Then
Label8.Caption = ad004
ad004 = ""
GoTo error
End If
If amkl = 3 Then
Label8.Caption = ad003
ad003 = ""
GoTo error
End If
If amkl = 2 Then
Label8.Caption = ad002
ad002 = ""
GoTo error
End If
If amkl = 1 Then
Label8.Caption = ad001
ad001 = ""
GoTo error
End If
End If



If GTX = 6 Then
If amkl = 16 Then
Label7.Caption = ad016
ad016 = ""
GoTo error
End If
If amkl = 15 Then
Label7.Caption = ad015
ad015 = ""
GoTo error
End If
If amkl = 14 Then
Label7.Caption = ad014
ad014 = ""
GoTo error
End If
If amkl = 13 Then
Label7.Caption = ad013
ad013 = ""
GoTo error
End If
If amkl = 12 Then
Label7.Caption = ad012
ad012 = ""
GoTo error
End If
If amkl = 11 Then
Label7.Caption = ad011
ad011 = ""
GoTo error
End If
If amkl = 10 Then
Label7.Caption = ad010
ad010 = ""
GoTo error
End If
If amkl = 9 Then
Label7.Caption = ad009
ad009 = ""
GoTo error
End If
If amkl = 8 Then
Label7.Caption = ad008
ad008 = ""
GoTo error
End If
If amkl = 7 Then
Label7.Caption = ad007
ad007 = ""
GoTo error
End If
If amkl = 6 Then
Label7.Caption = ad006
ad006 = ""
GoTo error
End If
If amkl = 5 Then
Label7.Caption = ad005
ad005 = ""
GoTo error
End If
If amkl = 4 Then
Label7.Caption = ad004
ad004 = ""
GoTo error
End If
If amkl = 3 Then
Label7.Caption = ad003
ad003 = ""
GoTo error
End If
If amkl = 2 Then
Label7.Caption = ad002
ad002 = ""
GoTo error
End If
If amkl = 1 Then
Label7.Caption = ad001
ad001 = ""
GoTo error
End If
End If


If GTX = 5 Then
If amkl = 16 Then
Label6.Caption = ad016
ad016 = ""
GoTo error
End If
If amkl = 15 Then
Label6.Caption = ad015
ad015 = ""
GoTo error
End If
If amkl = 14 Then
Label6.Caption = ad014
ad014 = ""
GoTo error
End If
If amkl = 13 Then
Label6.Caption = ad013
ad013 = ""
GoTo error
End If
If amkl = 12 Then
Label6.Caption = ad012
ad012 = ""
GoTo error
End If
If amkl = 11 Then
Label6.Caption = ad011
ad011 = ""
GoTo error
End If
If amkl = 10 Then
Label6.Caption = ad010
ad010 = ""
GoTo error
End If
If amkl = 9 Then
Label6.Caption = ad009
ad009 = ""
GoTo error
End If
If amkl = 8 Then
Label6.Caption = ad008
ad008 = ""
GoTo error
End If
If amkl = 7 Then
Label6.Caption = ad007
ad007 = ""
GoTo error
End If
If amkl = 6 Then
Label6.Caption = ad006
ad006 = ""
GoTo error
End If
If amkl = 5 Then
Label6.Caption = ad005
ad005 = ""
GoTo error
End If
If amkl = 4 Then
Label6.Caption = ad004
ad004 = ""
GoTo error
End If
If amkl = 3 Then
Label6.Caption = ad003
ad003 = ""
GoTo error
End If
If amkl = 2 Then
Label6.Caption = ad002
ad002 = ""
GoTo error
End If
If amkl = 1 Then
Label6.Caption = ad001
ad001 = ""
GoTo error
End If
End If


If GTX = 4 Then
If amkl = 16 Then
Label5.Caption = ad016
ad016 = ""
GoTo error
End If
If amkl = 15 Then
Label5.Caption = ad015
ad015 = ""
GoTo error
End If
If amkl = 14 Then
Label5.Caption = ad014
ad014 = ""
GoTo error
End If
If amkl = 13 Then
Label5.Caption = ad013
ad013 = ""
GoTo error
End If
If amkl = 12 Then
Label5.Caption = ad012
ad012 = ""
GoTo error
End If
If amkl = 11 Then
Label5.Caption = ad011
ad011 = ""
GoTo error
End If
If amkl = 10 Then
Label5.Caption = ad010
ad010 = ""
GoTo error
End If
If amkl = 9 Then
Label5.Caption = ad009
ad009 = ""
GoTo error
End If
If amkl = 8 Then
Label5.Caption = ad008
ad008 = ""
GoTo error
End If
If amkl = 7 Then
Label5.Caption = ad007
ad007 = ""
GoTo error
End If
If amkl = 6 Then
Label5.Caption = ad006
ad006 = ""
GoTo error
End If
If amkl = 5 Then
Label5.Caption = ad005
ad005 = ""
GoTo error
End If
If amkl = 4 Then
Label5.Caption = ad004
ad004 = ""
GoTo error
End If
If amkl = 3 Then
Label5.Caption = ad003
ad003 = ""
GoTo error
End If
If amkl = 2 Then
Label5.Caption = ad002
ad002 = ""
GoTo error
End If
If amkl = 1 Then
Label5.Caption = ad001
ad001 = ""
GoTo error
End If
End If


If GTX = 3 Then
If amkl = 16 Then
Label4.Caption = ad016
ad016 = ""
GoTo error
End If
If amkl = 15 Then
Label4.Caption = ad015
ad015 = ""
GoTo error
End If
If amkl = 14 Then
Label4.Caption = ad014
ad014 = ""
GoTo error
End If
If amkl = 13 Then
Label4.Caption = ad013
ad013 = ""
GoTo error
End If
If amkl = 12 Then
Label4.Caption = ad012
ad012 = ""
GoTo error
End If
If amkl = 11 Then
Label4.Caption = ad011
ad011 = ""
GoTo error
End If
If amkl = 10 Then
Label4.Caption = ad010
ad010 = ""
GoTo error
End If
If amkl = 9 Then
Label4.Caption = ad009
ad009 = ""
GoTo error
End If
If amkl = 8 Then
Label4.Caption = ad008
ad008 = ""
GoTo error
End If
If amkl = 7 Then
Label4.Caption = ad007
ad007 = ""
GoTo error
End If
If amkl = 6 Then
Label4.Caption = ad006
ad006 = ""
GoTo error
End If
If amkl = 5 Then
Label4.Caption = ad005
ad005 = ""
GoTo error
End If
If amkl = 4 Then
Label4.Caption = ad004
ad004 = ""
GoTo error
End If
If amkl = 3 Then
Label4.Caption = ad003
ad003 = ""
GoTo error
End If
If amkl = 2 Then
Label4.Caption = ad002
ad002 = ""
GoTo error
End If
If amkl = 1 Then
Label4.Caption = ad001
ad001 = ""
GoTo error
End If
End If


If GTX = 2 Then
If amkl = 16 Then
Label3.Caption = ad016
ad016 = ""
GoTo error
End If
If amkl = 15 Then
Label3.Caption = ad015
ad015 = ""
GoTo error
End If
If amkl = 14 Then
Label3.Caption = ad014
ad014 = ""
GoTo error
End If
If amkl = 13 Then
Label3.Caption = ad013
ad013 = ""
GoTo error
End If
If amkl = 12 Then
Label3.Caption = ad012
ad012 = ""
GoTo error
End If
If amkl = 11 Then
Label3.Caption = ad011
ad011 = ""
GoTo error
End If
If amkl = 10 Then
Label3.Caption = ad010
ad010 = ""
GoTo error
End If
If amkl = 9 Then
Label3.Caption = ad009
ad009 = ""
GoTo error
End If
If amkl = 8 Then
Label3.Caption = ad008
ad008 = ""
GoTo error
End If
If amkl = 7 Then
Label3.Caption = ad007
ad007 = ""
GoTo error
End If
If amkl = 6 Then
Label3.Caption = ad006
ad006 = ""
GoTo error
End If
If amkl = 5 Then
Label3.Caption = ad005
ad005 = ""
GoTo error
End If
If amkl = 4 Then
Label3.Caption = ad004
ad004 = ""
GoTo error
End If
If amkl = 3 Then
Label3.Caption = ad003
ad003 = ""
GoTo error
End If
If amkl = 2 Then
Label3.Caption = ad002
ad002 = ""
GoTo error
End If
If amkl = 1 Then
Label3.Caption = ad001
ad001 = ""
GoTo error
End If
End If


If GTX = 1 Then
If amkl = 16 Then
Label2.Caption = ad016
ad016 = ""
GoTo error
End If
If amkl = 15 Then
Label2.Caption = ad015
ad015 = ""
GoTo error
End If
If amkl = 14 Then
Label2.Caption = ad014
ad014 = ""
GoTo error
End If
If amkl = 13 Then
Label2.Caption = ad013
ad013 = ""
GoTo error
End If
If amkl = 12 Then
Label2.Caption = ad012
ad012 = ""
GoTo error
End If
If amkl = 11 Then
Label2.Caption = ad011
ad011 = ""
GoTo error
End If
If amkl = 10 Then
Label2.Caption = ad010
ad010 = ""
GoTo error
End If
If amkl = 9 Then
Label2.Caption = ad009
ad009 = ""
GoTo error
End If
If amkl = 8 Then
Label2.Caption = ad008
ad008 = ""
GoTo error
End If
If amkl = 7 Then
Label2.Caption = ad007
ad007 = ""
GoTo error
End If
If amkl = 6 Then
Label2.Caption = ad006
ad006 = ""
GoTo error
End If
If amkl = 5 Then
Label2.Caption = ad005
ad005 = ""
GoTo error
End If
If amkl = 4 Then
Label2.Caption = ad004
ad004 = ""
GoTo error
End If
If amkl = 3 Then
Label2.Caption = ad003
ad003 = ""
GoTo error
End If
If amkl = 2 Then
Label2.Caption = ad002
ad002 = ""
GoTo error
End If
If amkl = 1 Then
Label2.Caption = ad001
ad001 = ""
GoTo error
End If
End If


error:
End Sub
Public Sub AMLS()
On Error GoTo error
If Label8 <> "" Then
Stlr
End If
If Label7 <> "" Then
GTX = 7
AMLT
GoTo error
End If
If Label6 <> "" Then
GTX = 6
AMLT
GoTo error
End If
If Label5 <> "" Then
GTX = 5
AMLT
GoTo error
End If
If Label4 <> "" Then
GTX = 4
AMLT
GoTo error
End If
If Label3 <> "" Then
GTX = 3
AMLT
GoTo error
End If
If Label2 <> "" Then
GTX = 2
AMLT
GoTo error
End If
If Label1 <> "" Then
GTX = 1
AMLT
GoTo error
End If
error:
If Label1.Caption <> "" Then
Label1.BackColor = RGB(0, 0, 0)
End If
If Label2.Caption <> "" Then
Label2.BackColor = RGB(0, 0, 0)
End If
If Label3.Caption <> "" Then
Label3.BackColor = RGB(0, 0, 0)
End If
If Label4.Caption <> "" Then
Label4.BackColor = RGB(0, 0, 0)
End If
If Label5.Caption <> "" Then
Label5.BackColor = RGB(0, 0, 0)
End If
If Label6.Caption <> "" Then
Label6.BackColor = RGB(0, 0, 0)
End If
If Label7.Caption <> "" Then
Label7.BackColor = RGB(0, 0, 0)
End If
If Label8.Caption <> "" Then
Label8.BackColor = RGB(0, 0, 0)
End If
End Sub
Public Sub Sfxt()
On Error GoTo error
If Label1.Caption = "" Then
Label1.BackColor = RGB(255, 0, 0)
End If
If Label2.Caption = "" Then
Label2.BackColor = RGB(255, 0, 0)
End If
If Label3.Caption = "" Then
Label3.BackColor = RGB(255, 0, 0)
End If
If Label4.Caption = "" Then
Label4.BackColor = RGB(255, 0, 0)
End If
If Label5.Caption = "" Then
Label5.BackColor = RGB(255, 0, 0)
End If
If Label6.Caption = "" Then
Label6.BackColor = RGB(255, 0, 0)
End If
If Label7.Caption = "" Then
Label7.BackColor = RGB(255, 0, 0)
End If
If Label8.Caption = "" Then
Label8.BackColor = RGB(255, 0, 0)
End If
error:
End Sub
Public Sub Stlr()
On Error GoTo error
If Label1 <> "" And Label2 = "" Then
End
End If
If Label8 <> "" Then
Label1.Caption = Label2.Caption
Label2.Caption = Label3.Caption
Label3.Caption = Label4.Caption
Label4.Caption = Label5.Caption
Label5.Caption = Label6.Caption
Label6.Caption = Label7.Caption
Label7.Caption = Label8.Caption
Label8.Caption = ""
Sfxt
Exit Sub
End If
If Label7 <> "" And Label8 = "" Then
Label1.Caption = Label2.Caption
Label2.Caption = Label3.Caption
Label3.Caption = Label4.Caption
Label4.Caption = Label5.Caption
Label5.Caption = Label6.Caption
Label6.Caption = Label7.Caption
Label7.Caption = ""
Sfxt
Exit Sub
End If
If Label6 <> "" And Label7 = "" Then
Label1.Caption = Label2.Caption
Label2.Caption = Label3.Caption
Label3.Caption = Label4.Caption
Label4.Caption = Label5.Caption
Label5.Caption = Label6.Caption
Label6.Caption = ""
Sfxt
Exit Sub
End If
If Label5 <> "" And Label6 = "" Then
Label1.Caption = Label2.Caption
Label2.Caption = Label3.Caption
Label3.Caption = Label4.Caption
Label4.Caption = Label5.Caption
Label5.Caption = ""
Sfxt
Exit Sub
End If
If Label4 <> "" And Label5 = "" Then
Label1.Caption = Label2.Caption
Label2.Caption = Label3.Caption
Label3.Caption = Label4.Caption
Label4.Caption = ""
Sfxt
Exit Sub
End If
If Label3 <> "" And Label4 = "" Then
Label1.Caption = Label2.Caption
Label2.Caption = Label3.Caption
Label3.Caption = ""
Sfxt
Exit Sub
End If
If Label2 <> "" And Label3 = "" Then
Label1.Caption = Label2.Caption
Label2.Caption = ""
Sfxt
Exit Sub
End If
error:
End Sub


Private Sub Form_Activate()
On Error GoTo error
Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
Timer1.Interval = 5000
Timer1.Enabled = True
Timer2.Interval = 3000
Timer2.Enabled = True
error:
End Sub


Private Sub Form_Load()
On Error GoTo error
Form_Resize
MCTalkForm1.BackColor = RGB(255, 0, 0)
Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
SetWindowLong Me.hwnd, -20, GetWindowLong(Me.hwnd, -20) Or &H80000
SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
SetLayeredWindowAttributes Me.hwnd, vbRed, 0, 1
Label1.BackColor = RGB(0, 0, 0)
Label1.ForeColor = RGB(255, 255, 255)
Label2.BackColor = RGB(0, 0, 0)
Label2.ForeColor = RGB(255, 255, 255)
Label3.BackColor = RGB(0, 0, 0)
Label3.ForeColor = RGB(255, 255, 255)
Label4.BackColor = RGB(0, 0, 0)
Label4.ForeColor = RGB(255, 255, 255)
Label5.BackColor = RGB(0, 0, 0)
Label5.ForeColor = RGB(255, 255, 255)
Label6.BackColor = RGB(0, 0, 0)
Label6.ForeColor = RGB(255, 255, 255)
Label7.BackColor = RGB(0, 0, 0)
Label7.ForeColor = RGB(255, 255, 255)
Label8.BackColor = RGB(0, 0, 0)
Label8.ForeColor = RGB(255, 255, 255)
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label1.Caption = "MCTalk Bar Version 0.1.0.0 Made with VB6.0"
If Command = "" Then
ad001 = "Product by WMProject1217"
ad002 = "ßÙÁ¨ßÙÁ¨¸É±­~"
ad003 = "Minecraft Yes"
ad004 = "LBWNB"
ad005 = "Creeper? awwwwwwwwwwman"
ad006 = "Endding back has been enabled"
ad007 = "Ï²»¶³ª,Ìø,RAP,ÀºÇò,Music"
Sfxt
GoTo error
End If
Dim b() As String
b() = Split(Command, "=")
If b(0) = 1 Then
Label1.Caption = b(1)
End If
If b(0) = 2 Then
Label1.Caption = b(1)
ad001 = b(2)
End If
If b(0) = 3 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
End If
If b(0) = 4 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
End If
If b(0) = 5 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
End If
If b(0) = 6 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
End If
If b(0) = 7 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
ad006 = b(7)
End If
If b(0) = 8 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
ad006 = b(7)
ad007 = b(8)
End If
If b(0) = 9 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
ad006 = b(7)
ad007 = b(8)
ad008 = b(9)
End If
If b(0) = 10 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
ad006 = b(7)
ad007 = b(8)
ad008 = b(9)
ad009 = b(10)
End If
If b(0) = 11 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
ad006 = b(7)
ad007 = b(8)
ad008 = b(9)
ad009 = b(10)
ad010 = b(11)
End If
If b(0) = 12 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
ad006 = b(7)
ad007 = b(8)
ad008 = b(9)
ad009 = b(10)
ad010 = b(11)
ad011 = b(12)
End If
If b(0) = 13 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
ad006 = b(7)
ad007 = b(8)
ad008 = b(9)
ad009 = b(10)
ad010 = b(11)
ad011 = b(12)
ad012 = b(13)
End If
If b(0) = 14 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
ad006 = b(7)
ad007 = b(8)
ad008 = b(9)
ad009 = b(10)
ad010 = b(11)
ad011 = b(12)
ad012 = b(13)
ad013 = b(14)
End If
If b(0) = 15 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
ad006 = b(7)
ad007 = b(8)
ad008 = b(9)
ad009 = b(10)
ad010 = b(11)
ad011 = b(12)
ad012 = b(13)
ad013 = b(14)
ad014 = b(15)
End If
If b(0) = 16 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
ad006 = b(7)
ad007 = b(8)
ad008 = b(9)
ad009 = b(10)
ad010 = b(11)
ad011 = b(12)
ad012 = b(13)
ad013 = b(14)
ad014 = b(15)
ad015 = b(16)
End If
If b(0) = 17 Then
Label1.Caption = b(1)
ad001 = b(2)
ad002 = b(3)
ad003 = b(4)
ad004 = b(5)
ad005 = b(6)
ad006 = b(7)
ad007 = b(8)
ad008 = b(9)
ad009 = b(10)
ad010 = b(11)
ad011 = b(12)
ad012 = b(13)
ad013 = b(14)
ad014 = b(15)
ad015 = b(16)
ad016 = b(17)
End If
Sfxt
error:
End Sub

Private Sub Form_Resize()
On Error GoTo error
Label1.Left = 0
Label1.Width = 0.4 * MCTalkForm1.Width
Label1.Top = 0.84814814 * MCTalkForm1.Height
Label1.Height = 0.05432098 * MCTalkForm1.Height

Label2.Left = 0
Label2.Width = 0.4 * MCTalkForm1.Width
Label2.Top = 0.79876543 * MCTalkForm1.Height
Label2.Height = 0.05432098 * MCTalkForm1.Height

Label3.Left = 0
Label3.Width = 0.4 * MCTalkForm1.Width
Label3.Top = 0.74691358 * MCTalkForm1.Height
Label3.Height = 0.05432098 * MCTalkForm1.Height

Label4.Left = 0
Label4.Width = 0.4 * MCTalkForm1.Width
Label4.Top = 0.69629629 * MCTalkForm1.Height
Label4.Height = 0.05432098 * MCTalkForm1.Height

Label5.Left = 0
Label5.Width = 0.4 * MCTalkForm1.Width
Label5.Top = 0.64444444 * MCTalkForm1.Height
Label5.Height = 0.05432098 * MCTalkForm1.Height

Label6.Left = 0
Label6.Width = 0.4 * MCTalkForm1.Width
Label6.Top = 0.59135802 * MCTalkForm1.Height
Label6.Height = 0.05432098 * MCTalkForm1.Height

Label7.Left = 0
Label7.Width = 0.4 * MCTalkForm1.Width
Label7.Top = 0.5382716 * MCTalkForm1.Height
Label7.Height = 0.05432098 * MCTalkForm1.Height

Label8.Left = 0
Label8.Width = 0.4 * MCTalkForm1.Width
Label8.Top = 0.48765432 * MCTalkForm1.Height
Label8.Height = 0.05432098 * MCTalkForm1.Height
error:
End Sub

Private Sub Timer1_Timer()
On Error GoTo error
If adlst = 255 Then
GoTo error
End If
Stlr
error:
End Sub

Private Sub Timer2_Timer()
AMLS
End Sub
