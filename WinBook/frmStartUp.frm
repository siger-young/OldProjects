VERSION 5.00
Begin VB.Form frmStartUp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Starting Up..."
   ClientHeight    =   5415
   ClientLeft      =   8175
   ClientTop       =   7680
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStartUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer timSleep 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3240
      Top             =   4680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   2160
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   4440
      Width           =   1815
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim aaaa As Integer
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Dim Alpha As Integer
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long

Private Sub Form_Click()
App.Title = "WinBook " & App.Major & "." & App.Minor & "." & App.Revision & " " & RC
    Unload Me
    frmViewOSList.Show
End Sub

Private Sub Form_Load()
    aaaa = 0
    Timer2.Enabled = False
    Me.AutoRedraw = True
    Me.Picture = LoadResPicture(1, vbResBitmap)
    Dim Ret As Long
    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    Timer1.Interval = 20
    Timer2.Interval = 20
    
    Label2.Caption = "Ver " & App.Major & "." & App.Minor & "." & App.Revision & " " & RC
End Sub

Private Sub Timer1_Timer()
    Alpha = Alpha + 5 '控制速度
    If Alpha > 255 Then '控制透明度
        Timer1.Enabled = False
        timSleep.Enabled = True
        
        Exit Sub
    End If
    SetLayeredWindowAttributes Me.hwnd, 0, Alpha, LWA_ALPHA
    
End Sub

Private Sub Timer2_Timer()
    Alpha = Alpha - 5 '控制速度
    If Alpha < 1 Then '控制透明度
        Timer2.Enabled = False
        Unload Me
        frmViewOSList.Show
        Exit Sub
    End If
    SetLayeredWindowAttributes Me.hwnd, 0, Alpha, LWA_ALPHA
    
End Sub

Private Sub timSleep_Timer()
    If aaaa = 3 Then
        Timer2.Enabled = True
    Else
        aaaa = aaaa + 1
    End If
End Sub
