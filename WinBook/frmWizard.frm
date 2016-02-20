VERSION 5.00
Begin VB.Form frmWizard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WinBook v0.4.8 "
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "否"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "是"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "是否要以向导模式启动Win Book?"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   840
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "欢迎使用WinBook !"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.NewXing.com
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim unloadmode As Integer

Private Sub Command1_Click()
    frmWizardMain.Show 1
End Sub

Private Sub Command2_Click()
unloadmode = 1
Unload Me
frmViewOSList.Show
End Sub

Private Sub Form_Load()

    Image1.Picture = LoadPicture(App.Path & "\Winbook32.ico")
Me.Caption = "WinBook " & App.Major & "." & App.Minor & "." & App.Revision & " " & RC

End Sub

Private Sub Timer1_Timer()
    Dim p As POINTAPI
    Dim f As RECT
    GetCursorPos p '得到MOUSE位置
    GetWindowRect Me.hwnd, f '得到窗体的位置
    If Me.WindowState <> 1 Then
        If p.x > f.Left And p.x < f.Right And p.y > f.Top And p.y < f.Bottom Then
            'MOUSE 在窗体上
            If Me.Top < 0 Then
                Me.Top = -10
                Me.Show
            ElseIf Me.Left < 0 Then
                Me.Left = -10
                Me.Show
            ElseIf Me.Left + Me.Width >= Screen.Width Then
                Me.Left = Screen.Width - Me.Width + 10
                Me.Show
            End If
        Else
            If f.Top <= 4 Then
                Me.Top = 40 - Me.Height
            ElseIf f.Left <= 4 Then
                Me.Left = 40 - Me.Width
            ElseIf Me.Left + Me.Width >= Screen.Width - 4 Then
                Me.Left = Screen.Width - 40
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If unloadmode <> 1 Then
    a = MsgBox("真的要退出？", vbYesNo + vbQuestion, "提示")
    If a = vbYes Then
        End
    ElseIf a = vbNo Then
        Cancel = 1
    End If
    End If
End Sub
