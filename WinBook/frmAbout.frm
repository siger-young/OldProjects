VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2160
      Top             =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "http://winbook.usr.me"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   240
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   180
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   540
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Caption = "关于"
    SaveFileFromRes 101, "CUSTOM", App.Path & "\WinBook32.ico"
Image1.Picture = LoadPicture(App.Path & "\Winbook32.ico")
Label1 = WinBookVer & vbCrLf & "开发者名单:" & vbCrLf & vbCrLf & "数据库开发:th1r5bvn23" & vbCrLf & vbCrLf & "程序:Ysc" & vbCrLf & vbCrLf & "贡献者名单" & vbCrLf & "hotrz" & vbCrLf & "driver1998" & vbCrLf & "随便问我" & vbCrLf & "huoqianyu" & vbCrLf & "632481545" & vbCrLf & "以上排名不分先后"
End Sub

Private Sub Label2_Click()
Shell "explorer" & " http://winbook.usr.me"
End Sub

Private Sub Timer1_Timer()

'Dim speed As Integer
Randomize
'speed = Int((25 * Rnd) + 1)
If Label1.Top <= -40 - Label1.Height Then
Label1.Top = Me.Height
Else
Label1.Top = Label1.Top - 15
End If
End Sub
'Private Sub Timer1_Timer()
'Static i As Byte
'i = i + 1
'If i >= 255 Then
'i = i - 255
'End If
'Label1.Top = Label1.Top - 18
'Label1.ForeColor = RGB(255 - i, i, 255 - i)
'End Sub

