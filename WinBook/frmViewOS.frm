VERSION 5.00
Begin VB.Form frmViewOS 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmViewOS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5790
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   600
      TabIndex        =   28
      Text            =   "Text8"
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查看"
      Height          =   375
      Left            =   4440
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Text9"
      Top             =   4080
      Width           =   4695
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Text7"
      Top             =   2640
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text6"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text5"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text3"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   720
      Width           =   1695
   End
   Begin VB.ListBox lstOS 
      Height          =   2040
      Left            =   1800
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label19 
      Height          =   255
      Left            =   4080
      TabIndex        =   25
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label18 
      Height          =   255
      Left            =   3360
      TabIndex        =   22
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label17 
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "版本关键字"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "BIOS日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "截图"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "修复"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "代号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "序列号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "架构"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "版本"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "语言"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "阶段"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "SKU"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "下载链接"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "产品名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblCurrentDir 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmViewOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
