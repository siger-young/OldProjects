VERSION 5.00
Begin VB.Form frmViewOSVer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2640
   Icon            =   "frmViewOSVer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2640
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ComboBox cbOSVerShaixuan 
      Appearance      =   0  'Flat
      Height          =   300
      ItemData        =   "frmViewOSVer.frx":000C
      Left            =   0
      List            =   "frmViewOSVer.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.ListBox lstOSVer 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmViewOSVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dpdown As Boolean
Private Sub cbOSVerShaixuan_Click()
If dpdown = True Then
getLie2 Me.lstOSVer, SystemItem, Me.cbOSVerShaixuan.Text
End If
VerORTag = Me.cbOSVerShaixuan.Text
End Sub

Private Sub cbOSVerShaixuan_DropDown()
dpdown = True
End Sub

Private Sub Form_Load()
With Me.cbOSVerShaixuan
.AddItem "Version"
.AddItem "BuildTag"
End With
getLie2 Me.lstOSVer, SystemItem, "Version"
End Sub


Private Sub lstOSVer_DblClick()

For i = 0 To lstOSVer.ListCount - 1
If lstOSVer.Selected(i) = True Then
SystemBuild = lstOSVer.List(i)
SystemIndex = lstOSVer.ListIndex + 1
Dim frmViewOS2 As New frmViewOS
NT6AddBar frmViewOS2.lblCurrentDir, SystemFen, SystemBuild
getHang frmViewOS2.lstOS, SystemItem, SystemIndex
viewOS frmViewOS2.lstOS, frmViewOS2.Text2, frmViewOS2.Text1, frmViewOS2.Text3, frmViewOS2.Text4, frmViewOS2.Text5, frmViewOS2.Label16, frmViewOS2.Label17, frmViewOS2.Label18, frmViewOS2.Text6, frmViewOS2.Text7, frmViewOS2.Text8, frmViewOS2.Label19, frmViewOS2.Command1, frmViewOS2.Text9
frmViewOS2.Caption = "WinBook " & App.Major & "." & App.Minor & "." & App.Revision & " " & RC
frmViewOS2.Show
End If
Next
End Sub
