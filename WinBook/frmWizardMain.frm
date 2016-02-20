VERSION 5.00
Begin VB.Form frmWizardMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Wizard"
   ClientHeight    =   17595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   17595
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Image imgBack 
      Height          =   495
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
End
Attribute VB_Name = "frmWizardMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Sub Form_Load()
    'SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    imgBack.Picture = LoadResPicture(2, vbResBitmap)
    MsgBox "欢迎使用Win Book 设置向导", vbSystemModal, "Win Book"
End Sub

Private Sub Form_Resize()
    imgBack.Top = 0
    imgBack.Left = 0
    imgBack.Width = frmWizardMain.Width
    imgBack.Height = frmWizardMain.Height
End Sub
