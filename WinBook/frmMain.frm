VERSION 5.00
Begin VB.Form frmViewOSList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5790
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdFeedBack 
      Caption         =   "反馈"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "关于"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "检查更新"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "设置"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox lstOSList 
      Height          =   3660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
   Begin VB.ListBox lstOSTable 
      Height          =   420
      Left            =   4440
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "双击您要查看的系统查看，有些系统尚未完成，更多精彩功能，敬请期待下一个版本！"
      Height          =   855
      Left            =   3480
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Menu mnuLstOS 
      Caption         =   "1"
      Visible         =   0   'False
      Begin VB.Menu mnuView 
         Caption         =   "查看(&V)"
         Index           =   1
      End
      Begin VB.Menu menuFeedback 
         Caption         =   "反馈(&F)"
      End
   End
End
Attribute VB_Name = "frmViewOSList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type POINTAPI
        x As Long
        y As Long
End Type
Dim p As POINTAPI
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Sub cmdAbout_Click()
frmAbout.Show
End Sub

Private Sub Form_Load()
getLie2 lstOSTable, "ProductList", "TableName"
getLieOS
Me.Caption = "WinBook " & App.Major & "." & App.Minor & "." & App.Revision & " " & RC
End Sub
Sub getTableName()
   Dim rs As New ADODB.Recordset
   Dim CN As New ADODB.Connection
   Set CN = New ADODB.Connection
   CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= winbook_0.3.0.2.dat;Persist Security Info=False"

   Set rs = CN.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, Empty))
   Do Until rs.EOF
        If Left(rs!table_name, 4) <> "MSys" Then
            lstOSList.AddItem rs!table_name
        End If
        rs.MoveNext
   Loop
   rs.Close
   Set rs = Nothing
   CN.Close
   Set CN = Nothing
End Sub
Sub getLieOS()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=winbook_0.3.0.2.dat;Persist Security Info=False;Jet OLEDB:DataBase Password=siger@siger"
rs.Open "ProductList", con, 3, 2
';User ID=Creator  ;Jet OLEDB:System database=WinBook.mdw
While rs.EOF = False
lstOSList.AddItem rs.Fields("ProductName").Value
rs.MoveNext
Wend
End Sub
Sub getFieldName()
    Dim rs As ADODB.Recordset
    Dim CN As ADODB.Connection
    Dim FN As ADODB.Field
    Set CN = New ADODB.Connection
    Set rs = New ADODB.Recordset
      
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=winbook_0.3.0.2.dat;Persist Security Info=False"
    rs.Open "ProductlstOSList", CN
    For Each FN In rs.Fields
        lstOSList.AddItem FN.Name
    Next
    rs.Close
    Set rs = Nothing
    CN.Close
    Set CN = Nothing
End Sub
Sub getLie()

Dim Conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
Dim mdbPath As String
mdbPath = App.Path & "\winbook_0.3.0.2.dat"
 
Dim i As Integer
    Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                 mdbPath & ";Persist Security Info=False"
   Conn.CursorLocation = adUseClien '我加上了这行！！！
    rs.Open "select * from [user]", Conn, adOpenKeyset, adLockOptimistic  '以上连接数据库，应该没什么问题
    Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
  Values = rs.GetRows(-1, 0, "姓名") '取线路编号列的数据存入values数组中
    For i = 0 To UBound(Values, 2)
        lstOSList.AddItem Values(0, i)
    Next
  
   rs.Close
   Conn.Close
End Sub
Function getFilePathName(dlgObj As Object, setFilter As String) As String
'添加　commondialog1
'On Error GoTo err
    With dlgObj
        .DialogTitle = "请指定文件夹"
        .Filter = "文本文件(*." & setFilter & ")|*." & setFilter
        .ShowOpen
        getFilePathName = .FileName
    End With
 Exit Function
err:
 MsgBox "您没有选择文件或者文件夹中没有" & setFilter & "文件"
End Function

Private Sub lblHelp1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
a = MsgBox("您确定要退出WinBook?", vbYesNo + vbInformation, "提示")
If a = vbNo Then
Cancel = 1
Else
End
End If
End Sub

Private Sub lstOSList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
GetCursorPos p
End Sub

Private Sub lstOSList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
GetCursorPos p
PopupMenu mnuLstOS, , x, y
End If
End Sub

Private Sub lstOSList_DblClick()
For i = 0 To lstOSList.ListCount - 1
If lstOSList.Selected(i) = True Then
SystemItem = lstOSTable.List(lstOSList.ListIndex)
SystemFen = lstOSList.List(i)
Dim frmViewOSVer2 As New frmViewOSVer
frmViewOSVer2.Caption = lstOSList.List(i)
frmViewOSVer2.Show
End If
Next

End Sub
