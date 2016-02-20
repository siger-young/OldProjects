Attribute VB_Name = "modAccess"
Public SystemItem As String
Public SystemFen As String
Public SystemBuild As String
Public SystemIndex As Integer
Dim scshot, Down, Fixes
Public VerORTag As String

Public Sub getTableName(List As Variant)
   Dim rs As New ADODB.Recordset
   Dim CN As New ADODB.Connection
   Set CN = New ADODB.Connection
   CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= winbook_0.3.0.2.dat;Persist Security Info=False"

   Set rs = CN.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, Empty))
   Do Until rs.EOF
        If Left(rs!table_name, 4) <> "MSys" Then
            List.AddItem rs!table_name
        End If
        rs.MoveNext
   Loop
   rs.Close
   Set rs = Nothing
   CN.Close
   Set CN = Nothing
End Sub
Public Sub getHang(lst As ListBox, TableName As String, SysIdx As Integer)

lst.Clear
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=winbook_0.3.0.2.dat;Jet OLEDB:DataBase Password=siger@siger"
sql1 = "select top " & SystemIndex & " * from " & TableName & ";"
sql2 = "select " & VerORTag & " from " & TableName & " where " & VerORTag & "='x'"
rs.Open sql1, con, 3, 2
';User ID=Creator  ;Jet OLEDB:System database=WinBook.mdw
rs.MoveLast
If rs("Screenshot").Value & "" = "" Then
scshot = "无"
ElseIf rs("Screenshot").Value & "" <> "" Then
scshot = rs("Screenshot").Value
End If
If rs("DownloadLink").Value & "" = "" Then
Down = "敬请期待下一个版本"
ElseIf rs("DownloadLink").Value & "" <> "" Then
Down = rs("DownloadLink").Value
End If
With lst
.AddItem rs("ProductName").Value & ""
.AddItem rs("Codename").Value & ""
.AddItem rs("Version").Value & ""
.AddItem rs("Stage").Value & ""
.AddItem rs("BuildTag").Value & ""
.AddItem rs("Architecture").Value & ""
.AddItem rs("Edition").Value & ""
.AddItem rs("Language").Value & ""
.AddItem rs("BIOSDate").Value & ""
.AddItem rs("SerialNumber").Value & ""
.AddItem rs("Fixes").Value & ""
.AddItem scshot
.AddItem Down
End With
End Sub
Public Sub getLie2(lst As ListBox, TableName As String, Lie As String)
lst.Clear
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=winbook_0.3.0.2.dat;Jet OLEDB:DataBase Password=siger@siger"
rs.Open TableName, con, 3, 2
';User ID=Creator  ;Jet OLEDB:System database=WinBook.mdw
While rs.EOF = False
lst.AddItem rs.Fields(Lie).Value
rs.MoveNext
Wend
End Sub
Public Sub viewOS(List As ListBox, ProductName As TextBox, Codename As TextBox, Version As TextBox, Stage As TextBox, BuildTag As TextBox, Architecture As Label, Edition As Label, Language As Label, BIOSDate As TextBox, Sn As TextBox, Fixes As TextBox, lblScreenshot As Label, Screenshotbtn As CommandButton, DownloadLink As TextBox)
If scshot = "无" Then
lblScreenshot.Visible = True
Screenshotbtn.Visible = False
ElseIf scshot <> "无" Then
lblScreenshot.Visible = False
Screenshotbtn.Visible = True
End If
ProductName = List.List(0)
Codename = List.List(1)
Version = List.List(2)
Stage = List.List(3)
BuildTag = List.List(4)
Architecture = List.List(5)
Edition = List.List(6)
Language = List.List(7)
BIOSDate = List.List(8)
Sn = List.List(9)
Fixes = List.List(10)
lblScreenshot = scshot
DownloadLink = List.List(12)
End Sub


Public Sub getFieldName(List As Variant)
    Dim rs As ADODB.Recordset
    Dim CN As ADODB.Connection
    Dim FN As ADODB.Field
    Set CN = New ADODB.Connection
    Set rs = New ADODB.Recordset
      
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=winbook_0.3.0.2.dat;Persist Security Info=False"

    rs.Open "ProductlstOSList", CN
    For Each FN In rs.Fields
        List.AddItem FN.Name
    Next
    rs.Close
    Set rs = Nothing
    CN.Close
    Set CN = Nothing
End Sub
Public Sub getLie(List As Variant)

Dim Conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
Dim mdbPath As String
mdbPath = App.Path & "\winbook_0.3.0.2.dat"
 
Dim i As Integer
    Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                 mdbPath & ";Persist Security Info=False;"
   Conn.CursorLocation = adUseClien '我加上了这行！！！
    rs.Open "select * from [user]", Conn, adOpenKeyset, adLockOptimistic  '以上连接数据库，应该没什么问题
    Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
  Values = rs.GetRows(-1, 0, "姓名") '取线路编号列的数据存入values数组中
    For i = 0 To UBound(Values, 2)
        List.AddItem Values(0, i)
    Next
  
   rs.Close
   Conn.Close
End Sub
Public Function getFilePathName(dlgObj As Object, setFilter As String) As String
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
'-----------------------NT6地址栏
Public Sub NT6AddBar(lblTxt As Label, add1 As String, add2 As String)
If add2 = "" Then
lblTxt.Caption = add1
Else
lblTxt.Caption = add1 & " > " & add2
End If
End Sub

