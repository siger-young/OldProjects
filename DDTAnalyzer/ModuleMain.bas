Attribute VB_Name = "ModuleMain"
Option Explicit
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub mouse_event Lib "user32 " (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'��ָ�����ڵĽṹ��ȡ����Ϣ
'����ָ���Ľ���
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'��ϵͳע��һ��ָ�����ȼ�
Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'ȡ���ȼ����ͷ�ռ�õ���Դ
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long
'�������API������ע��ϵͳ���ȼ�������ģ�����ʵ�ֹ����������ʾ
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_ABSOLUTE = &H8000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA As Long = &H2
Public Const WS_EX_LAYERED As Long = &H80000
Public Const ANGLE_FORMULA_LINE = 2
Public Const POWER_FORMULA_LINE = 3
Public Const TABLE_LINE = 4
Public Const PI = 3.1415926
  '�ȼ���־����,�����жϵ����̰���������ʱ�Ƿ������������趨���ȼ�
Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = (-4)

'����ϵͳ���ȼ�,ԭ�жϱ�ʾ,�����ص���Ŀ���
Public preWinProc As Long, MyhWnd As Long, uVirtKey As Long

'�ȼ����ع���
Public Type POINTAPI
  X As Long
  Y As Long
End Type
Public PowerAngleData As String
Public PowerAngleDataArray() As String


Public Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then     '������ص��ȼ���־����
        Select Case wParam
        Case 1
            OperationOfArrows vbKeyNumpad2
        Case 2
            OperationOfArrows vbKeyNumpad8
        Case 3
            OperationOfArrows vbKeyNumpad4
        Case 4
            OperationOfArrows vbKeyNumpad6
        End Select
      End If
    '��������ȼ�,���߲����������õ��ȼ�,��������Ȩ��ϵͳ,��������ȼ�
    WndProc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam)
End Function

Public Sub OperationOfArrows(Arrow As KeyCodeConstants)
    Select Case Arrow
    Case vbKeyNumpad2
        MoveMouse 0, 1
    Case vbKeyNumpad8
        MoveMouse 0, -1
    Case vbKeyNumpad4
        MoveMouse -1, 0
    Case vbKeyNumpad6
        MoveMouse 1, 0
    End Select
    ClickMouse
End Sub

Sub Main()
    Dim Modifiers As Long
    PowerAngleData = GetFile("data.dat")
    PowerAngleDataArray = GetDataArray(PowerAngleData)
    FormMain.Show
    preWinProc = GetWindowLong(FormMain.hWnd, GWL_WNDPROC)
    SetWindowLong FormMain.hWnd, GWL_WNDPROC, AddressOf WndProc
    RegisterHotKey FormMain.hWnd, 1, Modifiers, vbKeyNumpad2
    RegisterHotKey FormMain.hWnd, 2, Modifiers, vbKeyNumpad8
    RegisterHotKey FormMain.hWnd, 3, Modifiers, vbKeyNumpad4
    RegisterHotKey FormMain.hWnd, 4, Modifiers, vbKeyNumpad6
End Sub

Public Function GetDataFromArray(ID As Integer)
    Dim CurrentDataArray() As String, CurrentData As String
    CurrentData = PowerAngleDataArray(ID)
    CurrentDataArray = Split(CurrentData, vbCrLf)
End Function

Public Function GetLineFromDataById(ID As Integer, LineNumber As Integer)
    GetLineFromDataById = Split(PowerAngleDataArray(ID), vbCrLf)(LineNumber - 1)
End Function

Public Function GetAngleFormulaById(ID As Integer)
    GetAngleFormulaById = Replace(GetLineFromDataById(ID, ANGLE_FORMULA_LINE), "ANGLE=", "")
End Function

Public Function GetPowerFormulaById(ID As Integer)
    GetPowerFormulaById = Replace(GetLineFromDataById(ID, POWER_FORMULA_LINE), "POWER=", "")
End Function

Public Function GetPowerFromTableByDistance(ID As Integer, Distance As Integer)
    If Distance <= 0 Or Distance > 20 Then Exit Function
    Dim PowerTable() As String, TempString As String
    TempString = GetLineFromDataById(ID, TABLE_LINE)
    If TempString = "" Then
        GetPowerFromTableByDistance = Replace(GetPowerFormulaById(ID), "POWER=", "")
    Else
        PowerTable = Split(TempString, " ")
        GetPowerFromTableByDistance = Val(PowerTable(Distance - 1))
    End If
End Function

Public Function GetPowerByDistanceAndWind(ID As Integer, Distance As Integer, Wind As Single)
    Dim PowerNormal As Integer
    Dim TempString As String
    PowerNormal = GetPowerFromTableByDistance(ID, Distance)
    TempString = GetPowerFormulaById(ID)
    If IsNumeric(TempString) Then
        GetPowerByDistanceAndWind = Val(TempString)
    Else
        TempString = Replace(Replace(Replace(Replace(TempString, "WIND", Wind), "DISTANCE", Distance), "POWER", PowerNormal), "NONE", PowerNormal)
        GetPowerByDistanceAndWind = Val(Eval(TempString))
    End If
End Function

Public Function GetAngleByDistanceAndWind(ID As Integer, Distance As Integer, Wind As Single)
    Dim TempString As String
    TempString = GetAngleFormulaById(ID)
    If IsNumeric(TempString) Then
        GetAngleByDistanceAndWind = TempString
    Else
        TempString = Replace(Replace(TempString, "WIND", Wind), "DISTANCE", Distance)
        GetAngleByDistanceAndWind = Eval(TempString)
    End If
End Function

Public Function GetItemsFromData(DataArray() As String)
    Dim ItemsArray() As String, ItemsCount As Integer, i As Integer
    ItemsCount = UBound(DataArray) - LBound(DataArray) + 1
    ReDim ItemsArray(ItemsCount - 1) As String
    For i = 0 To ItemsCount - 1
        ItemsArray(i) = Split(DataArray(i), vbCrLf)(0)
    Next
    GetItemsFromData = ItemsArray
End Function

Private Function GetFile(FileName As String)
    Dim FileContent As String, FileContentTemp As String
    Open App.Path & "\" & FileName For Input As #1
    While Not EOF(1)
        Line Input #1, FileContentTemp
        FileContent = FileContent & FileContentTemp & vbCrLf
    Wend
    Close #1
    GetFile = FileContent
End Function

Private Function GetDataArray(Data As String)
    Dim DataArray() As String
    DataArray = Split(Data, "###" + vbCrLf)
    GetDataArray = DataArray
End Function

Public Sub ClickMouse()
    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0&, 0&  '�����������ظ�һ�¾���˫����
End Sub

Public Sub MoveMouse(AbsX As Integer, AbsY As Integer)
    Dim Position As POINTAPI
    GetCursorPos Position
    SetCursorPos Position.X + AbsX, Position.Y + AbsY
    ClickMouse
End Sub

Public Function Eval(ByVal Expressions As String) As Double
    Dim Mssc As Object
    Set Mssc = CreateObject("MSScriptControl.ScriptControl")
    Mssc.Language = "vbscript"
    Eval = Mssc.Eval(Expressions)
End Function
