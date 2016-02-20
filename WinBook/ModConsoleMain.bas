Attribute VB_Name = "ModConsoleMain"
Option Explicit
Private Declare Function GetStdHandle Lib "kernel32" (ByVal HandleType As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal cToWrite As Long, ByRef cWritten As Long, Optional ByVal lpOverlapped As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&
Dim mhInput   As Long
Dim mhOutPut  As Long
Public Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Sub Main()
    Dim cmdline As String
    cmdline = Command()
    If InStr(cmdline, "/console") = 1 Then

    Dim sName As String
    mhInput = GetStdHandle(STD_INPUT_HANDLE)                                    '┓初始化
    mhOutPut = GetStdHandle(STD_OUTPUT_HANDLE)                                  '┣标准输入、标准输出的句柄
    SetConsoleTextAttribute mhOutPut, 14                                        '┛VB中的QBColor函数颜色
    SetConsoleTitle "WinBook v0.4.8 Alpha Console Program"
    OutPutLine "WinBook-v0.4.8 Alpha"
    OutPutLine vbCrLf
    OutPutLine "Please type " & Chr(34) & "help" & Chr(34) & "to see the help ! " & sName
    OutPutLine vbCrLf
    OutPutLine ">"
    If InPutLine <> "" Then
    MsgBox InPutLine
    End If
    If InPutLine = "help" Or "HELP" Or "Help" Then
    MsgBox "**********WinBook-v0.4.8 Alpha Help " + vbCrLf + "chkupt Check the lastest version of the winbook", vbOKOnly + vbInformation, "Help"
    'Win Book 帮助部分
    End If
    Call InPutLine(1)
    CloseHandle mhInput
    CloseHandle mhOutPut
    Else
    FreeConsole
    CloseHandle mhInput
    CloseHandle mhOutPut
    frmStartUp.Show
    End If
End Sub

Public Function OutPutLine(ByVal sString As String) As Long                     '┓
    WriteFile mhOutPut, sString, lstrlen(sString), OutPutLine                   '┣标准输出
End Function                                                                    '┛

Public Function InPutLine(Optional ByVal lenBuffer As Long = 1024) As String    '┓
    Dim bChar() As Byte, lReadChars As Long                                     '┃
    ReDim bChar(lenBuffer)                                                      '┃
    ReadFile mhInput, bChar(0), lenBuffer, lReadChars, ByVal 0&                 '┃
    lReadChars = lReadChars - 2                                                 '┣标准输入
    If lReadChars < 1 Then Exit Function                                        '┃
    ReDim Preserve bChar(lReadChars - 1)                                        '┃
    InPutLine = StrConv(bChar, vbUnicode)                                       '┃
End Function                                                                    '┛

