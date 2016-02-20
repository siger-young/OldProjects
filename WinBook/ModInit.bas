Attribute VB_Name = "ModInit"
Option Explicit
'download by http://www.codefans.net
'Help
Public WinBookVer As String
Public Const Cmd_Cls_Help = "清除屏幕。\n\nCLS\n"
Public Const Cmd_Echo_Help = "显示信息，或将命令回显打开或关上。\n\n  Echo [ON | OFF]\n  Echo [message]\n\n要显示当前回显设置，键入不带参数的 ECHO。"
Public RC As String

''定义API函数中用到的所有常量
''GetStdHandle函数的 nStdHandle参数的取值
Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Const STD_ERROR_HANDLE = -12&
''SetConsoleTextAttribute函数的wAttributes参数的取值（按RGB方式组合）
Public Const FOREGROUND_BLUE = &H1
Public Const FOREGROUND_GREEN = &H2
Public Const FOREGROUND_RED = &H4
Public Const FOREGROUND_INTENSITY = &H8
Public Const BACKGROUND_BLUE = &H10
Public Const BACKGROUND_GREEN = &H20
Public Const BACKGROUND_RED = &H40
Public Const BACKGROUND_INTENSITY = &H80
''SetConsoleMode的输入模式
Public Const ENABLE_LINE_INPUT = &H2
Public Const ENABLE_ECHO_INPUT = &H4
Public Const ENABLE_MOUSE_INPUT = &H10
Public Const ENABLE_PROCESSED_INPUT = &H1
Public Const ENABLE_WINDOW_INPUT = &H8
''SetConsoleMode的输出模式
Public Const ENABLE_PROCESSED_OUTPUT = &H1
Public Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

Public Const CONSOLE_TEXTMODE_BUFFER = 1
Public Const CREATE_NEW_CONSOLE = &H10

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public Type CONSOLE_CURSOR_INFO
    dwSize As Long
    bVisible As Long
End Type
Public Type CHAR_INFO
    Char As Integer
    Attributes As Integer
End Type
Public Type COORD
    x As Integer
    y As Integer
End Type
Public Type SMALL_RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type
Public Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD
    dwCursorPosition As COORD
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long

'清除控制台输入缓冲区
Private Declare Function FlushConsoleInputBuffer Lib "kernel32" (ByVal hConsoleInput As Long) As Long
'设置屏幕文本属性
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
'返回鼠标按钮数
Private Declare Function GetNumberOfConsoleMouseButtons Lib "kernel32" (lpNumberOfMouseButtons As Long) As Long
'设置控制台光标位置
Private Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, dwCursorPosition As COORD) As Long
'将字符写入屏幕缓冲区
Private Declare Function FillConsoleOutputCharacter Lib "kernel32" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long
'将属性写入屏幕缓冲区
Private Declare Function FillConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttribute As Long, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
'返回屏幕缓冲区信息
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
'改变屏幕缓冲区大小
Private Declare Function SetConsoleScreenBufferSize Lib "kernel32" (ByVal hConsoleOutput As Long, dwSize As COORD) As Long
'改变显示屏幕缓冲区
Private Declare Function SetConsoleActiveScreenBuffer Lib "kernel32" (ByVal hConsoleOutput As Long) As Long
'滚动屏幕缓冲区中的数据
Private Declare Function ScrollConsoleScreenBuffer Lib "kernel32" Alias "ScrollConsoleScreenBufferA" (ByVal hConsoleOutput As Long, lpScrollRectangle As SMALL_RECT, lpClipRectangle As SMALL_RECT, dwDestinationOrigin As COORD, lpFill As CHAR_INFO) As Long

'为当前进程建立 / 释放 -- 控制台
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
'设置 / 返回 -- 控制台窗口标题
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
'读 / 写 -- 控制台缓冲区                 '输入数据'写控制台屏幕缓冲区
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumherOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
'读 / 写 -- 控制台缓冲区                 读屏幕缓冲区数据'直接控制屏幕缓冲区
Private Declare Function ReadConsoleOutput Lib "kernel32" Alias "ReadConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpReadRegion As SMALL_RECT) As Long
Private Declare Function WriteConsoleOutput Lib "kernel32" Alias "WriteConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpWriteRegion As SMALL_RECT) As Long
'设置 / 返回 -- 控制台输入输出模式
Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function GetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, lpMode As Long) As Long
'获取 / 设置 -- 控制台输出代码页
Private Declare Function GetConsoleOutputCP Lib "kernel32" () As Long
Private Declare Function SetConsoleOutputCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
'获取 / 设置 -- 控制台输入代码页
Private Declare Function GetConsoleCP Lib "kernel32" () As Long
Private Declare Function SetConsoleCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
'返回 / 设置 -- 控制台光标大小
Private Declare Function GetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare Function SetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
'读 / 写 -- 控制台属性字符串
Private Declare Function ReadConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, lpAttribute As Long, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfAttrsRead As Long) As Long
Private Declare Function WriteConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, lpAttribute As Integer, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
'读 / 写 -- 控制台屏幕缓冲区字符串
Private Declare Function ReadConsoleOutputCharacter Lib "kernel32" Alias "ReadConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfCharsRead As Long) As Long
Private Declare Function WriteConsoleOutputCharacter Lib "kernel32" Alias "WriteConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long

'返回窗口尺寸的最大可能性 / 设置控制台窗口大小
Private Declare Function GetLargestConsoleWindowSize Lib "kernel32" (ByVal hConsoleOutput As Long) As COORD
Private Declare Function SetConsoleWindowInfo Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal bAbsolute As Long, lpConsoleWindow As SMALL_RECT) As Long

'将句柄返回给新的屏幕缓冲区
Private Declare Function CreateConsoleScreenBuffer Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwFlags As Long, lpScreenBufferData As Any) As Long
'返回控制台队列事件数
Private Declare Function GetNumberOfConsoleInputEvents Lib "kernel32" (ByVal hConsoleInput As Long, lpNumberOfEvents As Long) As Long
'设置控制台进程的单个句柄
Private Declare Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long
'向控制台进程组发送信号
Private Declare Function GenerateConsoleCtrlEvent Lib "kernel32" (ByVal dwCtrlEvent As Long, ByVal dwProcessGroupId As Long) As Long

'any = SECURITY_ATTRIBUTES
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_ALL = &H10000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2


Public hConsoleIn As Long ''控制台窗口的 input handle
Public hConsoleOut As Long ''控制台窗口的output handle
Public hConsoleErr As Long ''控制台窗口的error handle
Public hConsole As Long
Public CmdLine As String
Public bEcho As Boolean
Public Function SaveFileFromRes(vntresourceId As Variant, sTYPE As String, sfilename As String) As Boolean
    Dim bytimage() As Byte
    Dim ifilenum As Integer
    On Error GoTo SaveFileFromRes_err
    SaveFileFromRes = True
    bytimage = LoadResData(vntresourceId, sTYPE)
    ifilenum = FreeFile
    
    Open sfilename For Binary As ifilenum
    Put #ifilenum, , bytimage
    Close ifilenum
    
SaveFileFromRes_err: Exit Function
    SaveFileFromRes = False: Exit Function
End Function

''主程序
Public Sub Main()
Shell "regsvr32 /s " & App.Path & "\msader15.dll"
Shell "regsvr32 /s " & App.Path & "\msado15.dll"
    CmdLine = Command$
    RC = "Alpha"
    WinBookVer = "WinBook " & App.Major & "." & App.Minor & "." & App.Revision & " " & RC
    If InStr(CmdLine, "/console") = 1 Then
        AllocConsole ''创建 console window
        SetConsoleTitle "VB控制台应用程序"
        '设置console window的标题
        '取得console window的三个句柄
        hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
        hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
        hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
        SetConsoleTextAttribute hConsoleOut, FOREGROUND_GREEN Or FOREGROUND_INTENSITY
        '前景：亮绿；背景：黑
        hConsole = CreateFile("CONOUT$", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, 3, 0, 0&)
        bEcho = True
        ConsoleMain
        '启动
        FreeConsole ''销毁 console window
    Else
        FreeConsole
        frmStartUp.Show
    End If
End Sub

Public Function Title(Txt As String)     '设置console window的标题
    SetConsoleTitle Txt
End Function

Public Sub mWrite(szOut As String)         '写 -- 控制台缓冲区
    szOut = Replace(szOut, "\n", vbCrLf)
    szOut = Replace(szOut, "\b", vbBack)
    szOut = Replace(szOut, "\r", vbCr)
    szOut = Replace(szOut, "\t", vbTab)
    szOut = Replace(szOut, "\\", "\")
    WriteConsole hConsoleOut, szOut, LenB(StrConv(szOut, vbFromUnicode)), vbNull, vbNull
End Sub

Public Function mRead() As String         '读 -- 控制台缓冲区
    Dim sUserInput As String * 256
    Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
    ''截掉字符串结尾的&H00和回车、换行符
    mRead = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
End Function

Public Function Cmd_Color(Color As Long)
    Cmd_Color = SetConsoleTextAttribute(hConsoleOut, Color)
End Function

Public Function Cmd_Cls(Cmd As String, param As String) As Boolean
    Dim csbi As CONSOLE_SCREEN_BUFFER_INFO
    Dim coPos As COORD
    Dim dwWritten As Long
    
    dwWritten = 0
    
    If param = "/?" Then
        mWrite Cmd_Cls_Help
        Cmd_Cls = False
        Exit Function
    End If
    
    GetConsoleScreenBufferInfo hConsole, csbi
    
    coPos.x = 0
    coPos.y = 0
    Call FillConsoleOutputAttribute(hConsole, 0&, csbi.dwSize.x * csbi.dwSize.y, coPos, dwWritten)
    Call FillConsoleOutputCharacter(hConsole, CByte(&H14), csbi.dwSize.x * csbi.dwSize.y, coPos, dwWritten)
    Call SetConsoleCursorPosition(hConsole, coPos)
    
    Cmd_Cls = True
End Function

Public Function Cmd_Echo(Cmd As String, param As String)
    Dim File As String
    
    If param = "/?" Then
        mWrite Cmd_Echo_Help
        Exit Function
    ElseIf param = "on" Then
        bEcho = True
        Exit Function
    ElseIf param = "off" Then
        bEcho = False
        Exit Function
    End If
    If InStr(param, ">") <> 0 Then
        If Mid(param, InStr(param, ">") - 2, 1) = "^" Then
            mWrite param
        Else
            Cmd = Mid(param, InStr(param, ">") - 1)
            File = Mid(param, InStr(param, ">") + 1)
            mWriteFile File, Cmd
        End If
    End If
    mWrite param
End Function
'download by http://www.codefans.net
Public Function mWriteFile(File As String, Info As String, Optional WriteType As String = "Output")
    Dim FileNum As Long
    FileNum = FreeFile
    If WriteType = "Output" Then
        
        Open File For Output As #FileNum
        Print #FileNum, Info
        
    ElseIf WriteType = "Append" Then
        
        Open File For Append As #FileNum
        Print #FileNum, Info
        
    ElseIf WriteType = "Binary" Then
        
        Open File For Binary As #FileNum
        Put #FileNum, , Info
        
    End If
    Close #FileNum
End Function
'download by http://www.codefans.net
Sub sleepFrm()
    Sleep 5000
End Sub
Function ReStr(s As String, Search As String) As String
    Dim i As Integer, res As String
    res = s
    Do While InStr(res, Search)
        i = InStr(res, Search)
        res = Left(res, i - 1) & Mid(res, i + 1)
    Loop
    ReStr = res
End Function
