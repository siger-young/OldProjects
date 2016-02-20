Attribute VB_Name = "ModRun"
'download by http://www.codefans.net
'Title 标题 | 设置控制台窗体标题 | Function Title(Txt As String)
'Echo 标题  | 输出控制台信息     | Sub Echo(szOut As String)
'EchoGet    | 得到控制台输入信息 |  Function EchoGet() As String
'返回 -- 控制台输入输出模式   Function GetMode(dwMode As Long)
'设置 -- 控制台输入输出模式   Function SetMode(lpMode As Long)
'返回 -- 控制台输出代码页     Function GetOutputCP(dwMode As Long)
'设置 -- 控制台输出代码页     Function SetOutputCP(wCodePageID As Long)
'获取 -- 控制台输入代码页     Function GetCP()
'设置 -- 控制台输入代码页     Function SetCP(wCodePageID As Long)
'获取 -- 控制台光标大小       Function GetCursor(By As CONSOLE_CURSOR_INFO)
'设置 -- 控制台光标大小       Function SetCursor(Size As Long, bHide As Long)
'返回窗口尺寸的最大可能性     Function GetMaxConsole(ByVal XY As COORD)
'写 -- 设置控制台窗口大小     Function SetConsoleWindow(Left As Long, Right As Long, Top As Long, Bottom As Long)
'返回控制台队列事件数         Function GetEvents()
'清除控制台输入缓冲区         Sub Cls()
'设置屏幕文本属性             Sub Color(Num As Long)
'返回鼠标按钮数               Function GetMouseButton()
'返回鼠标按钮数               Function Set_Cursor(Xx As Long, Yy As Long)
'改变屏幕缓冲区大小           Function ScreenSize(Xx As Long, Yy As Long)
'改变显示屏幕缓冲区           Function SetScreen()
'CmdLine    | 程序启动命令
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Public Function ConsoleMain()
    Dim mColors As Long
    Dim sInStr  As String
    Dim Tmp(2)  As String
    Dim lInStr As String
    Dim lang As String
    If Dir(App.Path & "\Settings.ini") = "" Then
        Open "Settings.ini" For Output As #1
        Print #1, "Language=English"
        lang = "English"
    ElseIf Dir(App.Path & "\Settings.ini") <> "" Then
        Open "Settings.ini" For Input As #2
        Input #2, lang
    End If
    If lang = "English" Then GoTo English
    '-----------------------
English:
    
    Title "Win Book"
    mWrite "Win Book " & "Ver " & App.Major & "." & App.Minor & "." & App.Revision & Space(1) & RC & "\n"
    mWrite "Program starting up...\n"
    Sleep 1000
    mWrite "Done.\n"
    mWrite "Please type 'help' to see help of the commands in Win Book Console program!\n"
Englishs:
    
    mWrite ">"
    sInStr = mRead()
    If LCase(sInStr) = "exit" Then Exit Function
    lInStr = LCase(sInStr)
    Select Case LCase(sInStr)
    Case "help":

        mWrite "exit  Exit Win Book\n"
        mWrite "gui  Quit Win Book and restart Win Book Gui \n"
        mWrite "help  Display help\n"
        mWrite "sys   Display the staus of your current operating system\n"
        mWrite "***** To see the usage of a command. Please add '/?' behind the command ****\n"
        
    Case Else
        mWrite "!!!Command Not Exist\n"
    End Select
    GoTo Englishs:
    
End Function


