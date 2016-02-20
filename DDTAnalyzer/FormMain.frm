VERSION 5.00
Begin VB.Form FormMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Siger Young DDTAnalyzer"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6015
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   3135
      Left            =   5760
      TabIndex        =   19
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame FrameStrength 
      Caption         =   "强化连点"
      Height          =   975
      Left            =   6000
      TabIndex        =   16
      Top             =   0
      Width           =   1815
      Begin VB.TextBox TextClickCount 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CommandClick 
         Caption         =   "点"
         Height          =   495
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Timer TimerClick 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2040
      Top             =   1440
   End
   Begin VB.CommandButton CommandCls 
      Caption         =   "清"
      Height          =   495
      Left            =   5400
      TabIndex        =   15
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton CommandFullPower 
      Caption         =   "满力角: "
      CausesValidation=   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Timer TimerSetPos 
      Interval        =   2000
      Left            =   2040
      Top             =   2160
   End
   Begin VB.TextBox TextLog 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3360
      Width           =   5655
   End
   Begin VB.TextBox TextWind 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   1935
   End
   Begin VB.ComboBox ComboModeSelect 
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton CommandHideExtraLine 
      Caption         =   "隐藏多余地图边框"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.OptionButton OptionPoints 
      Caption         =   "敌人和自己"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton OptionMapBox 
      Caption         =   "小地图框"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.OptionButton OptionMap 
      Caption         =   "小地图"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   2280
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.PictureBox PictureContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   2400
      MousePointer    =   2  'Cross
      ScaleHeight     =   2115
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.Line LineMapRight 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   1440
         X2              =   1440
         Y1              =   840
         Y2              =   1440
      End
      Begin VB.Line LineMapLeft 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   1440
         X2              =   1440
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Shape PointEnemy 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         Height          =   75
         Left            =   360
         Shape           =   2  'Oval
         Top             =   1320
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape PointSelf 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   75
         Left            =   360
         Shape           =   2  'Oval
         Top             =   1200
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Line LineMapBoxRight 
         BorderColor     =   &H000000FF&
         Visible         =   0   'False
         X1              =   1200
         X2              =   1200
         Y1              =   720
         Y2              =   1320
      End
      Begin VB.Line LineMapBoxLeft 
         BorderColor     =   &H80000001&
         Visible         =   0   'False
         X1              =   1200
         X2              =   1200
         Y1              =   1320
         Y2              =   1920
      End
   End
   Begin VB.Label LabelData 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label LabelPingCountInt 
      Caption         =   "屏"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label LabelPingCount 
      Caption         =   "屏"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label LabelPointDistance 
      Caption         =   "我敌水平距离: "
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label LabelMapBoxDistance 
      Caption         =   "小地图框宽度: "
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label LabelMapDistance 
      Caption         =   "小地图宽度: "
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClickI As Integer
Private Sub ComboModeSelect_Click()
    Dim ID As Integer
    ID = ComboModeSelect.ListIndex
    UpdateDistanceData
    'LabelData = GetLineFromDataById(Id, 4)
End Sub

Private Sub Command1_Click()
    If Command1.Caption = ">" Then
        Me.Width = Me.Width + FrameStrength.Width
        Command1.Caption = "<"
    ElseIf Command1.Caption = "<" Then
        Me.Width = 6105
        Command1.Caption = ">"
    End If
End Sub

Private Sub CommandCls_Click()
    PictureContainer.Cls
    PictureContainer.Picture = Nothing
End Sub

Private Sub CommandClick_Click()
    ClickI = 0
    If CommandClick.Caption = "点" Then
        TimerClick.Enabled = True
        CommandClick.Caption = "停"
    ElseIf CommandClick.Caption = "停" Then
        TimerClick.Enabled = False
        CommandClick.Caption = "点"
    End If
End Sub

Private Sub CommandFullPower_Click()
    PictureContainer.Cls
    PictureContainer.DrawWidth = 3
    PictureContainer.Line (PointSelf.Left + PointSelf.Width / 2, PointSelf.Top + PointSelf.Height / 2)-(PointEnemy.Left + PointEnemy.Width / 2, PointEnemy.Top + PointEnemy.Height / 2), RGB(0, 20, 255)
End Sub

Private Sub CommandHideExtraLine_Click()
    LineMapRight.Visible = False
    LineMapLeft.Visible = False
    LineMapBoxRight.Visible = False
    LineMapBoxLeft.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then OptionMapBox.Value = True
    If KeyCode = vbKeyF2 Then OptionPoints.Value = True
    If KeyCode = vbKeyF3 Then
        TextWind.SelStart = 0
        TextWind.SelLength = Len(TextWind.Text)
    End If
End Sub

Private Sub Form_Load()
    Dim p As Long
    p = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H2 Or &H1
    SetWindowLong Me.hWnd, GWL_EXSTYLE, p Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, 0, 110, LWA_ALPHA
    AddItemsToComboBox ComboModeSelect
    ComboModeSelect.ListIndex = 0
End Sub

Private Sub AddItemsToComboBox(Combo As ComboBox)
    Dim ItemsArray() As String, ItemsCount As Integer
    ItemsArray = GetItemsFromData(PowerAngleDataArray)
    ItemsCount = UBound(ItemsArray) - LBound(ItemsArray) + 1
    For i = 0 To ItemsCount - 1
        Combo.AddItem ItemsArray(i)
    Next
End Sub

Private Sub PictureContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Select Case GetSelectItem
        Case 0
            ChangeVerticalLineTopLeft LineMapRight, X, Y, PictureContainer.Height
        Case 1
            ChangeVerticalLineTopLeft LineMapBoxRight, X, Y, PictureContainer.Height
        Case 2
            ChangePointTopLeft PointEnemy, X, Y
        End Select
    End If
    If Button = vbLeftButton Then
        Select Case GetSelectItem
        Case 0
            ChangeVerticalLineTopLeft LineMapLeft, X, Y, PictureContainer.Height
        Case 1
            ChangeVerticalLineTopLeft LineMapBoxLeft, X, Y, PictureContainer.Height
        Case 2
            ChangePointTopLeft PointSelf, X, Y
        End Select
    End If
    UpdateDistanceData
End Sub

Private Sub ChangeVerticalLineTopLeft(Line As Line, Left As Single, Top As Single, Length As Integer)
    Line.Visible = True
    Line.X1 = Left
    Line.X2 = Left
    Line.Y1 = Top - Length / 2
    Line.Y2 = Top + Length / 2
End Sub

Private Sub ChangePointTopLeft(Point As Shape, Left As Single, Top As Single)
    Point.Visible = True
    Point.Left = Left - Point.Width / 2
    Point.Top = Top - Point.Height / 2
End Sub

Private Function GetPointX(Point As Shape)
    GetPointX = Point.Left + Point.Width / 2
End Function

Private Function GetPointY(Point As Shape)
    GetPointY = Point.Top + Point.Height / 2
End Function

Private Function GetSelectItem()
    If OptionMap.Value = True Then GetSelectItem = 0
    If OptionMapBox.Value = True Then GetSelectItem = 1
    If OptionPoints.Value = True Then GetSelectItem = 2
End Function

Private Sub UpdateDistanceData()
    Dim Map As Integer, MapBox As Integer, Point As Integer
    Map = Abs(LineMapLeft.X1 - LineMapRight.X1)
    MapBox = Abs(LineMapBoxLeft.X1 - LineMapBoxRight.X1)
    Point = Abs(PointSelf.Left - PointEnemy.Left)
    LabelMapDistance.Caption = "地图宽度: " & Map
    LabelMapBoxDistance.Caption = "地图框宽度: " & MapBox
    LabelPointDistance.Caption = "我敌宽度: " & Point
    If MapBox > 0 Then
        Dim PingCount As Single
        PingCount = Point / (MapBox / 10)
        LabelPingCount.Caption = Format(PingCount, "0.00") & "屏"
        LabelPingCountInt.Caption = Round(PingCount) & "屏"
    End If
    Dim Power As Integer, Angle As Integer
    Angle = Round(GetAngleByDistanceAndWind(ComboModeSelect.ListIndex, Round(PingCount), Val(TextWind.Text)))
    Power = Round(GetPowerByDistanceAndWind(ComboModeSelect.ListIndex, Round(PingCount), Val(TextWind.Text)))
    LabelData = "角度: " & Angle & "  " & "力度: " & Power
    Dim X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer
    X1 = GetPointX(PointSelf)
    Y1 = GetPointY(PointSelf)
    X2 = GetPointX(PointEnemy)
    Y2 = GetPointY(PointEnemy)
    TextLog.Text = X1 & " " & Y1 & " " & X2 & " " & Y2
    CommandFullPower.Caption = "满力角: " & Int(Atn(GetKByPoints(X1, Y1, X2, Y2)) * (180 / PI)) & "度"
    TextLog.Text = TextLog.Text & vbCrLf & GetKByPoints(X1, Y1, X2, Y2)
    'TextLog.Text = TextLog.Text & "Angle: " & Angle & " " & "Power: " & Power & " " & "Distance: " & Round(PingCount) & vbCrLf
End Sub


Private Sub TextWind_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDistanceData
    If IsNumeric(TextWind) = False Then KeyCode = 0
End Sub

Private Sub TextWind_KeyUp(KeyCode As Integer, Shift As Integer)
    UpdateDistanceData
    If IsNumeric(TextWind) = False Then KeyCode = 0
End Sub

Private Sub TimerSetPos_Timer()
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H2 Or &H1
End Sub

Private Function GetKByPoints(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
    Dim k As Single
    If X1 - X2 = 0 Then
        k = 0
    Else
        k = Format((Y1 - Y2) / (X1 - X2), "0.000")
    End If
    Debug.Print k & ": " & X1 & Y1 & X2 & Y2
    GetKByPoints = k
End Function

Private Sub TimerClick_Timer()
    If ClickI > Val(TextClickCount.Text) Then TimerClick.Enabled = False
    ClickMouse
    ClickI = ClickI + 1
    TextLog = TextLog & vbCrLf & ClickI
End Sub
