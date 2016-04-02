VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "线性电路求解"
   ClientHeight    =   5775
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7065
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "求解"
      Height          =   5535
      Left            =   3480
      TabIndex        =   18
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton EleSolve 
         Caption         =   "求解"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox ResultList 
         Height          =   3975
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "结果:"
         Height          =   180
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "元件"
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton EleClr 
         Caption         =   "清空"
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton EleDel 
         Caption         =   "删除选中"
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton EleAdd 
         Caption         =   "添加"
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1320
         TabIndex        =   12
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1320
         TabIndex        =   11
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1320
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "电阻"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "恒流电源"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "恒压电源"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Ele2 
         Height          =   270
         Left            =   1080
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Ele1 
         Height          =   270
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.ListBox EleList 
         Height          =   2400
         Left            =   240
         TabIndex        =   1
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ω"
         Height          =   180
         Left            =   2280
         TabIndex        =   14
         Top             =   2040
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A"
         Height          =   180
         Left            =   2280
         TabIndex        =   13
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "V"
         Height          =   300
         Left            =   2280
         TabIndex        =   10
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "负极结点:"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "正极结点:"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   810
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Ele1_GotFocus()
    Ele1.SelStart = 0
    Ele1.SelLength = Len(Ele1.Text)
End Sub

Private Sub Ele2_GotFocus()
    Ele2.SelStart = 0
    Ele2.SelLength = Len(Ele2.Text)
End Sub

Private Sub EleAdd_Click()
    Dim TempStyle As EleStyle, TempValue As Double
    If IsNumeric(Ele1.Text) And IsNumeric(Ele2.Text) Then
        If Option1.Value = True Then TempStyle = Vol: TempValue = Val(Text1.Text)
        If Option2.Value = True Then TempStyle = Cur: TempValue = Val(Text2.Text)
        If Option3.Value = True Then TempStyle = Res: TempValue = Val(Text3.Text)
        With Eles(EleCount)
            .Node1 = Val(Ele1.Text)
            .Node2 = Val(Ele2.Text)
            .Style = TempStyle
            .Value = TempValue
        End With
        EleCount = EleCount + 1
        ShowEleList
        EleList.ListIndex = EleCount - 1
    Else
        MsgBox "请输入结点号!"
    End If
End Sub

Private Sub EleClr_Click()
    EleCount = 0
    ShowEleList
    ResultList.Text = ""
End Sub

Private Sub EleDel_Click()
    Dim i As Integer
    If EleList.ListIndex <> -1 Then
        EleCount = EleCount - 1
        If EleList.ListIndex < EleCount Then
            For i = EleList.ListIndex To EleCount - 1
                Eles(i) = Eles(i + 1)
            Next
        End If
    End If
    ShowEleList
End Sub

Private Sub EleSolve_Click()
    Dim i As Integer, j As Integer
    Dim XCount As Integer
    Dim Mat1() As Double
    Dim Mat2() As Double
    Dim NodeMax As Integer
    Dim VolCount As Integer
    '统计元件使用的节点数
    NodeMax = 0
    VolCount = 0
    For i = 0 To EleCount - 1
        If Eles(i).Node1 > NodeMax Then NodeMax = Eles(i).Node1
        If Eles(i).Node2 > NodeMax Then NodeMax = Eles(i).Node2
        If Eles(i).Style = Vol Then VolCount = VolCount + 1
    Next
    XCount = NodeMax + VolCount + 1
    '建立矩阵
    ReDim Mat1(XCount, XCount) As Double
    ReDim Mat2(XCount) As Double
    For i = 1 To XCount
        For j = 1 To XCount
            Mat1(i, j) = 0
        Next
        Mat2(i) = 0
    Next
    '列节点的电流方程
    For i = 0 To NodeMax
        For j = 0 To EleCount - 1
            With Eles(j)
                If .Style = Res Then
                    If .Node1 = i Then
                        Mat1(i + 1, i + 1) = Mat1(i + 1, i + 1) + 1 / .Value
                        Mat1(i + 1, .Node2 + 1) = Mat1(i + 1, .Node2 + 1) - 1 / .Value
                    End If
                    If .Node2 = i Then
                        Mat1(i + 1, i + 1) = Mat1(i + 1, i + 1) + 1 / .Value
                        Mat1(i + 1, .Node1 + 1) = Mat1(i + 1, .Node1 + 1) - 1 / .Value
                    End If
                ElseIf .Style = Cur Then
                    If .Node1 = i Then Mat2(i + 1) = Mat2(i + 1) + .Value
                    If .Node2 = i Then Mat2(i + 1) = Mat2(i + 1) - .Value
                End If
            End With
        Next
    Next
    '考虑恒压电源
    VolCount = 1
    For i = 0 To EleCount - 1
        If Eles(i).Style = Vol Then
            With Eles(i)
                Mat1(.Node1 + 1, VolCount + NodeMax + 1) = Mat1(.Node1 + 1, VolCount + NodeMax + 1) - 1
                Mat1(.Node2 + 1, VolCount + NodeMax + 1) = Mat1(.Node2 + 1, VolCount + NodeMax + 1) + 1
                Mat1(VolCount + NodeMax + 1, .Node1 + 1) = Mat1(VolCount + NodeMax + 1, .Node1 + 1) + 1
                Mat1(VolCount + NodeMax + 1, .Node2 + 1) = Mat1(VolCount + NodeMax + 1, .Node2 + 1) - 1
                Mat2(VolCount + NodeMax + 1) = .Value
                VolCount = VolCount + 1
            End With
        End If
    Next
    '设置电势零点
    For i = 1 To XCount
        Mat1(1, i) = 0
    Next
    Mat2(1) = 0
    Mat1(1, 1) = 1
    '输出调试信息（矩阵）
    For i = 1 To XCount
        For j = 1 To XCount
            Debug.Print Mat1(i, j),
        Next
        Debug.Print Mat2(i)
    Next
    '输出结果
    ResultList.Text = ""
    If LEGauss(XCount, Mat1(), Mat2()) Then
        For i = 1 To NodeMax + 1
            ResultList.Text = ResultList.Text & "结点" & i - 1 & " 电势" & Mat2(i) & "V" & vbCrLf
        Next
        VolCount = 1
        For i = 0 To EleCount - 1
            With Eles(i)
                ResultList.Text = ResultList.Text & EleStyleName(.Style) & i & " 电压"
                ResultList.Text = ResultList.Text & Abs(Mat2(.Node1 + 1) - Mat2(.Node2 + 1)) & "V"
                ResultList.Text = ResultList.Text & " 电流"
                If .Style = Vol Then
                    ResultList.Text = ResultList.Text & Abs(Mat2(VolCount + NodeMax + 1))
                    VolCount = VolCount + 1
                ElseIf .Style = Cur Then
                    ResultList.Text = ResultList.Text & Abs(.Value)
                Else
                    ResultList.Text = ResultList.Text & Abs((Mat2(.Node1 + 1) - Mat2(.Node2 + 1)) / .Value)
                End If
                ResultList.Text = ResultList.Text & "A" & vbCrLf
            End With
        Next
    Else
        MsgBox "求解失败"
    End If
End Sub

Private Sub Form_Load()
    EleCount = 0
    EleStyleName(0) = "恒压电源"
    EleStyleName(1) = "恒流电源"
    EleStyleName(2) = "电阻"
    EleStyleUnit(0) = "V"
    EleStyleUnit(1) = "A"
    EleStyleUnit(2) = "Ω"
End Sub

Sub ShowEleList()
    Dim i As Integer
    EleList.Clear
    If EleCount > 0 Then
        For i = 0 To EleCount - 1
            With Eles(i)
                EleList.AddItem "[" & i & "] {" & .Node1 & "," & .Node2 & "} " & EleStyleName(.Style) & " " & .Value & EleStyleUnit(.Style)
            End With
        Next
    End If
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Option1.Value = True
End Sub

Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Option2.Value = True
End Sub

Private Sub Text3_GotFocus()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Option3.Value = True
End Sub
