VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "结束"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确认"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r, s As Integer         '定义窗体级各模块共享的变量

Private Sub Command1_Click()
    Label1.Caption = "请输入一个100以内的正整数"
    Randomize                       '对随机数生成器做初始化动作
    r = Int((100 * Rnd) + 1)        '随机生成100以内的正整数
    s = 1
    Text1.Locked = False            '设置文本框为可编辑状态
    Command1.Enabled = False
    Command2.Enabled = True
    Text1.Text = ""
    Text1.SetFocus                  '设置文本框焦点
End Sub

Private Sub Command2_Click()
    If Text1.Text = "" Or (Not IsNumeric(Text1.Text)) Then      'IsNumeric函数判断是否为数字字符串
        Label1.Caption = "空字符或非数字字符！请重新输入"
        Text1.Text = ""                                         '清空文本框
    ElseIf Val(Text1.Text) > r Then                             'Val函数将数字字符串转换成数字
        Label1.Caption = Text1.Text & "大了，已猜了" & s & "次"
        s = s + 1
        Text1.Text = ""
    ElseIf Val(Text1.Text) < r Then
        Label1.Caption = Text1.Text & "小了，已猜了" & s & "次"
        s = s + 1
        Text1.Text = ""
    Else
        Label1.Caption = "恭喜您答对了！共猜了" & s & "次"
        Text1.Locked = True
        Command1.Enabled = True
        Command2.Enabled = False
    End If
    Text1.SetFocus             '设置文本框焦点
End Sub

Private Sub Command3_Click()
    Unload Form1
    
End Sub

Private Sub Form_Load()
    Label1.Caption = "请单击“开始”按钮启动游戏"
    Text1.Text = ""
    Command1.Enabled = True
    Command2.Enabled = False
    
End Sub
