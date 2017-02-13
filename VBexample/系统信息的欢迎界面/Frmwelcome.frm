VERSION 5.00
Begin VB.Form Frmwelcome 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7290
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "楷体"
      Size            =   20.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   6000
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6000
      Top             =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "学生信息管理系统欢迎你！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4320
      TabIndex        =   0
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "Frmwelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim j As Integer   '设置电波发送的循环变量
Dim r As Byte, g As Byte, b As Byte
Dim i As Integer    '设置生成直线的循环变量

Private Sub Form_Click()
    Timer1.Interval = 0   '设置定时器的时间间隔为0
    Cls                    '清除屏幕
    Frmwelcome.Hide
    
End Sub

Private Sub Form_Load()
    BackColor = RGB(0, 0, 0)    '实际运行中可设置背景色为黑色
                                '塔身的生成
    j = 0
    DrawWidth = 1               '设置线宽
    PSet (ScaleWidth / 3, 1000) '在三分之一宽度和高度1000处画圆
    ForeColor = RGB(0, 255, 0)   '设置前景色为绿色
    For i = 1 To 50 Step 5        '设置直线宽度值为循环变量
        DrawWidth = i            '设置直线宽度值为循环变量
        Line -Step(0, ScaleHeight / 10) '从当前位置按步幅scaleHeight/10划线
    Next i
    
                            '生成塔上部的突起
    DrawWidth = 1           '重新设置线宽
    FillStyle = 6           '设置填充样式为十字线
    FillColor = ForeColor   '设置填充线的颜色为前景色
    Circle (ScaleWidth / 3, 2000), 300, , , , 0.5 '画横向椭圆
    
End Sub

Private Sub Timer1_Timer()
   
    FillStyle = 1     '图形方法生成的圆或方框的模式为透明
    DrawStyle = 2     '输入的线性样式为虚数
    r = 255 * Rnd       '生成的红色随机参数
    g = 255 * Rnd       '生成的绿色随机参数
    b = 255 * Rnd       '生成的蓝色随机参数
    j = j + 1
    Circle (ScaleWidth / 3, 1000), 300 * j, RGB(r, g, b) '以定时器规定的时间间隔画半径相差300的颜色随机的圆
    
    If j = 10 Then j = 0
End Sub

Private Sub Timer2_Timer()
    For i = 0 To 10000 Step 50
    Label1.Top = (Label1.Top - 1)  '使控件位置发生变化
    Label1.ForeColor = RGB(r, g, b) '使控件前景色发生变化
    If Label1.Top = -Label1.Height Then Label1.Top = ScaleHeight
    Next i
End Sub
