VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   5535
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   1815
      Left            =   7080
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   4200
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   720
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()        '"确认"按钮的单击事件
    Text2.Text = Text1.Text        '将text1的text属性的值赋给text2的text属性
    Text1.Text = ""
    Label1.Caption = "左边文本框的文字已复制到右边的文本框"
    Command2.Enabled = True
    
End Sub


Private Sub Command2_Click()
Unload Form1

End Sub

Private Sub Form_Load()
    Label1.FontName = "楷体"
    Label1.FontSize = 12
    Label1.ForeColor = vbRed
    Label1.Caption = "请在左边文本框输入文字，然后单击“确认”按钮"
    Text1.Text = ""
    Text2.Text = ""
    Text2.Locked = True             '文本框text2被锁定，不能进行文字编辑
    
End Sub


Private Sub Text1_GotFocus()        '文本框text1获得焦点事件
    Label1.Caption = "请在左边文本框输入文字，然后单击“确认按钮”"
    Text2.Text = ""
    Command2.Enabled = False        '此时“退出”按钮不可用
    
    

End Sub
