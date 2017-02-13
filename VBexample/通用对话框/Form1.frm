VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
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
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "设置字体"
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存文件"
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开文件"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   2640
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "请选择右键按钮进行操作"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yn As Integer  '定义一个整型变量

Private Sub Command1_Click()
    CommonDialog1.Filter = "文档(*.doc;*.rtf;*.txt)|*.doc;*.ref;*.txt|所有文件(*.*)|*.*" '设置文件列表框中所显示文件的类型
    CommonDialog1.Action = 1    '调用“打开”对话框
    Label1.Caption = "打开" + CommonDialog1.FileName  '设置标签标题
    Text1.Text = ""   '设置文本框初始值
    Open CommonDialog1.FileName For Input As #1  '打开选择的文件
    Do While Not EOF(1)
        Line Input #1, inputdata   '读一行数据
        Text1.Text = Text1.Text + inputdata + vbCrLf
    Loop
    Close #1        '关闭文件
End Sub

Private Sub Command2_Click()
    CommonDialog1.FileName = "default.txt"  '保存文件的默认文件名
    CommonDialog1.DefaultExt = "txt" '默认的扩展名
    CommonDialog1.Action = 2   '调用“另存为”对话框
    Label1.Caption = "保存" + CommonDialog1.FileName
    Open CommonDialog1.FileName For Output As #1  '打开文件写入数据
    Print #1, Text1.Text    '将文本框内的文本写入文件
    Close #1    '关闭文件
End Sub

Private Sub Command3_Click()
    CommonDialog1.Flags = 3   '设置显示字体为屏幕字体或打印机字体均可
    CommonDialog1.Action = 4  '调用“字体”对话框
    Label1.Caption = "为文件" + CommonDialog1.FileName + "设置字体"
    Text1.FontName = CommonDialog1.FontName '设置文本字体
    Text1.FontSize = CommonDialog1.FontSize '设置文本字号
    Text1.FontBold = CommonDialog1.FontBold '设置文本粗体
    Text1.FontItalic = CommonDialog1.FontItalic '设置文本斜体
    Text1.FontStrikethru = CommonDialog1.FontStrikethru  '设置文本删除线
    Text1.FontUnderline = CommonDialog1.FontUnderline '设置文本下划线
    Text1.ForeColor = CommonDialog1.Color  '设置文本颜色
End Sub

Private Sub Command4_Click()
    yn = MsgBox("在退出之前您的文件保存了吗？", 4, "提示")
    If yn = 6 Then
        End
    End If
End Sub
