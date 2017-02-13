VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " 查看日志"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查看"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F As String
Dim H As Long

Private Sub Command1_Click()
    F = App.Path & "\" & "xsxx.log"
    H = FreeFile
    Open F For Input As #H   '以顺序方式打开文件
    Text1.Text = ""
    Do Until EOF(1)     '文件未到尾部
        Line Input #H, newline   '读文件中的一行到变量newline中
        Text1.Text = Text1.Text + newline + Chr(13) + Chr(10) '将变量值显示在文本框内，每行尾部加回车换行符。Chr(13)表示回车，Chr(10)表示换行，回车回到当前行的行首，接着输入的话，本行以前的内容会被逐一覆盖。
    Loop       '循环直到文件尾
    Close #H     '关闭文件
End Sub

Private Sub Command2_Click()
    Unload Me     '退出当前窗体，并不退出程序
End Sub
