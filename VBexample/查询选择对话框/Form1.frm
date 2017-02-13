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
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      Caption         =   "按班级"
      Height          =   300
      Left            =   2880
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "按姓名"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "按学号"
      Height          =   180
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "请选择查询方式"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Option1.Value = True Then
        yn = MsgBox("选择学号查询码？", 67, "提示")
        If yn = vbNo Then
            End
        ElseIf yn = vbYes Then
            tr = InputBox$("请输入学号（0~9数字）：", "输入学号")
                If Len(tr) = 10 Then
                    yn = MsgBox("您输入的学号是" & tr, 64, "提示")
                Else
                    MsgBox "输入错误！不能查询！", 16, "特别提示"
                End If
        ElseIf yn = vbCancel Then
            MsgBox "按学号查询操作被取消！", 48, "警告"
        End If
    End If
    
    If Option2.Value = True Then
        tr = InputBox$("请输入姓名", "输入姓名")
            yn = MsgBox("您输入的姓名是" & tr, 64, "提示")
    End If
    
    If Option3.Value = True Then
        tr = InputBox$("请输入班级名称：", "输入班级")
        yn = MsgBox("您输入的班级是" & tr, 64, "提示")
    End If
    
                    
End Sub

Private Sub Form_Load()
    Form1.Caption = "输入对话框和消息框"
End Sub
