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
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   960
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   15
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "选课器"
      Height          =   3495
      Left            =   4200
      TabIndex        =   6
      Top             =   1800
      Width           =   5775
      Begin VB.CommandButton Command8 
         Caption         =   "《"
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "》"
         Height          =   495
         Left            =   2400
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "<"
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   ">"
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.ListBox List2 
         Height          =   1860
         ItemData        =   "Form1.frx":0000
         Left            =   3480
         List            =   "Form1.frx":0002
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   1950
         ItemData        =   "Form1.frx":0004
         Left            =   360
         List            =   "Form1.frx":0006
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "通过单击相应按钮，选择本学期需要选学的课程"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   3000
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "选学课程"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "本学期待选课程"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   495
      Left            =   7680
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "提交"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "选课"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "请输入学号"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.Print Text1.Text & "学生已选课程如下："
End Sub

Private Sub Command2_Click()
    List1.AddItem "邓小平理论": List1.AddItem "大学英语"
    List1.AddItem "高等数学": List1.AddItem "信息技术导论"
    List1.AddItem "VB程序设计": List1.AddItem "网站开发技术"
    List1.AddItem "数据库原理与应用": List1.AddItem "多媒体课件制作"
    List1.AddItem "平面设计": List1.AddItem "信息安全"
    List1.AddItem "电子商务"
End Sub

Private Sub Command3_Click()
    MsgBox "选课成功，已经写入数据库", , "提示"
End Sub

Private Sub Command4_Click()
    End
End Sub

Private Sub Command5_Click()
     i = 0
     Do While i < List1.ListCount
        If List1.Selected(i) = True Then
            List2.AddItem List1.List(i)
            List1.RemoveItem i
        Else
            i = i + 1
        End If
     Loop
     
        
End Sub

Private Sub Command6_Click()
    i = 0
     Do While i < List2.ListCount
        If List2.Selected(i) = True Then
            List1.AddItem List2.List(i)
            List2.RemoveItem i
        Else
            i = i + 1
        End If
     Loop
     
End Sub

Private Sub Command7_Click()
    For i = 0 To List1.ListCount - 1
        List2.AddItem List1.List(i)
    Next
    List1.Clear
End Sub

Private Sub Command8_Click()
    For i = 0 To List2.ListCount - 1
        List1.AddItem List2.List(i)
    Next
    List2.Clear
End Sub
