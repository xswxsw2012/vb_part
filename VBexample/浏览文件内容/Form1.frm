VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "文件系统控件的使用"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   3840
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   480
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   1770
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   1710
      Left            =   1080
      Pattern         =   "*.txt;*.jpg;*.gif;*.bmp"
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As String

Private Sub Dir1_Change()
    File1.Pattern = "*.txt;*.jpg;*.gif;*.bmp"       '设置可显示的文本模式
    File1.Path = Dir1.Path                          '将目录列表框的路径赋给文件列表框
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive                        '将驱动器的Drive属性值赋给目录路径
End Sub

Private Sub File1_Click()
     p = File1.Path & "\" & File1.FileName            '文件列表框内选择的文件的路径
     If LCase$(Right(p, 3)) = "txt" Then            '判断字符串尾部的3个字符
        Image1.Visible = False               '图片字符串尾部的3个字符
        Text1.Visible = True               '文本框显示
        Open p For Input As #1             '打开文件
        Text1.Text = ""
        Do Until EOF(1)
            Line Input #1, newline          '逐行读取文件到变量newline中
            Text1.Text = Text1.Text + newline + Chr(13) + Chr(10)
        Loop
        Close #1
    Else
        Image1.Visible = True              '否则，图像框显现
        Text1.Visible = False             '文本框隐藏
        Image1.Picture = LoadPicture(p)    '装载图片
    End If
End Sub
