VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   7080
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "跳动区域"
      Height          =   2535
      Left            =   4440
      TabIndex        =   2
      Top             =   840
      Width           =   2055
      Begin VB.Shape Shape2 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   720
         Shape           =   3  'Circle
         Top             =   1560
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   720
         Shape           =   3  'Circle
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "选择跳动时间"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer

Private Sub Command1_Click()
    t = Val(Combo1.Text)
    Timer1.Enabled = True
    Timer1.Interval = 200
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 10 To 100
        Combo1.AddItem i
    Next
    Shape1.Visible = False
    Shape2.Visible = False
    Combo1.Text = 10
End Sub

Private Sub Timer1_Timer()
    t = t - 1
    Label1.Caption = "倒计时：" & t & "秒"
    If t Mod 2 = 0 Then
        Shape1.Visible = True
        Shape2.Visible = False
    Else
        Shape1.Visible = False
        Shape2.Visible = True
    End If
    If t = 0 Then
        Timer1.Enabled = False
        t = MsgBox("时间到！", , "消息框")
        Shape1.Visible = False
        Shape2.Visible = False
    End If
End Sub
