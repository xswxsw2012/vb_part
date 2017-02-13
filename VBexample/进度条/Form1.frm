VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Left            =   6720
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "初始化数组"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   5880
      TabIndex        =   0
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s(0 To 2500) As String
Dim counter, n As Integer

Private Sub Command1_Click()
    ProgressBar1.Min = LBound(s)
    ProgressBar1.Max = UBound(s)
    ProgressBar1.Visible = True
    ProgressBar1.Value = ProgressBar1.Min
    counter = LBound(s)
    Timer1.Enabled = True
    n = Int(UBound(s) - LBound(s)) / 100
    Label2.Caption = "正在进行，请稍候！"
End Sub

Private Sub Form_Load()
    Timer1.Enabled = False
    Timer1.Interval = 10
End Sub

Private Sub Timer1_Timer()
    If counter <= UBound(s) Then
        s(counter) = Int(Rnd(1) * 1000)
        ProgressBar1.Value = counter
        Label1.Caption = "已完成" & Int(ProgressBar1.Value / n) & "%"
        counter = counter + 1
    End If
    If counter > UBound(s) Then
        Label2.Visible = False
    End If
    
End Sub
