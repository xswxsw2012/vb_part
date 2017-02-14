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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2400
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   600
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(10) As Integer
Dim b() As Integer
Dim c() As Integer
Dim m As String
Dim n As Integer
Dim s As String

Private Sub Command1_Click()
    For i = 1 To 10
        
        If a(i) > 0 Then
            m = m & " " & a(i)
            n = n + a(i)
        ElseIf a(i) < 0 Then
            k = k & " " & a(i)
            u = u + a(i)
        End If
    Next i
   
    Label2.Caption = m
    Label3.Caption = k
    Label4.Caption = n
    Label5.Caption = u
End Sub

Private Sub Command2_Click()
    Unload Form1
End Sub

Private Sub Form_Load()
        a(1) = -2
        a(2) = 73
        a(3) = 82
        a(4) = -76
        a(5) = -1
        a(6) = 24
        a(7) = 321
        a(8) = -25
        a(9) = 89
        a(10) = -20
        n = 0
  Frame1.Caption = "正负数"
  Frame2.Caption = "正数"
  Frame3.Caption = "负数"
  Frame4.Caption = "正数和"
  Frame5.Caption = "负数和"
  Command1.Caption = "开始"
  Command2.Caption = "结束"
  Form1.Caption = "正负数求和"
  Form1.Width = 5000
  Form1.Height = 6000
  
    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
For v = 0 To 10
    s = s & " " & a(v)  '显示数组数据
Next v
 Label1.Caption = s
  
End Sub
