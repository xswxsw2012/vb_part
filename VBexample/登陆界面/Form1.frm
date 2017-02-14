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
      Caption         =   "退出"
      Height          =   420
      Left            =   2880
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "密码："
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "用户名："
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "欢迎登陆"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Label5.Caption = Text1.Text
    Text1.Text = ""
    Label4.Caption = Text2.Text
    Text2.Text = ""
    Label1.Caption = "您的密码是："
    Label6.Caption = "欢迎"
    Command1.Enabled = False
    
    
End Sub

Private Sub Command2_Click()
    Unload Form1
    
End Sub
