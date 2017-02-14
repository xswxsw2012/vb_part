VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "简易计算器"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdNum 
      Caption         =   "0"
      Height          =   375
      Index           =   9
      Left            =   1680
      TabIndex        =   17
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton CmdExpr 
      Caption         =   "+/-"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton ComEq 
      Caption         =   "="
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton CmdOp 
      Caption         =   "+"
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   14
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton CmdOp 
      Caption         =   "-"
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   13
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdOp 
      Caption         =   "*"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   12
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton CmdOp 
      Caption         =   "/"
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "9"
      Height          =   375
      Index           =   8
      Left            =   3480
      TabIndex        =   10
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "8"
      Height          =   375
      Index           =   7
      Left            =   2520
      TabIndex        =   9
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "7"
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "6"
      Height          =   375
      Index           =   5
      Left            =   3480
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "5"
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "3"
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "2"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "c"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox TxtShow 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As String
Dim opd As String
Dim b As Integer

Private Sub CmdExpr_Click()
    Minus TxtShow
    
End Sub

Private Sub CmdNum_Click(Index As Integer)
    If TxtShow.Text = "0" Or b = -1 Then
        TxtShow.Text = ""
        b = 0
    End If
    TxtShow.Text = TxtShow.Text + CmdNum(Index).Caption
    
End Sub

Private Sub ComEq_Click()
    TxtShow.Text = Operater(TxtShow.Text)
    b = 0
    op = ""
End Sub

Private Sub Command1_Click()
    TxtShow.Text = ""
End Sub

Private Sub Minus(Txt As Object)
Dim t As Integer
If Txt.Text = "0" Or b = -1 Then
Txt.Text = ""
b = 0
End If
t = Len(Txt.Text)
If Left(Txt.Text, 1) = "-" Then
    Txt.Text = Right(Txt.Text, t - 1)
Else
    Txt.Text = "-" & Txt.Text
End If
End Sub

Private Function Operater(s As String) As String
    Dim Value As Variant
    If op <> "" Then
        Select Case op
            Case "/"
                If Val(s) <> 0 Then
                    Value = Val(opd) / (Val(s))
                Else
                    MsgBox "除数不能是0", 0 + 48, "警告"
                End If
                
            Case "*"
                Value = Val(opd) * (Val(s))
            Case "+"
                Value = Val(opd) + (Val(s))
            Case "-"
                Value = Val(opd) - (Val(s))
        End Select
    End If
    Operater = Str(Value)
End Function

