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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command4 
      Caption         =   "ÍË³ö"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "À¶"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "»Æ"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ºì"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Label1.ForeColor = RGB(255, 0, 0)
End Sub


Private Sub Command2_Click()
    Label1.ForeColor = RGB(0, 0, 255)
End Sub

Private Sub Command3_Click()
   Label1.ForeColor = RGB(0, 255, 0)
End Sub

Private Sub Command4_Click()
    Unload Form1
    
End Sub


Private Sub Form_Load()
    Label1.Caption = "»¶Ó­Ê¹ÓÃVisual Basic 6.0"
End Sub

