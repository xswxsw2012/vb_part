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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Form1.Picture = LoadPicture("C:\Users\Public\Pictures\Sample Pictures\psb.jpg")
    'Form1.Icon = LoadPicture("C:\Users\Public\Pictures\Sample Pictures\psb.jpg")
    Label1.Caption = "欢迎使用Visual Basic 6.0"
    Label1.BackStyle = 0 - transparent
    
End Sub
