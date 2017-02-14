VERSION 5.00
Begin VB.Form Form1 
   Caption         =   $"Form1.frx":0000
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   5535
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   5535
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Dim a(10) As Integer

Private Sub Command1_Click()
    Dim i, j, min, t As Integer
    For i = 1 To 9
        min = i
        For j = i + 1 To 10
            If a(j) < a(min) Then min = j
        Next j
        t = a(i): a(i) = a(min): a(min) = t
    Next i
    s = ""
    For i = 1 To 10
        s = s & " " & a(i)
    Next i
    Label2.Caption = s
End Sub

Private Sub Command2_Click()
    Unload Form1
End Sub

Private Sub Form_Load()
    Randomize
    For i = 1 To 10
        a(i) = Int((100 * Rnd) + 1)
        s = s & " " & a(i)
    Next i
    Label1.Caption = s
    Label2.Caption = ""
    Frame1.Caption = "ÅÅÐòÇ°"
    Frame2.Caption = "ÅÅÐòºó"
    Command1.Caption = "ÅÅÐò"
    Command2.Caption = "ÍË³ö"
    Form1.Caption = "ÅÅÐò"
End Sub
