VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��ʾ���ʽ����"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame4 
      Caption         =   "�߼����ʽ"
      Height          =   1695
      Left            =   6480
      TabIndex        =   8
      Top             =   2280
      Width           =   3615
      Begin VB.Label Label5 
         Height          =   1215
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "��ϵ���ʽ"
      Height          =   1935
      Left            =   6480
      TabIndex        =   7
      Top             =   240
      Width           =   3375
      Begin VB.Label Label4 
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�������ʽ"
      Height          =   1935
      Left            =   2880
      TabIndex        =   6
      Top             =   2040
      Width           =   3375
      Begin VB.Label Label3 
         Height          =   1455
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   1575
      Left            =   2880
      TabIndex        =   5
      Top             =   360
      Width           =   3255
      Begin VB.Label Label2 
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʾ"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "������������"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y As Integer

Private Sub Command1_Click()
    Label3.Caption = "x+y��ֵ�ǣ�" & x + y
    Label3.Caption = Label3.Caption & Chr(13) & "x-y��ֵ�ǣ�" & x - y
    Label3.Caption = Label3.Caption & Chr(13) & "x*y��ֵ�ǣ�" & x * y
    Label3.Caption = Label3.Caption & Chr(13) & "x/y��ֵ�ǣ�" & x / y
    Label3.Caption = Label3.Caption & Chr(13) & "x\y��ֵ�ǣ�" & x \ y
    Label3.Caption = Label3.Caption & Chr(13) & "xMody��ֵ�ǣ�" & x Mod y
    Label3.Caption = Label3.Caption & Chr(13) & "x^3��ֵ�ǣ�" & x ^ 3
    
    Label4.Caption = "x=y��ֵ�ǣ�" & CStr(x = y)
    Label4.Caption = Label4.Caption & Chr(13) & "x>y��ֵ�ǣ�" & CStr(x > y)
    Label4.Caption = Label4.Caption & Chr(13) & "x>=y��ֵ�ǣ�" & CStr(x >= y)
    Label4.Caption = Label4.Caption & Chr(13) & "x<y��ֵ�ǣ�" & CStr(x < y)
    Label4.Caption = Label4.Caption & Chr(13) & "x<=y��ֵ�ǣ�" & CStr(x <= y)
    Label4.Caption = Label4.Caption & Chr(13) & "x<>y��ֵ�ǣ�" & CStr(x <> y)
    
    Label2.Caption = "x=" & x
    Label2.Caption = Label2.Caption & Chr(13) & "y=" & y
    Label2.Caption = Label2.Caption & Chr(13) & "a=(x+y)>100"
    Label2.Caption = Label2.Caption & Chr(13) & "a=(x-y)<10"
    
    a = x + y > 100
    b = x - y < 10
    
    Label5.Caption = "a��ֵ�ǣ�" & CStr(a)
    Label5.Caption = Label5.Caption & Chr(13) & "b��ֵ�ǣ�" & CStr(b)
    Label5.Caption = Label5.Caption & Chr(13) & "Not a��ֵ�ǣ�" & CStr(Not a)
    Label5.Caption = Label5.Caption & Chr(13) & "Not b��ֵ�ǣ�" & CStr(Not b)
    Label5.Caption = Label5.Caption & Chr(13) & "a And b��ֵ�ǣ�" & CStr(a And b)
    Label5.Caption = Label5.Caption & Chr(13) & "a Or b��ֵ�ǣ�" & CStr(a Or b)
    
End Sub

Private Sub Command2_Click()
    Unload Form1
    
End Sub

Private Sub Text1_Change()
    x = CInt(Text1.Text)
End Sub

Private Sub Text2_Change()
    y = CInt(Text2.Text)
    
End Sub
