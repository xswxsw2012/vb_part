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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   840
      Max             =   80
      Min             =   8
      TabIndex        =   5
      Top             =   3600
      Value           =   8
      Width           =   3975
   End
   Begin VB.Frame Frame3 
      Caption         =   "��ɫ"
      Height          =   1215
      Left            =   6120
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
      Begin VB.OptionButton Option8 
         Caption         =   "��ɫ"
         Height          =   180
         Left            =   1080
         TabIndex        =   19
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option7 
         Caption         =   "��ɫ"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "��ɫ"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "��ɫ"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "����"
      Height          =   1215
      Left            =   6120
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
      Begin VB.OptionButton Option4 
         Caption         =   "����"
         Height          =   180
         Left            =   960
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "��Բ"
         Height          =   180
         Left            =   960
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   975
      Left            =   6120
      TabIndex        =   1
      Top             =   240
      Width           =   1935
      Begin VB.CheckBox Check4 
         Caption         =   "ɾ����"
         Height          =   180
         Left            =   840
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         Caption         =   "�»���"
         Height          =   180
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "б��"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "�����С��8-80��"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    Text1.FontBold = Check1.Value
End Sub

Private Sub Check2_Click()
    Text1.FontItalic = Check2.Value
End Sub

Private Sub Check3_Click()
    Text1.FontUnderline = Check3.Value
End Sub

Private Sub Check4_Click()
    Text1.FontStrikethru = Check4.Value
End Sub

Private Sub Command1_Click()
    Text1.Text = ""
End Sub

Private Sub Command2_Click()
 End
End Sub

Private Sub Form_Load()
    ch = Chr(13) + Chr(10)
    Text1.Text = "��ǰ���¹�" & ch & "���ǵ���˪" & ch & "��ͷ������" & ch & "��ͷ˼����"
End Sub

Private Sub HScroll1_Change()      '�õ�����������ֵ
    Text1.FontSize = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()       '���ٹ������еĶ�̬�仯
    Text1.FontSize = HScroll1.Value
End Sub

Private Sub Option1_Click()
    Text1.FontName = Option1.Caption
End Sub

Private Sub Option2_Click()
    Text1.FontName = Option2.Caption
End Sub

Private Sub Option3_Click()
    Text1.FontName = Option3.Caption
End Sub

Private Sub Option4_Click()
    Text1.FontName = Option4.Caption
End Sub

Private Sub Option5_Click()
    Text1.ForeColor = vbRed
End Sub

Private Sub Option6_Click()
    Text1.ForeColor = vbBlue
End Sub

Private Sub Option7_Click()
    Text1.ForeColor = vbGreen
End Sub

Private Sub Option8_Click()
    Text1.ForeColor = vbBlack
End Sub
