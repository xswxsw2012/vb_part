VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   5535
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   1815
      Left            =   7080
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   4200
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   720
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()        '"ȷ��"��ť�ĵ����¼�
    Text2.Text = Text1.Text        '��text1��text���Ե�ֵ����text2��text����
    Text1.Text = ""
    Label1.Caption = "����ı���������Ѹ��Ƶ��ұߵ��ı���"
    Command2.Enabled = True
    
End Sub


Private Sub Command2_Click()
Unload Form1

End Sub

Private Sub Form_Load()
    Label1.FontName = "����"
    Label1.FontSize = 12
    Label1.ForeColor = vbRed
    Label1.Caption = "��������ı����������֣�Ȼ�󵥻���ȷ�ϡ���ť"
    Text1.Text = ""
    Text2.Text = ""
    Text2.Locked = True             '�ı���text2�����������ܽ������ֱ༭
    
End Sub


Private Sub Text1_GotFocus()        '�ı���text1��ý����¼�
    Label1.Caption = "��������ı����������֣�Ȼ�󵥻���ȷ�ϰ�ť��"
    Text2.Text = ""
    Command2.Enabled = False        '��ʱ���˳�����ť������
    
    

End Sub
