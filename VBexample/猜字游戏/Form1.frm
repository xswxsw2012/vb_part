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
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r, s As Integer         '���崰�弶��ģ�鹲��ı���

Private Sub Command1_Click()
    Label1.Caption = "������һ��100���ڵ�������"
    Randomize                       '�����������������ʼ������
    r = Int((100 * Rnd) + 1)        '�������100���ڵ�������
    s = 1
    Text1.Locked = False            '�����ı���Ϊ�ɱ༭״̬
    Command1.Enabled = False
    Command2.Enabled = True
    Text1.Text = ""
    Text1.SetFocus                  '�����ı��򽹵�
End Sub

Private Sub Command2_Click()
    If Text1.Text = "" Or (Not IsNumeric(Text1.Text)) Then      'IsNumeric�����ж��Ƿ�Ϊ�����ַ���
        Label1.Caption = "���ַ���������ַ�������������"
        Text1.Text = ""                                         '����ı���
    ElseIf Val(Text1.Text) > r Then                             'Val�����������ַ���ת��������
        Label1.Caption = Text1.Text & "���ˣ��Ѳ���" & s & "��"
        s = s + 1
        Text1.Text = ""
    ElseIf Val(Text1.Text) < r Then
        Label1.Caption = Text1.Text & "С�ˣ��Ѳ���" & s & "��"
        s = s + 1
        Text1.Text = ""
    Else
        Label1.Caption = "��ϲ������ˣ�������" & s & "��"
        Text1.Locked = True
        Command1.Enabled = True
        Command2.Enabled = False
    End If
    Text1.SetFocus             '�����ı��򽹵�
End Sub

Private Sub Command3_Click()
    Unload Form1
    
End Sub

Private Sub Form_Load()
    Label1.Caption = "�뵥������ʼ����ť������Ϸ"
    Text1.Text = ""
    Command1.Enabled = True
    Command2.Enabled = False
    
End Sub
