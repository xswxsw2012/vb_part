VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
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
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��������"
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�����ļ�"
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���ļ�"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   2640
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "��ѡ���Ҽ���ť���в���"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yn As Integer  '����һ�����ͱ���

Private Sub Command1_Click()
    CommonDialog1.Filter = "�ĵ�(*.doc;*.rtf;*.txt)|*.doc;*.ref;*.txt|�����ļ�(*.*)|*.*" '�����ļ��б��������ʾ�ļ�������
    CommonDialog1.Action = 1    '���á��򿪡��Ի���
    Label1.Caption = "��" + CommonDialog1.FileName  '���ñ�ǩ����
    Text1.Text = ""   '�����ı����ʼֵ
    Open CommonDialog1.FileName For Input As #1  '��ѡ����ļ�
    Do While Not EOF(1)
        Line Input #1, inputdata   '��һ������
        Text1.Text = Text1.Text + inputdata + vbCrLf
    Loop
    Close #1        '�ر��ļ�
End Sub

Private Sub Command2_Click()
    CommonDialog1.FileName = "default.txt"  '�����ļ���Ĭ���ļ���
    CommonDialog1.DefaultExt = "txt" 'Ĭ�ϵ���չ��
    CommonDialog1.Action = 2   '���á����Ϊ���Ի���
    Label1.Caption = "����" + CommonDialog1.FileName
    Open CommonDialog1.FileName For Output As #1  '���ļ�д������
    Print #1, Text1.Text    '���ı����ڵ��ı�д���ļ�
    Close #1    '�ر��ļ�
End Sub

Private Sub Command3_Click()
    CommonDialog1.Flags = 3   '������ʾ����Ϊ��Ļ������ӡ���������
    CommonDialog1.Action = 4  '���á����塱�Ի���
    Label1.Caption = "Ϊ�ļ�" + CommonDialog1.FileName + "��������"
    Text1.FontName = CommonDialog1.FontName '�����ı�����
    Text1.FontSize = CommonDialog1.FontSize '�����ı��ֺ�
    Text1.FontBold = CommonDialog1.FontBold '�����ı�����
    Text1.FontItalic = CommonDialog1.FontItalic '�����ı�б��
    Text1.FontStrikethru = CommonDialog1.FontStrikethru  '�����ı�ɾ����
    Text1.FontUnderline = CommonDialog1.FontUnderline '�����ı��»���
    Text1.ForeColor = CommonDialog1.Color  '�����ı���ɫ
End Sub

Private Sub Command4_Click()
    yn = MsgBox("���˳�֮ǰ�����ļ���������", 4, "��ʾ")
    If yn = 6 Then
        End
    End If
End Sub
