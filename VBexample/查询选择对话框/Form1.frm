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
   Begin VB.CommandButton Command1 
      Caption         =   "��ѯ"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      Caption         =   "���༶"
      Height          =   300
      Left            =   2880
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "������"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "��ѧ��"
      Height          =   180
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "��ѡ���ѯ��ʽ"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Option1.Value = True Then
        yn = MsgBox("ѡ��ѧ�Ų�ѯ�룿", 67, "��ʾ")
        If yn = vbNo Then
            End
        ElseIf yn = vbYes Then
            tr = InputBox$("������ѧ�ţ�0~9���֣���", "����ѧ��")
                If Len(tr) = 10 Then
                    yn = MsgBox("�������ѧ����" & tr, 64, "��ʾ")
                Else
                    MsgBox "������󣡲��ܲ�ѯ��", 16, "�ر���ʾ"
                End If
        ElseIf yn = vbCancel Then
            MsgBox "��ѧ�Ų�ѯ������ȡ����", 48, "����"
        End If
    End If
    
    If Option2.Value = True Then
        tr = InputBox$("����������", "��������")
            yn = MsgBox("�������������" & tr, 64, "��ʾ")
    End If
    
    If Option3.Value = True Then
        tr = InputBox$("������༶���ƣ�", "����༶")
        yn = MsgBox("������İ༶��" & tr, 64, "��ʾ")
    End If
    
                    
End Sub

Private Sub Form_Load()
    Form1.Caption = "����Ի������Ϣ��"
End Sub
