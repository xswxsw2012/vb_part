VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " �鿴��־"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�鿴"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F As String
Dim H As Long

Private Sub Command1_Click()
    F = App.Path & "\" & "xsxx.log"
    H = FreeFile
    Open F For Input As #H   '��˳��ʽ���ļ�
    Text1.Text = ""
    Do Until EOF(1)     '�ļ�δ��β��
        Line Input #H, newline   '���ļ��е�һ�е�����newline��
        Text1.Text = Text1.Text + newline + Chr(13) + Chr(10) '������ֵ��ʾ���ı����ڣ�ÿ��β���ӻس����з���Chr(13)��ʾ�س���Chr(10)��ʾ���У��س��ص���ǰ�е����ף���������Ļ���������ǰ�����ݻᱻ��һ���ǡ�
    Loop       'ѭ��ֱ���ļ�β
    Close #H     '�ر��ļ�
End Sub

Private Sub Command2_Click()
    Unload Me     '�˳���ǰ���壬�����˳�����
End Sub
