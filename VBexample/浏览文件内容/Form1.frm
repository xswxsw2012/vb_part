VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ļ�ϵͳ�ؼ���ʹ��"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   3840
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   480
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   1770
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   1710
      Left            =   1080
      Pattern         =   "*.txt;*.jpg;*.gif;*.bmp"
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As String

Private Sub Dir1_Change()
    File1.Pattern = "*.txt;*.jpg;*.gif;*.bmp"       '���ÿ���ʾ���ı�ģʽ
    File1.Path = Dir1.Path                          '��Ŀ¼�б���·�������ļ��б��
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive                        '����������Drive����ֵ����Ŀ¼·��
End Sub

Private Sub File1_Click()
     p = File1.Path & "\" & File1.FileName            '�ļ��б����ѡ����ļ���·��
     If LCase$(Right(p, 3)) = "txt" Then            '�ж��ַ���β����3���ַ�
        Image1.Visible = False               'ͼƬ�ַ���β����3���ַ�
        Text1.Visible = True               '�ı�����ʾ
        Open p For Input As #1             '���ļ�
        Text1.Text = ""
        Do Until EOF(1)
            Line Input #1, newline          '���ж�ȡ�ļ�������newline��
            Text1.Text = Text1.Text + newline + Chr(13) + Chr(10)
        Loop
        Close #1
    Else
        Image1.Visible = True              '����ͼ�������
        Text1.Visible = False             '�ı�������
        Image1.Picture = LoadPicture(p)    'װ��ͼƬ
    End If
End Sub
