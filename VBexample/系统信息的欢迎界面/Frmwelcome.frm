VERSION 5.00
Begin VB.Form Frmwelcome 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7290
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   20.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   6000
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6000
      Top             =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "ѧ����Ϣ����ϵͳ��ӭ�㣡"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4320
      TabIndex        =   0
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "Frmwelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim j As Integer   '���õ粨���͵�ѭ������
Dim r As Byte, g As Byte, b As Byte
Dim i As Integer    '��������ֱ�ߵ�ѭ������

Private Sub Form_Click()
    Timer1.Interval = 0   '���ö�ʱ����ʱ����Ϊ0
    Cls                    '�����Ļ
    Frmwelcome.Hide
    
End Sub

Private Sub Form_Load()
    BackColor = RGB(0, 0, 0)    'ʵ�������п����ñ���ɫΪ��ɫ
                                '���������
    j = 0
    DrawWidth = 1               '�����߿�
    PSet (ScaleWidth / 3, 1000) '������֮һ��Ⱥ͸߶�1000����Բ
    ForeColor = RGB(0, 255, 0)   '����ǰ��ɫΪ��ɫ
    For i = 1 To 50 Step 5        '����ֱ�߿��ֵΪѭ������
        DrawWidth = i            '����ֱ�߿��ֵΪѭ������
        Line -Step(0, ScaleHeight / 10) '�ӵ�ǰλ�ð�����scaleHeight/10����
    Next i
    
                            '�������ϲ���ͻ��
    DrawWidth = 1           '���������߿�
    FillStyle = 6           '���������ʽΪʮ����
    FillColor = ForeColor   '��������ߵ���ɫΪǰ��ɫ
    Circle (ScaleWidth / 3, 2000), 300, , , , 0.5 '��������Բ
    
End Sub

Private Sub Timer1_Timer()
   
    FillStyle = 1     'ͼ�η������ɵ�Բ�򷽿��ģʽΪ͸��
    DrawStyle = 2     '�����������ʽΪ����
    r = 255 * Rnd       '���ɵĺ�ɫ�������
    g = 255 * Rnd       '���ɵ���ɫ�������
    b = 255 * Rnd       '���ɵ���ɫ�������
    j = j + 1
    Circle (ScaleWidth / 3, 1000), 300 * j, RGB(r, g, b) '�Զ�ʱ���涨��ʱ�������뾶���300����ɫ�����Բ
    
    If j = 10 Then j = 0
End Sub

Private Sub Timer2_Timer()
    For i = 0 To 10000 Step 50
    Label1.Top = (Label1.Top - 1)  'ʹ�ؼ�λ�÷����仯
    Label1.ForeColor = RGB(r, g, b) 'ʹ�ؼ�ǰ��ɫ�����仯
    If Label1.Top = -Label1.Height Then Label1.Top = ScaleHeight
    Next i
End Sub
