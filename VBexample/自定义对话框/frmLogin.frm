VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "�û�����(&U):"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "����(&P):"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit    'ǿ����ʾ����ģ���е����б���
Public i As Integer
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    '����ȫ�ֱ���Ϊ false
    '����ʾʧ�ܵĵ�¼
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If txtUserName = "vb" Then
        If i < 3 Then
        '�����ȷ������
        If txtPassword = "123456" Then
        '������������ﴫ��
        '�ɹ��� calling ����
        '����ȫ�ֱ���ʱ�����׵�
        LoginSucceeded = True
        MsgBox "��ȷ�����룬��ӭ����", , "��¼"
        Me.Hide
        Load frmSplash
        frmSplash.Show
    Else
        MsgBox "��Ч�����룬������!", , "��¼"
        txtPassword = ""
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        i = i + 1
    End If
    Else
        MsgBox "������������ѵ������ܵ�¼��", , "��ʾ"
        LoginSucceeded = False
        Me.Hide
    End If
    Else
        MsgBox "�û���������������룡", , "��ʾ"
        txtUserName = ""
        txtUserName.SetFocus
    End If
End Sub
Private Sub frmlogin_load()
    i = 1
End Sub
End Sub
