VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "��Ϣϵͳ������"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.Menu browerlog 
      Caption         =   "�鿴��־"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()   '������
    nstr = "����ϵͳ"
    nstr = Format(Now, "yyyy-mm-dd hh:mm:ss") & nstr '������ֵΪϵͳʱ�������
    F = App.Path & "\" & "xsxx.log" '��ȡ�ļ���·�����ļ���
    H = FreeFile       '�ú������Ŀǰ��С��δʹ�õ��ļ���
    Open F For Append As #H   '��׷�ӷ�ʽ���ļ�
    Print #H, nstr       '��������ֵд���ļ�
    Close #H       '�ر��ļ�
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    nstr = "�˳�ϵͳ"
    nstr = Format(Now, "yyyy-mm-dd hh:mm:ss") & nstr '������ֵΪϵͳʱ�������
    F = App.Path & "\" & "xsxx.log"  '��ȡ�ļ���·�����ļ���
    H = FreeFile   '�ú������Ŀǰ��С��δʹ�õ��ļ���
    Open F For Append As #H        '��׷�ӷ�ʽ���ļ�
    Print #H, nstr        '��������ֵд���ļ�
    Close #H      '�ر��ļ�
End Sub
Private Sub browerlog_Click()
    Form1.Show
End Sub
