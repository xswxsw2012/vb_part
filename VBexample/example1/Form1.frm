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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()  '���弤���¼�
    Form1.Print                'Print �����޲���ʱ���һ����
    Form1.Print Tab(5); "����ѧϰ Visual Basic��"  'Tab��n������n���ַ���λ��������ַ���
End Sub


Private Sub Form_Click()  '���嵥���¼�
Form1.Cls '���������������
Form1.BackColor = RGB(0, 255, 0) '��ʾ��ɫ����ɫ�ķ�ֵΪ0�����Ϊ��ɫ

Form1.ForeColor = RGB(255, 0, 0)
Form1.FontName = "����"
Form1.Print Chr(13); Tab(5); "����ѧϰ Visual Basic��" 'Chr(13)��ʾ�Ȼ��������
End Sub


Private Sub Form_DblClick()      'ж�ش����¼�Form1
    Unload Form1
End Sub


Private Sub Form_Load()       '���������¼�
    Form1.ForeColor = RGB(0, 255, 0)
End Sub
