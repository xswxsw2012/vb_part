VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   18000
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   2055
      Left            =   9480
      TabIndex        =   1
      Top             =   5280
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2055
      Left            =   9120
      TabIndex        =   0
      Top             =   2160
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim swApp As Object
Dim moddoc As ModelDoc2
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim drawdoc As DrawingDoc
Dim View As Object



Set swApp = CreateObject("Sldworks.Application") '���Ӧ�ó������
swApp.Visible = True

Set moddoc = swApp.OpenDoc6(App.Path & "\" & "С����.SLDPRT", 1, 0, "", longstatus, longwarnings)
Set moddoc = swApp.ActivateDoc2("С����.SLDPRT", False, longstatus)

Set drawdoc = swApp.NewDrawing2(13, App.Path & "\" & "�����ļ���\gb_a0.drwdot", 2, 0.2794, 0.4318)


Set View = drawdoc.CreateDrawViewFromModelView2(App.Path & "\" & "С����.SLDPRT", "*ǰ��", 0.19198374340949, 0.656111142355, 0)
boolstatus = moddoc.Extension.SelectByID2("������ͼ1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)

Set View = drawdoc.CreateUnfoldedViewAt3(0.75, 0.656111142355, 0, 0)
moddoc.ClearSelection2 True
boolstatus = moddoc.Extension.SelectByID2("������ͼ1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
moddoc.ClearSelection2 True

Set moddoc = swApp.ActiveDoc
boolstatus = moddoc.Extension.SelectByID2("������ͼ1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)

Set View = drawdoc.CreateUnfoldedViewAt3(0.19198374340949, 0.2821920652174, 0, 0)
boolstatus = moddoc.Extension.SelectByID2("������ͼ1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
moddoc.ClearSelection2 True    '��ͼָֽ��λ�����ɱ�׼����ͼ

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Me.Command1.Caption = "����"
Me.Command2.Caption = "����"
End Sub
