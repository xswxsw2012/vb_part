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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim swApp As Object
Dim a, b As Integer
Dim part As Object
Dim myModelView As Object
Dim skSegment As Object
Dim myDisplayDim As Object
Dim myDimension As Object
Dim myView As Object
Dim myFeature As Object
Dim myDisplayDim1 As Object
Dim moddoc As Object
Dim boolstatus As Boolean

Private Sub Command1_Click()
a = Val(Text1.Text)    'Բ��ֱ��
b = Val(Text2.Text)    'Բ�̺��

If a >= 60 And a <= 70 And b >= 10 And b <= 20 Then
Set swApp = CreateObject("SldWorks.Application")     '����sw
swApp.Visible = True
Set part = swApp.NewDocument(App.Path & "\" & "��׼ͼֽ" & "\" & "gb_part.prtdot", 0, 0, 0)
Set myModelView = part.ActiveView
boolstatus = part.Extension.SelectByID2("ǰ�ӻ�׼��", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
part.SketchManager.InsertSketch True    '��ʼ���Ʋ�ͼ
part.ClearSelection2 True              '���ѡ���б�

Set skSegment = part.SketchManager.CreateCircle(0#, 0#, 0#, 0.017934, 0.015738, 0#)      '��Բ
part.ClearSelection2 True
boolstatus = part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 2.23259369346084E-02, 6.95398035668131E-03, 0, False, 0, Nothing, 0)

Set myDisplayDim = part.AddDimension2(3.73318945463944E-02, 5.85598345825794E-03, 0)    '��ע��ͼ�ߴ�(Ĭ�ϵ�)
part.ClearSelection2 True
boolstatus = part.Extension.SelectByID2("D1@��ͼ1@���3.SLDPRT", "DIMENSION", 3.73318945463944E-02, 5.85598345825794E-03, 0, False, 0, Nothing, 0)   'ѡ���ͼ

Set myDimension = part.Parameter("D1@��ͼ1")
myDimension.SystemValue = a / 1000               '����Բ�ĳߴ�ֱ������ʱȡֵΪ0.06-0.07
part.ShowNamedView2 "*���¶��Ƚ����", 8
part.ClearSelection2 True
boolstatus = part.Extension.SelectByID2("D1@��ͼ1@���3.SLDPRT", "DIMENSION", 3.73318945463944E-02, 5.85598345825794E-03, 0, False, 0, Nothing, 0)

Set myFeature = part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, b / 1000, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, True, True, True, 0, 0, False) '��6���������������ĺ�ȣ���ʱȡֵ0.01-0.02��̫���̫С��Ӱ���Ӱ��ͼֽ�����������ڼ���
part.SelectionManager.EnableContourSelection = False
part.ShowNamedView2 "*�����", 7
part.ShowNamedView2 "*�����", 7
longstatus = part.SaveAs3(App.Path & "\" & "Բ��.SLDPRT", 0, 2)     '�������ͼ



Dim longwarnings As Long
Set moddoc = swApp.NewDocument(App.Path & "\" & "��׼ͼֽ\gb_a0.drwdot", 12, 0.841, 1.189)
Set moddoc = swApp.ActiveDoc

Set myView = moddoc.CreateDrawViewFromModelView3(App.Path & "\" & "Բ��.SLDPRT", "*ǰ��", 0.26, 0.613, 0)
boolstatus = moddoc.Extension.SelectByID2("����ͼ��ͼ1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = moddoc.ActivateView("����ͼ��ͼ1")

Set myView = moddoc.CreateUnfoldedViewAt3(0.26, 0.2, 0, False)
moddoc.ClearSelection2 True
boolstatus = moddoc.Extension.SelectByID2("����ͼ��ͼ1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
moddoc.ClearSelection2 True
boolstatus = moddoc.Extension.SelectByID2("", "EDGE", 0.26 + a / 1000 / 2 * 5, 0.613, 400.005, False, 0, Nothing, 0) 'ѡ��������λ��

Set myDisplayDim1 = moddoc.AddDimension2(0.26 + 0.01 * 5, 0.613 + 0.01 * 5, 0)   '��עλ��
moddoc.ClearSelection2 True
boolstatus = moddoc.ActivateSheet("ͼֽ1")
boolstatus = moddoc.ActivateView("����ͼ��ͼ2")
boolstatus = moddoc.Extension.SelectByID2("", "SILHOUETTE", 0.26 + a / 1000 / 2 * 5, 0.2 + b / 1000 / 2 * 5, 400, False, 0, Nothing, 0)
Set myDisplayDim1 = moddoc.AddDimension2(0.26 + (a / 1000 / 2 + 0.005) * 5, 0.2, 0)

boolstatus = moddoc.Extension.SelectByID2("RD1@����ͼ��ͼ1", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = moddoc.ActivateSheet("ͼֽ1")
boolstatus = moddoc.ActivateView("����ͼ��ͼ2")
boolstatus = moddoc.ActivateSheet("ͼֽ1")
moddoc.ClearSelection2 True
moddoc.ViewZoomtofit2
moddoc.SheetPrevious
moddoc.ViewZoomTo2 0, 0, 0, 0.1, 0.1, 0.1
moddoc.ViewZoomTo2 0, 0, 0, 0.1, 0.1, 0.1
moddoc.ViewZoomTo2 0, 0, 0, 0.1, 0.1, 0.1
moddoc.ViewZoomtofit2
longstatus = moddoc.SaveAs3(App.Path & "\" & "Բ��.DWG", 0, 0)
Else
    MsgBox ("ֱ������60-70���������10-20")
End If
End Sub

Private Sub Command2_Click()
    Unload Form1
End Sub

Private Sub Form_Load()
    Form1.Width = 10000
    Form1.Height = 5000
    Label1.Caption = "ֱ��"
    Label2.Caption = "���"
    Frame1.Caption = "Բ�̲�������λ��mm��"
    Command1.Caption = "��ʼ"
    Command2.Caption = "����"
    Text1.Text = "65"
    Text2.Text = "15"
End Sub
