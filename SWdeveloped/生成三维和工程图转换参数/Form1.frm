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
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim swApp As Object           '��������λ��û��������Щûר�Ŷ���
Dim part As Object
Dim boolstatus As Boolean
'����ӿ�
Set swApp = CreateObject("SldWorks.Application")   '��ȡSolidWorks����ӿڣ����ڵ���û�а�װsw�������ʱ����
Set part = swApp.NewPart()         '�½�һ������ļ�
swApp.Visible = True

Dim myModelView As Object     '����һ������
Set myModelView = part.ActiveView       '������ֵ
myModelView.FrameState = swWindowState_e.swWindowMaximized
boolstatus = part.Extension.SelectByID2("ǰ�ӻ�׼��", "PLANE", 0, 0, 0, False, 0, Nothing, 0)  '����ǰ�ӻ�׼��
part.SketchManager.InsertSketch True    '��ʼ���Ʋ�ͼ
part.ClearSelection2 True              '���ѡ���б�

Dim skSegment As Object
Set skSegment = part.SketchManager.CreateCircle(0#, 0#, 0#, 0.017934, 0.015738, 0#)      '��Բ
part.ClearSelection2 True
boolstatus = part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 2.23259369346084E-02, 6.95398035668131E-03, 0, False, 0, Nothing, 0)

Dim myDisplayDim As Object
Set myDisplayDim = part.AddDimension2(3.73318945463944E-02, 5.85598345825794E-03, 0)    '��ӳߴ�
part.ClearSelection2 True
boolstatus = part.Extension.SelectByID2("D1@��ͼ1@���3.SLDPRT", "DIMENSION", 3.73318945463944E-02, 5.85598345825794E-03, 0, False, 0, Nothing, 0)   'ѡ���ͼ

Dim myDimension As Object
Set myDimension = part.Parameter("D1@��ͼ1")
myDimension.SystemValue = 0.07                 '����Բ�ĳߴ�ֱ������ʱȡֵΪ0.06-0.07����Ϊȡֵ���������ͼֽ�����Զ��仯��Ӱ����㣬Ҳ��̫����ʵ������
part.ShowNamedView2 "*���¶��Ƚ����", 8
part.ClearSelection2 True
boolstatus = part.Extension.SelectByID2("D1@��ͼ1@���3.SLDPRT", "DIMENSION", 3.73318945463944E-02, 5.85598345825794E-03, 0, False, 0, Nothing, 0)

Dim myFeature As Object        'feature�˴���������������
Set myFeature = part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.02, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, True, True, True, 0, 0, False)  '��6���������������ĺ�ȣ���ʱȡֵ0.01-0.02��̫���̫С��Ӱ���Ӱ��ͼֽ�����������ڼ���
part.SelectionManager.EnableContourSelection = False
part.ShowNamedView2 "*�����", 7
part.ShowNamedView2 "*�����", 7
longstatus = part.SaveAs3(App.Path & "\" & "Բ��.SLDPRT", 0, 2)     '�������ͼ�������ںͳ���һ����ļ�����



Dim longwarnings As Long
Set moddoc = swApp.NewDocument(App.Path & "\" & "��׼ͼֽ\gb_a0.drwdot", 12, 0.841, 1.189)  '��1��������ʾ����ģ���λ�ã�����3��������ʾ12��ͼֽ����841mm���1189mm���½�����ͼ�����ⲿ�ļ���ʹ��sw�Դ�ģ��ͼ����ԭ��δ֪
Set moddoc = swApp.ActiveDoc

Dim myView As Object
'Set myView = moddoc.CreateDrawViewFromModelView3(App.Path & "\" & "Բ��.SLDPRT", "*ǰ��", 0.265882219679634, 0.61352585812357, 0) '��1��������ʾ������ά���ͼ��λ�ã���2��������ʾ���ͼ��ǰ��ͼ��Ϊ�½�����ͼ������ͼ������3��������ʾ���������ͼ�ڹ���ͼ��λ�ã��½�����ͼ������ԭ��Ϊͼֽ�����½ǣ�����ͼ����Ļ���Ϊ�����ͼ�����ģ�Ҫ����ͼֽ����ͼ�Ĵ�Сȷ������������
Set myView = moddoc.CreateDrawViewFromModelView3(App.Path & "\" & "Բ��.SLDPRT", "*ǰ��", 0.26, 0.613, 0)
boolstatus = moddoc.Extension.SelectByID2("����ͼ��ͼ1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = moddoc.ActivateView("����ͼ��ͼ1")

'Set myView = moddoc.CreateUnfoldedViewAt3(0.265882219679634, 0.202936956521739, 0, False)  '���븩��ͼ��ǰ3��������ʾ����ֵ�����һ������false��ʾ�븸��ͼ���˴�������ͼ������
Set myView = moddoc.CreateUnfoldedViewAt3(0.26, 0.2, 0, False)
moddoc.ClearSelection2 True
boolstatus = moddoc.Extension.SelectByID2("����ͼ��ͼ1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
moddoc.ClearSelection2 True
'boolstatus = moddoc.Extension.SelectByID2("", "EDGE", 0.411469519450801, 0.656056979407828, 400.005, False, 0, Nothing, 0)
boolstatus = moddoc.Extension.SelectByID2("", "EDGE", 0.26 + 0.035 * 5, 0.613, 400.005, False, 0, Nothing, 0)

Dim myDisplayDim1 As Object
'Set myDisplayDim1 = moddoc.AddDimension2(0.265882219679634, 0.61352585812357, 0)     'ֱ����עλ�õ����꣨��ȷ��mm����Ȼ���ܲ���ʾ��ע����ע����ӣ�����ͼ�ߴ�
Set myDisplayDim1 = moddoc.AddDimension2(0.26 + 0.01 * 5, 0.613 + 0.01 * 5, 0)
moddoc.ClearSelection2 True
boolstatus = moddoc.ActivateSheet("ͼֽ1")
boolstatus = moddoc.ActivateView("����ͼ��ͼ2")
'boolstatus = moddoc.Extension.SelectByID2("", "SILHOUETTE", 0.41474114416476, 0.212751830666409, 400, False, 0, Nothing, 0)
boolstatus = moddoc.Extension.SelectByID2("", "SILHOUETTE", 0.41474114416476 + 0.005 * 5, 0.212751830666409, 400, False, 0, Nothing, 0) '��ע�ߴ�ʱ��ѡ���ұ߱���
'boolstatus = moddoc.Extension.SelectByID2("", "EDGE", 0.261, 0.21, 400, False, 0, Nothing, 0)
'boolstatus = moddoc.Extension.SelectByID2("", "EDGE", 0.261, 0.21 - 0.01 * 5, 400, True, 0, Nothing, 0)
'Set myDisplayDim1 = moddoc.AddDimension2(0.44254995423341, 0.206208581235698, 0)     '��ע����ӣ�����ͼ�ߴ�
Set myDisplayDim1 = moddoc.AddDimension2(0.26, 0.2, 0)   '���쳤�ȱ�עλ�õ����꣬����5����Ϊͼ����5:1
'Set myDisplayDim1 = moddoc.AddDimension2(0.26, 0.2, 0)
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
longstatus = moddoc.SaveAs3(App.Path & "\" & "Բ��.DWG", 0, 0)   '���湤��ͼ


End Sub

Private Sub Command2_Click()
    Unload Form1
End Sub

Private Sub Form_Load()
    Command1.Caption = "��ʼ"
    Command2.Caption = "����"
    Label1.Caption = "ע���ļ��������Զ����ɵ���άͼ�͹���ͼ"
    Form1.Caption = "������άͼ�͹���ͼ"
End Sub
