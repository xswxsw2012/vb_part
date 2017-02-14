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
   StartUpPosition =   3  '窗口缺省
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
Dim boolstatus As Boolean


Private Sub Command1_Click()
a = Val(Text1.Text)    '圆盘直径
b = Val(Text2.Text)    '圆盘厚度
If a >= 60 And a <= 70 And b >= 10 And b <= 20 Then
Set swApp = CreateObject("SldWorks.Application")     '启动sw
Set part = swApp.NewPart()
swApp.Visible = True

Dim myModelView As Object
Set myModelView = part.ActiveView
boolstatus = part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)  '设置前视基准面
part.SketchManager.InsertSketch True    '开始绘制草图
part.ClearSelection2 True              '清除选择列表

Dim skSegment As Object
Set skSegment = part.SketchManager.CreateCircle(0#, 0#, 0#, 0.017934, 0.015738, 0#)      '画圆
part.ClearSelection2 True
boolstatus = part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 2.23259369346084E-02, 6.95398035668131E-03, 0, False, 0, Nothing, 0)

Dim myDisplayDim As Object
Set myDisplayDim = part.AddDimension2(3.73318945463944E-02, 5.85598345825794E-03, 0)    '添加尺寸
part.ClearSelection2 True
boolstatus = part.Extension.SelectByID2("D1@草图1@零件3.SLDPRT", "DIMENSION", 3.73318945463944E-02, 5.85598345825794E-03, 0, False, 0, Nothing, 0)   '选择草图

Dim myDimension As Object
Set myDimension = part.Parameter("D1@草图1")
myDimension.SystemValue = a / 1000               '设置圆的尺寸直径，暂时取值为0.06-0.07，因为取值过大会引起图纸比例自动变化，影响计算，也不太符合实际需求。
part.ShowNamedView2 "*上下二等角轴测", 8
part.ClearSelection2 True
boolstatus = part.Extension.SelectByID2("D1@草图1@零件3.SLDPRT", "DIMENSION", 3.73318945463944E-02, 5.85598345825794E-03, 0, False, 0, Nothing, 0)

Dim myFeature As Object        'feature此处的特征是做拉伸
Set myFeature = part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, b / 1000, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, True, True, True, 0, 0, False) '第6个参数控制拉升的厚度，暂时取值0.01-0.02，太大或太小会影响会影响图纸比例，不利于计算
part.SelectionManager.EnableContourSelection = False
part.ShowNamedView2 "*等轴测", 7
part.ShowNamedView2 "*等轴测", 7
longstatus = part.SaveAs3(App.Path & "\" & "圆盘.SLDPRT", 0, 2)     '保存零件图，保存在和程序一起的文件夹中



Dim longwarnings As Long
Set moddoc = swApp.NewDocument(App.Path & "\" & "标准图纸\gb_a0.drwdot", 12, 0.841, 1.189)  '第1个参数表示调用模板的位置，后面3个参数表示12号图纸长度841mm宽度1189mm，新建工程图，用外部文件，使用sw自带模板图报错，原因未知
Set moddoc = swApp.ActiveDoc

Dim myView As Object
'Set myView = moddoc.CreateDrawViewFromModelView3(App.Path & "\" & "圆盘.SLDPRT", "*前视", 0.265882219679634, 0.61352585812357, 0) '第1个参数表示调用三维零件图的位置，第2个参数表示零件图的前视图作为新建工程图的主视图，后面3个参数表示插入的主视图在工程图的位置，新建工程图的坐标原点为图纸的左下角，主视图插入的基点为零件视图的中心，要依据图纸和视图的大小确定插入点的坐标
Set myView = moddoc.CreateDrawViewFromModelView3(App.Path & "\" & "圆盘.SLDPRT", "*前视", 0.26, 0.613, 0)
boolstatus = moddoc.Extension.SelectByID2("工程图视图1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = moddoc.ActivateView("工程图视图1")

'Set myView = moddoc.CreateUnfoldedViewAt3(0.265882219679634, 0.202936956521739, 0, False)  '插入俯视图，前3个参数表示坐标值，最后一个参数false表示与父视图（此处是主视图）对齐
Set myView = moddoc.CreateUnfoldedViewAt3(0.26, 0.2, 0, False)
moddoc.ClearSelection2 True
boolstatus = moddoc.Extension.SelectByID2("工程图视图1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
moddoc.ClearSelection2 True
'boolstatus = moddoc.Extension.SelectByID2("", "EDGE", 0.411469519450801, 0.656056979407828, 400.005, False, 0, Nothing, 0)
boolstatus = moddoc.Extension.SelectByID2("", "EDGE", 0.26 + a / 1000 / 2 * 5, 0.613, 400.005, False, 0, Nothing, 0) '选择点的坐标位置，横坐标圆心坐标，纵坐标圆心坐标加半径（乘以5是图的比例），z轴是宏得来的

Dim myDisplayDim1 As Object
'Set myDisplayDim1 = moddoc.AddDimension2(0.265882219679634, 0.61352585812357, 0)     '直径标注位置的坐标（精确到mm，不然可能不显示标注）标注（添加）工程图尺寸
Set myDisplayDim1 = moddoc.AddDimension2(0.26 + 0.01 * 5, 0.613 + 0.01 * 5, 0)   '标注位置，大概圆内任选一点就可
moddoc.ClearSelection2 True
boolstatus = moddoc.ActivateSheet("图纸1")
boolstatus = moddoc.ActivateView("工程图视图2")
'boolstatus = moddoc.Extension.SelectByID2("", "SILHOUETTE", 0.41474114416476, 0.212751830666409, 400, False, 0, Nothing, 0)
'boolstatus = moddoc.Extension.SelectByID2("", "SILHOUETTE", 0.41474114416476 + 0.005 * 5, 0.212751830666409, 400, False, 0, Nothing, 0) '标注尺寸时，选择右边边线
boolstatus = moddoc.Extension.SelectByID2("", "SILHOUETTE", 0.26 + a / 1000 / 2 * 5, 0.2 + b / 1000 / 2 * 5, 400, False, 0, Nothing, 0)
'boolstatus = moddoc.Extension.SelectByID2("", "EDGE", 0.261, 0.21, 400, False, 0, Nothing, 0)
'boolstatus = moddoc.Extension.SelectByID2("", "EDGE", 0.261, 0.21 - 0.01 * 5, 400, True, 0, Nothing, 0)
'Set myDisplayDim1 = moddoc.AddDimension2(0.44254995423341, 0.206208581235698, 0)     '标注（添加）工程图尺寸
Set myDisplayDim1 = moddoc.AddDimension2(0.26 + (a / 1000 / 2 + 0.002) * 5, 0.2, 0) '拉伸长度标注位置的坐标，乘以5是因为图比例5:1
'Set myDisplayDim1 = moddoc.AddDimension2(0.26, 0.2, 0)
boolstatus = moddoc.Extension.SelectByID2("RD1@工程图视图1", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = moddoc.ActivateSheet("图纸1")
boolstatus = moddoc.ActivateView("工程图视图2")
boolstatus = moddoc.ActivateSheet("图纸1")
moddoc.ClearSelection2 True
moddoc.ViewZoomtofit2
moddoc.SheetPrevious
moddoc.ViewZoomTo2 0, 0, 0, 0.1, 0.1, 0.1
moddoc.ViewZoomTo2 0, 0, 0, 0.1, 0.1, 0.1
moddoc.ViewZoomTo2 0, 0, 0, 0.1, 0.1, 0.1
moddoc.ViewZoomtofit2
longstatus = moddoc.SaveAs3(App.Path & "\" & "圆盘.DWG", 0, 0)   '保存工程图
Else
    MsgBox ("弹窗")
End If
End Sub

Private Sub Command2_Click()
    Unload Form1
End Sub

Private Sub Form_Load()
    Form1.Width = 10000
    Form1.Height = 5000
    Label1.Caption = "直径"
    Label2.Caption = "厚度"
    Frame1.Caption = "圆盘参数"
    Command1.Caption = "开始"
    Command2.Caption = "结束"
    Text1.Text = ""
    Text2.Text = ""
End Sub
