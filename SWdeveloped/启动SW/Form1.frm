VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "启动SW2014并画圆盘"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim swapp As Object
Dim part As Object
Dim boolstatus As Boolean
'定义接口
Set swapp = CreateObject("SldWorks.Application")
Set part = swapp.NewPart()
swapp.Visible = True
'获取SolidWorks软件接口并新建一个零件文件，对于电脑没有安装sw的情况暂时挂起
part.InsertSketch2 True
boolstatus = part.Extension.SelectByID("前视", "PLANE", 0, 0, 0, False, 0, Nothing)  '截止到此处可启动sw2014并创建一个前视窗口
part.InsertSketch2 True                                 '下列程序可以自动画圆盘
part.CreateCircle 0, 0, 0, 0, 50, 0
part.ShowNamedView2 "上下二等角轴测", 8
part.FeatureManager.FeatureExtrusion True, False, False, 0, 0, 10000 / 1000, 0.01, False, False, False, False, 0, 0, False, False, False, False, 1, 1, 1
End Sub

Private Sub Command2_Click()
    End
End Sub
