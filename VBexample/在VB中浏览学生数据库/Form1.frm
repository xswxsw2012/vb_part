VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2055
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command3 
      Caption         =   "选课信息"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "课程信息"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "基本信息"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\desk\xs.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Data1.RecordSource = "基本信息"    '设置Data控件可以访问的数据为“基本信息”表
    Data1.Refresh                     '刷新
    Data1.Caption = "基本信息"
    Form1.Caption = "学生基本信息浏览"
End Sub

Private Sub Command2_Click()
    Data1.RecordSource = "课程表"    '设置Data控件可以访问的数据为“课程表”表
    Data1.Refresh                     '刷新
    Data1.Caption = "课程信息"
    Form1.Caption = "课程信息浏览"
End Sub


Private Sub Command3_Click()
    Data1.RecordSource = "选课信息"    '设置Data控件可以访问的数据为“选课信息”表
    Data1.Refresh                     '刷新
    Data1.Caption = "选课信息"
    Form1.Caption = "学生选课信息浏览"
End Sub
