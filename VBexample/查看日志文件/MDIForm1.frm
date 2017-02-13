VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "信息系统主界面"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Menu browerlog 
      Caption         =   "查看日志"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()   '主窗体
    nstr = "启动系统"
    nstr = Format(Now, "yyyy-mm-dd hh:mm:ss") & nstr '变量的值为系统时间与操作
    F = App.Path & "\" & "xsxx.log" '读取文件的路径与文件名
    H = FreeFile       '用函数求出目前最小的未使用的文件号
    Open F For Append As #H   '以追加方式打开文件
    Print #H, nstr       '将变量的值写入文件
    Close #H       '关闭文件
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    nstr = "退出系统"
    nstr = Format(Now, "yyyy-mm-dd hh:mm:ss") & nstr '变量的值为系统时间与操作
    F = App.Path & "\" & "xsxx.log"  '读取文件的路径与文件名
    H = FreeFile   '用函数求出目前最小的未使用的文件号
    Open F For Append As #H        '以追加方式打开文件
    Print #H, nstr        '将变量的值写入文件
    Close #H      '关闭文件
End Sub
Private Sub browerlog_Click()
    Form1.Show
End Sub
