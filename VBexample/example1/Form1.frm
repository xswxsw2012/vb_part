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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()  '窗体激活事件
    Form1.Print                'Print 方法无参数时输出一空行
    Form1.Print Tab(5); "快乐学习 Visual Basic！"  'Tab（n）跳过n个字符的位置再输出字符串
End Sub


Private Sub Form_Click()  '窗体单击事件
Form1.Cls '清除窗体所有文字
Form1.BackColor = RGB(0, 255, 0) '表示红色和蓝色的分值为0，结果为黄色

Form1.ForeColor = RGB(255, 0, 0)
Form1.FontName = "楷体"
Form1.Print Chr(13); Tab(5); "快乐学习 Visual Basic！" 'Chr(13)表示先换行再输出
End Sub


Private Sub Form_DblClick()      '卸载窗体事件Form1
    Unload Form1
End Sub


Private Sub Form_Load()       '窗体载入事件
    Form1.ForeColor = RGB(0, 255, 0)
End Sub
