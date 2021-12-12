VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "绘图"
   ClientHeight    =   2985
   ClientLeft      =   3090
   ClientTop       =   4350
   ClientWidth     =   5055
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     '变量使用前必须声明
'新手注意：使用前请确保 DX9VB.TLB组件已被引用，VBDX9BAS.bas模块已被添加

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
AltBUGRepair KeyCode    '修复按下ALT导致画面停止刷新的临时解决函数
End Sub

Private Sub Form_Load()

'新手注意：游戏编程中，
'通常将窗体的 BorderStyle 设置为“Fixed single”即不允许改变窗体大小
'通常将窗体的 MinButton 设置为“True”即允许最小化
'通常将窗体的 MaxButton 设置为“False”即禁止最大化

  '初始化引擎并设置引擎初始化窗体和引擎分辨率，但最好是电脑常用的分辨率比如 640,480 、 800,600 、 1024,768 、 1366,768
CWVBDX9Initialization Me, 800, 600, CW_Windowed
'初始化引擎（目标窗体，横向分辨率，纵向分辨率，窗口模式/全屏模式）
    
    Do While CWGameRun = True         '进入游戏循环
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '检测是否可以渲染（设备正常并且窗体未最小化时渲染）
    CWBeginScene    '准备好绘制场景
    
    '关于颜色:可以使用BAS模块中已经设置好的常用颜色,也可用CWColorARGB函数自行转换,A分量为不透明度,RGB分量对应的颜色请参考RGB颜色表
    '关于颜色权重:颜色权重越高,该点的颜色在渐变过程中所占的区域就越大
    
    CWDrawPoint 50, 50, CWRed
    '画点(横坐标,纵坐标,颜色)
    
    CWDrawline 80, 50, 150, 50, CWRed
    '画线(起点横坐标,起点纵坐标,终点横坐标,终点纵坐标,颜色)
         
    CWDrawlineEX 200, 50, 350, 50, 1, 2.5, CWRed, CWBlue
    '画渐变色线(起点横坐标,起点纵坐标,终点横坐标,终点纵坐标,起点颜色权重,重点颜色权重,起点颜色,终点颜色)
    
    CWDrawHRect 400, 10, 150, 80, CWYellow
    '画空心矩形(起点横坐标,起点纵坐标,宽度,高度,颜色)
    
    CWDrawSRect 600, 10, 150, 80, CWGreen
    '画实心矩形(起点横坐标,起点纵坐标,宽度,高度,颜色)
    
    CWDrawSRectXGC 50, 100, 150, 150, 1.5, 1, CWPurple, CWCyan
    '画横向渐变色实心矩形(起点横坐标,起点纵坐标,宽度,高度,左边颜色权重,右边颜色权重,左边颜色,右边颜色)
    
    CWDrawSRectYGC 250, 100, 150, 150, 1, 1.5, CWBlue, CWYellow
    '画纵向渐变色实心矩形(起点横坐标,起点纵坐标,宽度,高度,上边颜色权重,下边颜色权重,上边颜色,下边颜色)
    
    CWDrawSRectXCGC 450, 100, 150, 150, 2, 1, CWRed, CWYellow
    '画横向中心渐变色实心矩形(起点横坐标,起点纵坐标,宽度,高度,中心颜色权重,两边颜色权重,中心颜色, 两边颜色)
    
    CWDrawSRectYCGC 650, 100, 150, 150, 2, 1, CWPurple, CWBlue
    '画纵向中心渐变色实心矩形(起点横坐标,起点纵坐标,宽度,高度,中心颜色权重,两边颜色权重,中心颜色, 两边颜色)

    CWDrawSRectCGC 30, 300, 150, 150, 1.5, 1, CWRed, CWColorARGB(255, 125, 0, 125)
    '画中心渐变色实心矩形(起点横坐标,起点纵坐标,宽度,高度,中心颜色权重,周边颜色权重,中心颜色, 周边颜色)
    
    CWDrawHCircle 300, 400, 100, CWBlue
    '画空心圆(圆心横坐标,圆心纵坐标,半径,颜色)
    
    CWDrawSCircle 500, 500, 100, CWPurple
    '画实心圆(圆心横坐标,圆心纵坐标,半径,颜色)

    CWDrawSCircleEX 620, 400, 150, 3, 1, CWHA_Yellow, CWRed
    '画中心渐变色实心圆(圆心横坐标,圆心纵坐标,半径,中心颜色权重,周边颜色权重,中心颜色, 周边颜色)
         
    CWPresentScene   '呈现绘制的场景

'*******************************以下为固定写法，不要轻易改动***********************************
    Else                 '当不满足渲染条件时
        CWResetDevice       '修复设备
    End If

    Loop

        CWVBDX9Destory     '销毁CoolWind引擎
    End '退出
'*******************************以上为固定写法，不要轻易改动***********************************


End Sub

Private Sub Form_Unload(Cancel As Integer)
CWGameRun = False       '窗体被关闭时关闭引擎
MediaBUGRepair          '修复非正常退出时，正在播放的音乐不停止的临时解决函数（IDE下无效）
End Sub
