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
    
    '画线(起点横坐标,起点纵坐标,终点横坐标,终点纵坐标,颜色)
    CWDrawLine 80, 50, 150, 50, CWRed
    
    '画线2(起点横坐标,起点纵坐标,终点横坐标,终点纵坐标,颜色[,抗锯齿,样式,线宽,缩放])
    CWDrawLine2 80, 75, 150, 75, CWRed                          ' 与原版对比
    CWDrawLine2 80, 100, 150, 100, CWRed, , , 2!                ' 粗线
    CWDrawLine2 80, 125, 150, 125, CWRed, True                  ' 抗锯齿线
    CWDrawLine2 80, 150, 150, 150, CWRed, , CWLP_Dash           ' 虚线 ___ '
    CWDrawLine2 80, 175, 150, 175, CWRed, , CWLP_Dot            ' 虚线 _ '
    CWDrawLine2 80, 200, 150, 200, CWRed, , CWLP_DashDot        ' 虚线 ___ _ '
    CWDrawLine2 80, 225, 150, 225, CWRed, , CWLP_DashDotDot     ' 虚线 _ ___ _ '
    CWDrawLine2 80, 250, 150, 250, CWRed, , CWLP_Minus          ' 虚线 __ '
    CWDrawLine2 80, 275, 150, 275, CWRed, , CWLP_DashMinus      ' 虚线 ___ __ '
    CWDrawLine2 80, 300, 150, 300, CWRed, , CWLP_MinusDot       ' 虚线 __ _ '
    CWDrawLine2 80, 325, 150, 325, CWRed, , CWLP_MinusDotDot    ' 虚线 _ __ _ '
    CWDrawLine2 80, 350, 150, 350, CWRed, , CWLP_Point          ' 虚线. . . ’
    CWDrawLine2 80, 375, 150, 375, CWRed, , CWLP_InvPoint       ' 虚线 . . .’
    CWDrawLine2 80, 400, 150, 400, CWRed, , CWLP_DotPointPoint  ' 虚线 . _ .’
    
    '抗锯齿对比
    CWDrawLine 200, 50, 350, 100, CWGreen               ' 原版画线
    CWDrawLine2 200, 75, 350, 125, CWGreen, False       ' 新版画线（有锯齿）
    CWDrawLine2 200, 100, 350, 150, CWGreen, True       ' 新版画线（抗锯齿）
    CWDrawLine2 200, 125, 350, 175, CWGreen, False, , 2!    ' 新版画线（加粗、有锯齿）
    CWDrawLine2 200, 150, 350, 200, CWGreen, True, , 2!     ' 新版画线（加粗、抗锯齿）
    CWDrawLine2 200, 175, 350, 225, CWGreen, False, , 5!    ' 新版画线（更粗、有锯齿）
    CWDrawLine2 200, 200, 350, 250, CWGreen, True, , 5!     ' 新版画线（更粗、抗锯齿）
    CWDrawLine2 200, 225, 350, 275, CWGreen, False, CWLP_DashDot, 2!        ' 新版画线（有锯齿、虚线、加粗、拉长）
    CWDrawLine2 200, 250, 350, 300, CWGreen, True, CWLP_DashDot, 2!         ' 新版画线（抗锯齿、虚线、加粗、拉长）
    CWDrawLine2 200, 275, 350, 325, CWGreen, False, CWLP_DashDot, 2!, 2!    ' 新版画线（有锯齿、虚线、加粗、拉长）
    CWDrawLine2 200, 300, 350, 350, CWGreen, True, CWLP_DashDot, 2!, 2!     ' 新版画线（抗锯齿、虚线、加粗、拉长）
    
    '画渐变色线(起点横坐标,起点纵坐标,终点横坐标,终点纵坐标,起点颜色权重,重点颜色权重,起点颜色,终点颜色)
    CWDrawLineEx 400, 50, 550, 50, 1, 2.5, CWRed, CWBlue
    
    '画多段线(坐标数组,颜色[,抗锯齿,样式,线宽,缩放])
    Dim pts(0 To 5) As D3DXVECTOR2
    MakeStar pts, 100, 500, 200
    CWDrawLine2Ex pts, CWBlue
    MakeStar pts, 100, 500, 400
    CWDrawLine2Ex pts, CWBlue, True, CWLP_Dot, 2!, 2!
    
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

Private Sub MakeStar(pts() As D3DXVECTOR2, ByVal R!, ByVal X!, ByVal Y!)
    Dim I&, L&, Ang!: L = UBound(pts) - LBound(pts)
    For I = LBound(pts) To UBound(pts)
      With pts(I)
        Ang = I * 4! * Pi / L
        .X = X + R * Sin(Ang)
        .Y = Y - R * Cos(Ang)
      End With
    Next
    
End Sub
