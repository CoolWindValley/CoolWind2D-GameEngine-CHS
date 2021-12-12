VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文字"
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

Dim FontDemo1 As CWFont, FontDemo2 As CWFont, FontDemo3 As CWFont    '定义CoolWind引擎字体变量
Const HelloWorld$ = "Hello World!"

    CWLoadFont FontDemo1, "SimSun", 32, CWF_Bold, False
    CWLoadFont FontDemo2, "@SimHei", 32, CWF_Normal, False
    CWLoadFont FontDemo3, "Microsoft YaHei", 64, CWF_Normal, False
    '载入字体(字体变量,字体类型,字体大小,是否为粗体,是否为斜体)
        '注意:载入的字体类型必须是正在运行该游戏的系统上存在的字体.否则将默认按"宋体"载入
        '小技巧:在大多数字体名前加上@可以让字体横过来~!
    
    Do While CWGameRun = True         '进入游戏循环
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '检测是否可以渲染（设备正常并且窗体未最小化时渲染）
    CWBeginScene    '准备好绘制场景
        CWPaintPicBegin
        
        CWPrintFont FontDemo1, CWFPS, 0, 0, 800, 60, CWWhite, CWF_LeftAl    '显示当前FPS
        CWPrintFont FontDemo1, "队长～　　　是我～！", 0, 100, 800, 60, CWCyan, CWF_LeftAl
        CWPrintFont FontDemo2, "      别开枪        ", 0, 100, 800, 60, CWRed, CWF_LeftAl
        '输出字体(字体,目标文字,显示起点横坐标,显示起点纵坐标,显示区宽度,显示区高度,字体颜色,对齐方式)
               '注意:文字总长度超过显示区宽度将引起换行.   文字总长度超过显示区高度将引起裁剪,即超过部分不显示.
        Dim cx&, cy&    '计算文本显示大小，并绘制文本边框。
        CWCalcrFont FontDemo3, HelloWorld, cx, cy, True
        CWPrintFont FontDemo3, HelloWorld, 400 - cx * 0.5!, 300, cx, cy, CWYellow, DT_SINGLELINE Or DT_NOCLIP
        CWDrawHRect 400 - cx * 0.5!, 300, cx, cy, CWBlue
        
        CWPaintPicEnd
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
