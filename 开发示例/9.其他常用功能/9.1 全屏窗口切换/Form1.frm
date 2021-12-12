VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "全屏、窗口切换"
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

Dim FontDemo As CWFont      '定义CoolWind引擎字体变量

    CWLoadFont FontDemo, "SimSun", 32, CWF_Normal, False  '载入要显示的字体

    Do While CWGameRun = True         '进入游戏循环
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '检测是否可以渲染（设备正常并且窗体未最小化时渲染）
    CWBeginScene    '准备好绘制场景
         
         CWPrintFont FontDemo, "按下ALT+ENTER键切换全屏和窗口模式", 0, 0, 800, 60, CWWhite, CWF_LeftAl

         If CWKeyboard.Enter.PDown And CWKeyboard.Enter.PDownMoment Then          '检测回车键抬起瞬间
         CWWinFullScrSwitch         '执行全屏/窗口模式切换
         End If
         
'注意：窗口模式下鼠标键盘检测灵敏度较高（延迟<10ms）而引擎绘图性能略低（系统特别繁忙时易出现画面抖动）
'（若系统CPU使用率并不高而出现画面抖动，通常是迅雷、360等软件的悬浮窗干扰了DX9的画面刷新，暂时将其关闭即可。绝大多数DX9游戏都有此类问题。作者也在尝试解决。）
'    全屏模式下鼠标键盘检测灵敏度较低（延迟<30ms）而引擎绘图性能较高。（即使在掉帧的时候也能极好的抑制画面抖动）
'    请根据自己游戏的特点选择游戏启动时默认的显示模式

    '小科普：人类视觉的时间延迟为20ms左右，人类手按键反应延迟为100ms以上（即600APM为人类极限）
    
         
         If CWKeyboard.ESC.PDownMoment Then   '按下ESC键的瞬间退出程序
         CWGameRun = False           '关闭引擎开关
         End If
    
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
