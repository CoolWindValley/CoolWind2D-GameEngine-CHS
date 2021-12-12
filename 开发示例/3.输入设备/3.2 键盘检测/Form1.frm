VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "键盘"
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
         
         If CWKeyboard.A.PUP Then          '检测A键抬起状态
         CWPrintFont FontDemo, "A键处于抬起状态", 0, 0, 800, 60, CWWhite, CWF_LeftAl
         End If
         
         If CWKeyboard.Enter.PUPMoment Then          '检测回车键抬起瞬间
         CWPrintFont FontDemo, "回车键抬起的瞬间", 0, 100, 800, 60, CWRed, CWF_LeftAl
         End If
         
         If CWKeyboard.Space.PDownMoment Then    '检测空格键按下瞬间
         CWPrintFont FontDemo, "空格键按下的瞬间", 0, 200, 800, 60, CWGreen, CWF_LeftAl
         End If
         
         If CWKeyboard.UP.PDown Then          '检测方向键↑按下状态
         CWPrintFont FontDemo, "↑方向键处于按下状态", 0, 300, 800, 60, CWBlue, CWF_LeftAl
         End If
         
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
