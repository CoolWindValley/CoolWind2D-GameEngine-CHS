VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "环绕音效"
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

    Dim FontDemo As CWFont      '定义显示用的字体

    CWLoadFont FontDemo, "Microsoft YaHei", 32, CWF_Normal, False      '载入普通宋体32号字体

    Dim MusicDemo1 As CWMusic, MusicDemo2 As CWMusic    '定义CoolWind引擎音频变量
    
    '载入音乐(音频变量，音频文件路径，特效参数)
    CWLoadMusic MusicDemo1, App.Path & "\Music\" & "test.mp3", CWME_3DFX
    CWLoadMusic MusicDemo2, App.Path & "\Music\" & "effect.mp3", CWME_3DFX
    
    CWPlayMusic MusicDemo1, CWM_Repeat
        
    '本引擎尚未完善处:音频载入时间长,且载入时窗口会卡住.正式版将至少使编译后的文件载入更流畅.
    
    Do While CWGameRun = True         '进入游戏循环
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '检测是否可以渲染（设备正常并且窗体未最小化时渲染）
    CWBeginScene    '准备好绘制场景
      CWPaintPicBegin
        CWPrintFont FontDemo, "背景音乐自动在玩家周围环绕", 0, 0, 800, 60, CWFuchsia, CWF_CenterAl Or DT_VCENTER Or DT_NOCLIP
        CWPrintFont FontDemo, "点击鼠标左键可在点击位置播放音效" & vbNewLine _
            & "按住鼠标左键不放并拖动可改变音效的空间坐标", 0, 500, 800, 100, CWViolet, CWF_CenterAl Or DT_VCENTER Or DT_NOCLIP
        
        ' 画出玩家人头朝向的模拟
        CWDrawHTriangle 400, 280, 380, 312, 420, 312, CWColorRGBA(0, 255, 255, 120)
        
        Const R! = 240!, PxpMi = 100
        Dim t1!: t1 = t1 + 0.02!
        Dim X1!: X1 = R * Cos(t1)
        Dim Y1!: Y1 = R * Sin(t1)
        
        ' 画出环绕背景音乐的路径
        CWDrawHCircle 400, 300, R, CWYellow
        CWDrawLine2 400, 300, 400 + X1, 300 + Y1, CWColorARGB(100, 255, 255, 0), True, CWLP_Dot, 2
        ' 更新环绕背景音乐的空间坐标（注意前后是Z轴，不是Y轴）
        CWMusicSet3DPosition MusicDemo1, X1 / PxpMi, , Y1 / PxpMi
        
        Dim t2!, a2!, X2!, Y2!
        If CWMouse.LeftKey.PDownMoment Then
            ' 点击鼠标左键时开始播放音效并初始化数据
            CWPlayMusic MusicDemo2, CWM_Restart
            t2 = 0: a2 = 200
        End If
        If CWMouse.LeftKey.PDown Then
            ' 按住鼠标左键时移动音效的音源位置
            X2 = CWMouse.X: Y2 = CWMouse.Y
            CWMusicSet3DPosition MusicDemo2, (X2 - 400) / PxpMi, , (Y2 - 300) / PxpMi
        End If
        If a2 > 0 Then
            ' 当画出近似的声波扩散图
            t2 = t2 + 3.14159265358979
            CWDrawSCircle X2, Y2, t2, CWColorARGB(CByte(a2), 20, 200, 255)
            a2 = a2 - 2.44645364561314
        End If
        
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
