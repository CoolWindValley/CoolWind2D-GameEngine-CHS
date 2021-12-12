VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "音频播放"
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

    CWLoadFont FontDemo, "SimSun", 32, CWF_Normal, False      '载入普通宋体32号字体

    Dim MusicDemo1 As CWMusic, MusicDemo2 As CWMusic    '定义CoolWind引擎音频变量
    
            CWLoadMusic MusicDemo1, App.Path & "\Music\" & "游戏王解开封印.mp3"
            CWLoadMusic MusicDemo2, App.Path & "\Music\" & "游戏王激烈的决斗.mp3"
            '载入音乐(音频变量,音频文件路径)
            
            CWSetMusicVol MusicDemo1, 100
            CWSetMusicVol MusicDemo2, 200
             '设置音量(音频变量,声音大小)
                '若不设置声音大小默认是1000(最大值)
                
            '本引擎尚未完善处:音频载入时间长,且载入时窗口会卡住.正式版将至少使编译后的文件载入更流畅.
    
    Do While CWGameRun = True         '进入游戏循环
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '检测是否可以渲染（设备正常并且窗体未最小化时渲染）
    CWBeginScene    '准备好绘制场景
         
        CWPrintFont FontDemo, "P键播放音乐1，S键停止音乐1；M键播放音乐2，E键停止音乐2；B键暂停1和2", 0, 100, 800, 100, CWYellow, CWF_CenterAl
        If CWKeyboard.P.PDownMoment Then
        CWPlayMusic MusicDemo1, CWM_Repeat     'P键按下的瞬间循环播放音乐1
        '播放音乐(音频,播放模式[循环/一次])
        End If
        
        If CWKeyboard.S.PDownMoment Then
        CWStopMusic MusicDemo1              'S键按下的瞬间停止音乐1
        '停止音乐(音频)
        End If
        
        If CWKeyboard.M.PDownMoment Then
        CWPlayMusic MusicDemo2, CWM_Once    'M键按下的瞬间播放音乐2
        End If
        
        If CWKeyboard.E.PDownMoment Then
        CWStopMusic MusicDemo2             'E键按下的瞬间停止音乐2
        End If
        
        If CWKeyboard.B.PDownMoment Then
        CWPauseMusic MusicDemo1             'B键按下的瞬间暂停音乐1和2
        CWPauseMusic MusicDemo2
        '暂停音乐(音频)
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
