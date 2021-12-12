VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
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

Dim PicDemo As CWPic    '定义CoolWind引擎图片变量

    CWLoadPic PicDemo, App.Path & "\Pic\bgimg.png", CWColorNone     '载入地面图片

    Do While CWGameRun = True         '进入游戏循环
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '检测是否可以渲染（设备正常并且窗体未最小化时渲染）
    CWBeginScene    '准备好绘制场景
         
      CWPaintPicBegin     '贴图开始
    
        ' 设置采样模式为循环（可以方便实现平铺和循环滚屏）
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_WRAP
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_WRAP
        ' 画背景前先禁用Alpha混合（提高效率）
        CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, False
        CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, False

        CWPaintPicEx PicDemo, 0, 0, 0, 0, 800, 600, CWGrey '贴上地面贴图，混合灰色模仿黑夜

      CWPaintPicEnd       '贴图结束
 
      ' 贴图结束后不用提交精灵，直接开启Alpha混合
      CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, True
      CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, True

      LightEFOpen       '打开光照效果
         
        '光照效果开启后所有的贴图和绘图都会被作为光源处理，透明度（A）决定强度，颜色（RGB）决定色调
        CWDrawSCircle 400, 300, 192, CWWhite                    '画圆作为光源
        CWDrawSCircleEX 400, 300, 192, 2, 1, CWYellow, CWHA_Red '重叠的光源图形越多，光照越强
    
            '注意：文字绘制（后面的示例即将提到）不受光照影响

      LightEFClose      '关闭光照效果

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
