VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "贴图"
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


Dim PicDemo1 As CWPic, PicDemo2 As CWPic    '定义CoolWind引擎图片变量

    CWLoadPic PicDemo1, App.Path & "\Pic\M1.png", CWColorNone
    CWLoadPic PicDemo2, App.Path & "\Pic\M2.png", CWColorNone
    '载入图片(图片变量,图片路径,屏蔽色)
    '关于屏蔽色:顾名思义,就是不显示图片上的指定的某种颜色,通常设为无色即可(即不屏蔽任何颜色)
    
    Do While CWGameRun = True         '进入游戏循环
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '检测是否可以渲染（设备正常并且窗体未最小化时渲染）
    CWBeginScene CWBlue '准备好绘制场景
         
    CWPaintPicBegin     '贴图开始 必须与贴图结束成对出现，用于通知引擎准备贴图
    
    '贴图开始到贴图结束之间不要使用CWDraw绘图函数，否则将产生一些奇怪的效果。

         CWPaintPic PicDemo1, 100, 100
         CWPaintPic PicDemo2, 400, 100
         '贴图(图片,起点横坐标,起点纵坐标)
         
         CWPaintPicEX PicDemo1, 100, 400, 0, 0, 120, 137, CWYellow    '贴图扩展函数,能裁剪贴图并改变图像色调
         '高级贴图(图片,起点横坐标,起点纵坐标,裁剪起点横坐标,裁剪起点纵坐标,裁剪宽度,裁剪高度,图像色调)
            '注意:裁剪起点的坐标是以图片本身作为参照系的.   若不想改变图像色调将其设置为白色(CWWhite)即可

    CWPaintPicEXEX PicDemo2, 350, 450, 0, 0, 208, 145, 1.5, 1.5, 350, 450, CRad(-45), CWRed   '贴图扩展函数,能裁剪、旋转贴图并改变图像色调、大小
    '超高级贴图(图片,起点横坐标,起点纵坐标,裁剪起点横坐标,裁剪起点纵坐标,裁剪宽度,裁剪高度,横向缩放倍数,纵向缩放倍数, _
                旋转轴横坐标,旋转轴纵坐标,旋转角度,图像色调)
        '注意:旋转轴坐标是以游戏窗口作为参照系的.   缩放倍数大于一为放大,小于一为缩小.
                '旋转角度采用弧度制,正数为顺时针旋转,负数为逆时针旋转.  可用CRad函数将角度转化为弧度

    CWPaintPicEnd       '贴图结束 必须与贴图开始成对出现，用于通知引擎贴图完毕
   
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
