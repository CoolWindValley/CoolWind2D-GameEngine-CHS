VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "综合测试"
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

    Dim bg As CWPic, yun As CWPic, spider As CWPic, light As CWPic, fnt As CWFont, bgm As CWMusic
    Const SpiderX& = 400 - 104  ' 图像的X位置
    Const SpiderY& = 410        ' 图像的Y位置
    
    ' 加载资源
    CWLoadFont fnt, "Microsoft YaHei", 32, CWF_Bold, False
    CWLoadPic bg, App.Path & "\Pic\bgimg.png", CWColorNone
    CWLoadPic yun, App.Path & "\Pic\yun.png", CWColorNone
    CWLoadPic spider, App.Path & "\Pic\spider.png", CWColorNone
    CWLoadPic light, App.Path & "\Pic\light.png", CWColorNone
    CWLoadMusic bgm, App.Path & "\Snd\test.mp3"
    CWPlayMusic bgm, CWM_Repeat

    Do While CWGameRun         '进入游戏循环
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '检测是否可以渲染（设备正常并且窗体未最小化时渲染）
    CWBeginScene    '准备好绘制场景
    CWPaintPicBegin
    
        ' 设置采样模式为循环（可以方便实现平铺和循环滚屏）
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_WRAP
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_WRAP
        ' 画背景前先禁用Alpha混合（提高效率）
        CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, False
        CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, False

        Dim l!, R!, i&, j&
        l = l + 1!: If l > 0! Then l = l - 893!
        R = R + 1!: If R > 180! Then R = R - 360!
         
        CWPrintFontTop fnt, "Hello World!", 0, 0, 480, 500, CWBlue, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP

#Const Title = False    ' 平铺有两种方式实现
' True 为纹理采样方式实现（效率高，但只能实现单张图片）
' False 为循环贴图方式实现（效率低一点，但可以实现多张图随意切换）
#If Title Then
        CWPaintPicEx bg, 0, 0, 0, 0, 800, 600
#Else
        For i = 0 To 24
            For j = 0 To 18
                CWPaintPic bg, i * 32, j * 32
        Next j, i
#End If
        CWPaintPicFlush ' 背景画完后要先提交，再开启Alpha混合
        CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, True
        CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, True

        CWPaintPicEx yun, 0, 0, l, 0, 800, 325
        CWPaintPicFlush ' 改变D3D渲染模式之前都要先提交精灵渲染（不然会导致之前的画的内容也受模式修改影响）
        
        ' 设置采样模式为边框（边框颜色默认为透明色，适合非背景类贴图）
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_BORDER
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_BORDER

        CWPaintPicExEx spider, 400 - 104, 300 - 72.5, 0, 0, 208, 145, , , 400, 300, R * Pi / 180!
        CWPaintPic spider, SpiderX, SpiderY     ' 蜘蛛原图（这个区域用来做命中测试的）
    
        Dim cwc As CWColor      ' 获取鼠标光标位置的颜色值
        cwc = CWSplitColor(CWPicGetPixel(spider, CWMouse.X - SpiderX, CWMouse.Y - SpiderY))
        CWPrintFontTop fnt, "XY(" & CWMouse.X & ", " & CWMouse.Y & ")" & vbNewLine & "ARGB(" & cwc.Alpha & ", " & cwc.Red & ", " & cwc.Green & ", " & cwc.Blue & ")", 0, 560, 800, 0, CWGreen, DT_CENTER Or DT_VCENTER Or DT_NOCLIP

        LightEFOpen True
            ' 如果鼠标命中，就以一半亮度的蜘蛛图像作为光源照在上面（实现变亮的效果）
            If IsHitWnd And cwc.Alpha >= &H80 Then CWPaintPic spider, SpiderX, SpiderY, CWGrey
            ' 先旋转后缩放的圆形光斑（有伪3D的效果）
            CWPaintPicExEx1 light, 200, 460, 0, 0, 256, 256, 2!, 0.5!, 128, 128, -R * Pi / 180!
            ' 先缩放后旋转的圆形光斑（传统2D的效果）
            CWPaintPicExEx2 light, 600, 440, 0, 0, 256, 256, 2!, 0.5!, 128, 128, R * Pi / 180!
        LightEFClose
    
        CWDrawSRect 300, 200, 200, 200, &H40FFFF00
        
        CWPrintFontTop fnt, CWFPS, 0, 0, 0, 0, CWRed, DT_SINGLELINE Or DT_NOCLIP
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
