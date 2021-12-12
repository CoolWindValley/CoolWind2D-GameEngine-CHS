VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "旋转文字"
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

Const TestText$ = "Cool Wind 2D" & vbNewLine & "文字旋转测试"

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

Dim FontDemo As CWFont    '定义CoolWind引擎字体变量
Dim FontRoll As Single

    CWLoadFont FontDemo, "Microsoft YaHei", 32, CWF_Normal, False
    '载入字体(字体变量,字体类型,字体大小,是否为粗体,是否为斜体)
        '注意:载入的字体类型必须是正在运行该游戏的系统上存在的字体.否则将默认按"宋体"载入
        '小技巧:在大多数字体名前加上@可以让字体横过来~!
    
    Do While CWGameRun = True         '进入游戏循环
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '检测是否可以渲染（设备正常并且窗体未最小化时渲染）
    CWBeginScene CWColorRGBA(255, 240, 200)  '准备好绘制场景
        CWPaintPicBegin
        
        Dim rcos!, rsin!
        rcos = Cos(FontRoll)
        rsin = Sin(FontRoll)
        
        ' 注意：修改WorldTransform会对接下来的所有绘制产生改变，直到再次修改。
        With WorldTransform     ' 自定义世界变换（顺时针旋转）
            ' m11 为X缩放，m22为Y缩放，m12是X到Y切变，m21是Y到X切变。
            ' 当缩放为余弦（同±）且切变为正弦（异±）时，就可以实现旋转效果。
            ' 注：两个sin的±关系决定旋转方向，顺时针为+-，逆时针为-+。
            .m11 = rcos: .m12 = rsin
            .m21 = -rsin: .m22 = rcos
            .mdx = 200!: .mdy = 150!    ' 旋转后的平移量（注：这里才是真正的最终输出坐标）
        End With
        PrintFontEx FontDemo, CWHA_Red, CWGreen, CWBlue
        
        With WorldTransform     ' 自定义世界变换（逆时针旋转）
            .m11 = rcos: .m12 = -rsin
            .m21 = rsin: .m22 = rcos
            .mdx = 200!: .mdy = 350!    ' 旋转后的平移量（注：这里才是真正的最终输出坐标）
        End With
        PrintFontEx FontDemo, CWHA_Blue, CWRed, CWGreen
        
        ' 其它测试
        With WorldTransform     ' 自定义世界变换
            .m11 = rcos: .m12 = rsin
            .m21 = rsin: .m22 = rcos
            .mdx = 400!: .mdy = 150!
        End With
        PrintFontEx FontDemo, CWHA_Purple, CWYellow, CWCyan
        
        With WorldTransform     ' 自定义世界变换
            .m11 = rcos: .m12 = -rsin
            .m21 = -rsin: .m22 = rcos
            .mdx = 400!: .mdy = 350!
        End With
        PrintFontEx FontDemo, CWHA_Cyan, CWPurple, CWYellow
        
        With WorldTransform     ' 自定义世界变换
            .m11 = rcos: .m12 = rsin
            .m21 = rsin: .m22 = rcos
            .mdx = 600!: .mdy = 150!
            PrintFontEx FontDemo, CWHA_Orange, CWkelly, CWFuchsia
            .m21 = -.m21
            PrintFontEx FontDemo, CWHA_Violet, CWTurquoise, CWCyanine
        End With
        
        With WorldTransform     ' 自定义世界变换
            .m11 = rcos: .m12 = -rsin
            .m21 = -rsin: .m22 = rcos
            .mdx = 600!: .mdy = 350!
            PrintFontEx FontDemo, CWHA_kelly, CWFuchsia, CWOrange
            .m21 = -.m21
            PrintFontEx FontDemo, CWHA_Turquoise, CWViolet, CWCyanine
        End With
        
        With WorldTransform     ' 自定义世界变换
            .m11 = rcos: .m12 = 0!
            .m21 = 0!: .m22 = 1!
            .mdx = 200!: .mdy = 550!
        End With
        PrintFontEx FontDemo, CWHA_Green, CWBlue, CWRed
        
        With WorldTransform     ' 自定义世界变换
            .m11 = 1!: .m12 = 0!
            .m21 = 0!: .m22 = rcos
            .mdx = 400!: .mdy = 550!
        End With
        PrintFontEx FontDemo, CWHA_Yellow, CWCyan, CWPurple
        
        With WorldTransform     ' 自定义世界变换
            .m11 = rcos: .m12 = 0!
            .m21 = 0!: .m22 = rcos
            .mdx = 600!: .mdy = 550!
        End With
        PrintFontEx FontDemo, CWHA_kelly, CWCyanine, CWFuchsia
        
        FontRoll = FontRoll + 0.05!
        If FontRoll > Pi Then FontRoll = FontRoll - 2! * Pi
        
        ' 设置世界变换为单位矩阵（重置为原始状态）
        WorldTransform = MatrixIdentity
        CWPrintFontTop FontDemo, "FPS: " & CWFPS, 10, 10, 800, 60, CWViolet, CWF_LeftAl  '显示当前FPS
        
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

Private Sub PrintFontEx(FontDemo As CWFont, ByVal BackColor As CWColorConstants, ByVal TextColor As CWColorConstants, ByVal EdgeColor As CWColorConstants)
    Dim CX&, CY&    '计算文本显示大小，并绘制文本边框。
    CWCalcrFont FontDemo, TestText, CX, CY, False
    ' 提示：绘图函数的原始坐标为输入坐标（此坐标为旋转前，可用于定位旋转点）
    ' (0, 0)坐标表示绕左上角旋转，(-宽度, -高度)坐标表示绕右下角旋转，(-0.5 * 宽度, -0.5 * 高度)坐标表示绕中心点旋转。
    CWPrintFontTop FontDemo, TestText, CX * -0.5!, CY * -0.5!, CX, CY, TextColor, DT_CENTER Or DT_NOCLIP
    CX = CX + 6: CY = CY + 4
    CWDrawSRect CX * -0.5!, CY * -0.5!, CX, CY, BackColor
    CWDrawHRect CX * -0.5!, CY * -0.5!, CX, CY, EdgeColor
End Sub
