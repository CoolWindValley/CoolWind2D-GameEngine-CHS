Attribute VB_Name = "VBDX9BAS"
'******************************************************************************************************************
'
'   作者名单：
'           VBDX9BAS.BAS：秋枫萧萧（百度贴吧 VB吧 http://tieba.baidu.com/f?kw=vb）
'           VBDX9TLB.TLB：acme_pjz（VBGOOD论坛 http://www.vbgood.com）
'
'   技术支持：
'           秋枫萧萧：CoolWind2D原作者，建立引擎框架和实现基础功能
'           YY菌{3EA3E263-6945-4E1F-A573-492FB5A7799E}：修复了大量BUG和增加大量新功能
'
'   有任何的意见建议或者疑问可以到百度贴吧VB吧或者VBGOOD论坛发帖讨论，也可以加入CoolWind游戏编程研究会一起讨论
'           CoolWind游戏编程研究会 群号：112915633 欢迎各位游戏编程爱好者的加入
'******************************************************************************************************************

Option Explicit     '变量使用前必须声明

Public Declare Function CWSplitColor Lib "msvbvm60" Alias "#644" (ByVal Color As CWColorConstants) As CWColor
Public Declare Function GetMem4 Lib "msvbvm60" (ByVal Address As Long, ByRef Value As Any) As Long
Public Declare Function PutMem4 Lib "msvbvm60" (ByVal Address As Long, ByVal Value As Any) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function timeGetTime Lib "winmm" () As Long
Public Declare Function timeBeginPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Public Declare Function mciSendStringA Lib "winmm" (ByVal lpstrCommand As String, Optional ByVal lpstrReturnString As String, Optional ByVal uReturnLength As Long, Optional ByVal hwndCallback As Long) As Long
Public Declare Function mciSendStringW Lib "winmm" (ByVal lpstrCommand As Long, Optional ByVal lpstrReturnString As Long, Optional ByVal uReturnLength As Long, Optional ByVal hwndCallback As Long) As Long
Public Declare Function GetShortPathNameW Lib "kernel32" (ByVal LongPath As Long, ByVal ShortPath As Long, ByVal Length As Long) As Long
Public Declare Function GetShortPathNameA Lib "kernel32" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal DX As Long, ByVal DY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As D3DRECT) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As D3DRECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function joyGetNumDevs Lib "winmm" () As Long
Public Declare Function joyGetPosEx Lib "winmm" (ByVal uJoyID As Long, ByRef pji As JOYINFOEX) As Long
Public Declare Function CombineTransform Lib "gdi32" (ByRef MatOut As CWMatrix, MatLeft As CWMatrix, MatRight As CWMatrix) As Long

Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, ByRef Width As Long) As Long
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, ByRef Height As Long) As Long
Public Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal Bitmap As Long, Rect As Any, ByVal Flags As GpImageLockMode, ByVal Format As Long, LockedBitmapData As GpBitmapData) As Long
Public Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal Bitmap As Long, LockedBitmapData As GpBitmapData) As Long

Public Type POINTAPI    'API鼠标位置类
   X As Long
   Y As Long
End Type

Public Enum JOYFALGS    'API游戏手柄类
    JOY_RETURNX = &H1
    JOY_RETURNY = &H2
    JOY_RETURNZ = &H4
    JOY_RETURNR = &H8
    JOY_RETURNU = &H10
    JOY_RETURNV = &H20
    JOY_RETURNPOV = &H40
    JOY_RETURNBUTTONS = &H80
    JOY_RETURNRAWDATA = &H100
    JOY_RETURNPOVCTS = &H200
    JOY_RETURNCENTERED = &H400
    JOY_USEDEADZONE = &H800
    JOY_RETURNALL = JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS
End Enum

Public Enum JOYBUTTON    'API游戏手柄类
    JOY_BUTTON1 = &H1
    JOY_BUTTON2 = &H2
    JOY_BUTTON3 = &H4
    JOY_BUTTON4 = &H8
    JOY_BUTTON5 = &H10
    JOY_BUTTON6 = &H20
    JOY_BUTTON7 = &H40
    JOY_BUTTON8 = &H80
    JOY_BUTTON9 = &H100
    JOY_BUTTON10 = &H200
    JOY_BUTTON11 = &H400
    JOY_BUTTON12 = &H800
    JOY_BUTTON13 = &H1000
    JOY_BUTTON14 = &H2000
    JOY_BUTTON15 = &H4000
    JOY_BUTTON16 = &H8000
    JOY_BUTTON17 = &H10000
    JOY_BUTTON18 = &H20000
    JOY_BUTTON19 = &H40000
    JOY_BUTTON20 = &H80000
    JOY_BUTTON21 = &H100000
    JOY_BUTTON22 = &H200000
    JOY_BUTTON23 = &H400000
    JOY_BUTTON24 = &H800000
    JOY_BUTTON25 = &H1000000
    JOY_BUTTON26 = &H2000000
    JOY_BUTTON27 = &H4000000
    JOY_BUTTON28 = &H8000000
    JOY_BUTTON29 = &H10000000
    JOY_BUTTON30 = &H20000000
    JOY_BUTTON31 = &H40000000
    JOY_BUTTON32 = &H80000000
End Enum

Public Type JOYINFOEX    'API游戏手柄类
   dwSize As Long
   dwFlags As JOYFALGS
   dwXpos As Long
   dwYpos As Long
   dwZpos As Long
   dwRpos As Long
   dwUpos As Long
   dwVpos As Long
   dwButtons As JOYBUTTON
   dwButtonNumber As Long
   dwPOV As Long
   dwReserved1 As Long
   dwReserved2 As Long
End Type

Public Type Rect    '长整数矩形类
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type GpRect    'GDI+矩形类
    X As Long
    Y As Long
    Width As Long
    Height As Long
End Type

Const GpPixelFormat32bppARGB& = &H26200A

Public Enum GpImageLockMode
    GpImageLockModeRead = 1&
    GpImageLockModeWrite = 2&
    GpImageLockModeUserInputBuf = 4&
End Enum

Public Type GpBitmapData
    Width As Long
    Height As Long
    Stride As Long
    PixelFormat As Long
    Scan0 As Long
    Reserved As Long
End Type

Public Type D2DVector   '2D顶点类
    X As Single
    Y As Single
    Z As Single
    Rhw As Single
    Color As CWColorConstants
'    Specular As Long
'    Tu As Single
'    Tv As Single
End Type
 
Public Type CWMatrix   '2D矩阵
    m11 As Single: m12 As Single
    m21 As Single: m22 As Single
    mdx As Single: mdy As Single
End Type
 
Public Type CWColor   '颜色类
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Public Type CWPic       '图片类
    Tex As Direct3DTexture9
    PICSize As D3DRECT
End Type

Public Type CWFont      '字体类
    SNum As Long
End Type

Public Type CWMusic     '音乐类
    ID As Long
End Type

Public Type CWMusicObj     '音乐类
    IsLoop As Boolean
    mc As IMediaControl
    mp As IMediaPosition
    ba As IBasicAudio
    vw As IVideoWindow
    evt As IMediaEvent
    fx3d As IDirectSound3DBuffer
End Type

Public Enum CWMusicEffect
    CWME_3DFX = 1
End Enum

Public Type CWKeyState      '按键状态类
    PUP As Boolean
    PDown As Boolean
    PUPMoment As Boolean
    PDownMoment As Boolean
End Type

Public Type CWKeyStateSP        '鼠标滚轮状态类
    PUP As Boolean
    PDown As Boolean
    PUPMoment As Boolean
    PDownMoment As Boolean
    RollUP As Boolean
    RollDown As Boolean
End Type

Public Type CWMouseState        '鼠标类
    X As Single
    Y As Single
    LeftKey As CWKeyState
    RightKey As CWKeyState
    MidKey As CWKeyStateSP
    BackKey As CWKeyState
    ForwardKey As CWKeyState
End Type

Public Type CWKeyboardState     '键盘类
    ESC As CWKeyState
    F1 As CWKeyState
    F2 As CWKeyState
    F3 As CWKeyState
    F4 As CWKeyState
    F5 As CWKeyState
    F6 As CWKeyState
    F7 As CWKeyState
    F8 As CWKeyState
    F9 As CWKeyState
    F10 As CWKeyState
    F11 As CWKeyState
    F12 As CWKeyState
    Insert As CWKeyState
    Delete As CWKeyState
    PageUp As CWKeyState
    PageDown As CWKeyState
    Home As CWKeyState
    End As CWKeyState
    UP As CWKeyState
    Down As CWKeyState
    Left As CWKeyState
    Right As CWKeyState
    Tab As CWKeyState
    Shift As CWKeyState
    Ctrl As CWKeyState
    Alt As CWKeyState
    Space As CWKeyState
    BackSpace As CWKeyState
    Enter As CWKeyState
    Num1 As CWKeyState
    Num2 As CWKeyState
    Num3 As CWKeyState
    Num4 As CWKeyState
    Num5 As CWKeyState
    Num6 As CWKeyState
    Num7 As CWKeyState
    Num8 As CWKeyState
    Num9 As CWKeyState
    Num0 As CWKeyState
    A As CWKeyState
    B As CWKeyState
    C As CWKeyState
    D As CWKeyState
    E As CWKeyState
    F As CWKeyState
    G As CWKeyState
    H As CWKeyState
    I As CWKeyState
    j As CWKeyState
    K As CWKeyState
    L As CWKeyState
    M As CWKeyState
    N As CWKeyState
    O As CWKeyState
    P As CWKeyState
    Q As CWKeyState
    R As CWKeyState
    S As CWKeyState
    T As CWKeyState
    U As CWKeyState
    V As CWKeyState
    W As CWKeyState
    X As CWKeyState
    Y As CWKeyState
    Z As CWKeyState
End Type

Public Type CWJoystickState        '手柄类
    IsConnected As Boolean
    IsPov As Boolean
    X As Single
    Y As Single
    Z As Single
    R As Single
    Pov As Single
    Btn(1 To 30) As CWKeyState
End Type

Enum CWSpriteState
    Ended
    Begined
    Drawed
End Enum

Enum CWLinePattern
    CWLP_Transparent
    CWLP_Solid = &HFFFFFFFF
    CWLP_Dash = &H7E7E7E7E
    CWLP_Dot = &H66666666
    CWLP_DashDot = &H87E187E1
    CWLP_DashDotDot = &H67E667E6
    CWLP_Minus = &H3C3C3C3C
    CWLP_DashMinus = &HE3C7E3C7
    CWLP_MinusDot = &HBDBDBDBD
    CWLP_MinusDotDot = &HC663C663
    CWLP_Point = &HAAAAAAAA
    CWLP_InvPoint = &H55555555
    CWLP_DotPointPoint = &HA5A5A5A5
End Enum

Public WorldTransform As CWMatrix
Public MatrixIdentity As CWMatrix, CWSpState As CWSpriteState
Public CWD3D9 As Direct3D9, CWD3DDevice9 As Direct3DDevice9, CWSprite As D3DXSprite, CWSpriteSP As D3DXSprite
Public CWLine As D3DXLine
Public CWGameRun As Boolean '运行状态
Public CWDpp9 As D3DPRESENT_PARAMETERS '画布
Public CWD3Dc9 As D3DCAPS9  '设备特性
Public CWWindowSwitch As Boolean '全屏窗口切换

Public CWLongTime As Long, CWFrameCount As Long, CWTimeNow As Long
Public CWFPS As Integer

Public CWFrm As Object, CWHwnd As Long, CWDModelX As Integer, CWDModelY As Integer, CWDModelW As Integer
Public CWFrmHei As Long, CWFrmWid As Long, CWFrmSHei As Long, CWFrmSWid As Long, CWMTempC As Long

Public CWFontList() As D3DXFont, CWFontNum As Long  '文字处理列表
Public CWMusicList() As CWMusicObj, CWMusicNum As Long '音乐处理列表

Public IsActive As Boolean, IsHitWnd As Boolean
Public CWMouse As CWMouseState
Public CWKeyboard As CWKeyboardState
Public CWJoystick() As CWJoystickState

Public CWP_PubRollCD As D3DVECTOR   '精灵贴图常用固定值

Enum CWDisplayModel     '显示模式常量
    CW_Windowed = 1
    CW_FullScreen = 0
End Enum

Enum CWFAlign               '文字对齐常量
    CWF_LeftAl = DT_LEFT Or DT_WORDBREAK
    CWF_RightAl = DT_RIGHT Or DT_WORDBREAK
    CWF_CenterAl = DT_CENTER Or DT_WORDBREAK
End Enum

Enum CWFBStyle              '字体粗细常量
    CWF_Normal = 400
    CWF_Bold = 700
End Enum

Enum CWMPModel              '音乐播放模式常量
    CWM_Resume
    CWM_Once
    CWM_Repeat
    CWM_Restart
End Enum

Public Const CWP_FVFConst As Long = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE
Public Const CWP_SpriteConst As Long = D3DXSPRITE_DONOTMODIFY_RENDERSTATE Or D3DXSPRITE_DONOTSAVESTATE  '精灵参数常量
Public Const Pi As Single = 3.14159265358979                  'π常量
Public Const CWKD As Boolean = True
Public Const CWKU As Boolean = False
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const RightButtonUpDown = &H18

'*********************常用颜色列表***************************
Enum CWColorConstants
    CWColorNone = &H0          '完全无色
    CWTransparent = &HFFFFFF    '透明有色（方便位运算截取颜色部分）
    
    CWBlack = &HFF000000        '黑色
    CWWhite = &HFFFFFFFF        '白色
    CWGrey = &HFF808080         '灰色
    CWRed = &HFFFF0000          '红色
    CWGreen = &HFF00FF00        '绿色
    CWBlue = &HFF0000FF         '蓝色
    CWYellow = &HFFFFFF00       '黄色
    CWPurple = &HFFFF00FF       '紫色
    CWCyan = &HFF00FFFF         '青色
    CWOrange = &HFFFF8000       '橙色
    CWkelly = &HFF80FF00        '黄绿色
    CWFuchsia = &HFFFF0080      '紫红色
    CWViolet = &HFF8000FF       '蓝紫色
    CWTurquoise = &HFF00FF80    '青绿色
    CWCyanine = &HFF0080FF      '青蓝色
    
    CWHA_Black = &H80000000     '半透明黑色
    CWHA_White = &H80FFFFFF     '半透明白色
    CWHA_Grey = &H80808080      '半透明灰色
    CWHA_Red = &H80FF0000       '半透明红色
    CWHA_Green = &H8000FF00     '半透明绿色
    CWHA_Blue = &H800000FF      '半透明蓝色
    CWHA_Yellow = &H80FFFF00    '半透明黄色
    CWHA_Purple = &H80FF00FF    '半透明紫色
    CWHA_Cyan = &H8000FFFF      '半透明青色
    CWHA_Orange = &H80FF8000    '橙色
    CWHA_kelly = &H8080FF00     '黄绿色
    CWHA_Fuchsia = &H80FF0080   '紫红色
    CWHA_Violet = &H808000FF    '蓝紫色
    CWHA_Turquoise = &H8000FF80 '青绿色
    CWHA_Cyanine = &H800080FF   '青蓝色
End Enum
'其他颜色请参考RGB颜色列表用CWColorARGB函数自行转换
'*********************常用颜色列表***************************

Public Sub CWVBDX9Initialization(ByVal Frm As Object, ByVal ScrWidth As Integer, ByVal ScrHeight As Integer, ByVal IniState As CWDisplayModel, Optional ByVal Zoom As Single = 1!)
Dim I As Integer, Anti As Integer, JSorH As Long

    If App.PrevInstance = True Then
    MsgBox "游戏已经在运行", vbInformation, "重复运行"
    End
    Exit Sub
    End If
timeBeginPeriod 1
On Error GoTo CWFIniEDH

Frm.ScaleMode = 3     '窗口显示区大小按像素计算
Frm.BorderStyle = 0
Frm.Caption = Frm.Caption
Frm.Width = ScrWidth * Zoom * 15 '保证窗口显示区符合即将初始化的引擎分辨率，防止图形失真
Frm.Height = ScrHeight * Zoom * 15

If IniState = CW_Windowed Then
    Frm.BorderStyle = 1
    Frm.Caption = Frm.Caption
End If

Set CWFrm = Frm
CWHwnd = Frm.hWnd
CWDModelX = ScrWidth
CWDModelY = ScrHeight
CWDModelW = IniState
Frm.Show
CWFrmHei = Frm.Height
CWFrmSHei = Frm.ScaleHeight
CWFrmWid = Frm.Width
CWFrmSWid = Frm.ScaleWidth
CWMTempC = (CWDModelX * 15 - CWFrmWid) / 2

With MatrixIdentity
    .m11 = 1!: .m22 = 1! ': .m33 = 1!: .m44 = 1!
End With
WorldTransform = MatrixIdentity

On Error GoTo CWIniEHD
      Set CWD3D9 = Direct3DCreate9(D3D_SDK_VERSION)
      
    CWDpp9.BackBufferWidth = ScrWidth
    CWDpp9.BackBufferHeight = ScrHeight
    CWDpp9.Windowed = IniState
    CWDpp9.SwapEffect = D3DSWAPEFFECT_DISCARD
    CWDpp9.BackBufferCount = 1
    CWDpp9.BackBufferFormat = D3DFMT_X8R8G8B8
    CWDpp9.hDeviceWindow = CWHwnd
    CWDpp9.PresentationInterval = D3DPRESENT_INTERVAL_ONE        '开启垂直同步

'    On Error Resume Next
'    Anti = 0
'    For i = 16 To 2 Step -2         '检测抗锯齿倍数
'    Err.Clear
'     CWD3D9.CheckDeviceMultiSampleType 0, D3DDEVTYPE_HAL, D3DFMT_A8R8G8B8, 1, i
'     CWD3D9.CheckDeviceMultiSampleType 0, D3DDEVTYPE_HAL, D3DFMT_A8R8G8B8, 0, i
'        If Err.Number = 0 Then
'        Anti = i
'        Exit For
'        End If
'    Next
'    If Anti > 0 Then CWDpp9.MultiSampleType = Anti           '抗锯齿

On Error GoTo CWIniEHD

    CWD3D9.GetDeviceCaps 0, D3DDEVTYPE_HAL, CWD3Dc9
    JSorH = 0
    If CWD3Dc9.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT Then        '检测是否支持硬件顶点渲染
        If (CWD3Dc9.VertexShaderVersion And &HFFFF&) >= &H200& Then
            JSorH = D3DCREATE_HARDWARE_VERTEXPROCESSING
        End If
    End If
    If JSorH = 0 Then
    JSorH = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    CWD3Dc9.VertexShaderVersion = 0
    End If

     Set CWD3DDevice9 = CWD3D9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, CWHwnd, JSorH, CWDpp9)

    CWD3DDevice9.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    ' Alpha 混合的效果比 Alpha 测试好，但 Alpha 测试是直接剔除不透明像素，不需要做混合运算，可以提高性能。

    ' Alpha 混合（实现透明度功能，可以全透明、不透明、半透明）
    CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, True
    CWD3DDevice9.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    CWD3DDevice9.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    ' Alpha 测试（只能实现全透明和不透明，不能半透明，但是效率比 Alpha 混合高）
    CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, True
    CWD3DDevice9.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
    CWD3DDevice9.SetRenderState D3DRS_ALPHAREF, 1

    CWD3DDevice9.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
    CWD3DDevice9.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
    CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_BORDER
    CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_BORDER
    CWD3DDevice9.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

    CWD3DDevice9.SetRenderState D3DRS_LIGHTING, False  '描绘时不必使用光线
    CWD3DDevice9.SetRenderState D3DRS_ZWRITEENABLE, False   '描绘时不必使用Z-Buffer
    
        D3DXCreateSprite CWD3DDevice9, CWSprite
        D3DXCreateSprite CWD3DDevice9, CWSpriteSP
        D3DXCreateLine CWD3DDevice9, CWLine

        CWP_PubRollCD.X = 0
        CWP_PubRollCD.Y = 0
        CWP_PubRollCD.Z = 0
        
        CWFPS = 60
        CWFrameCount = 60
        CWLongTime = 0
        
    ReDim CWFontList(0)
        CWFontNum = 0
        CWMusicNum = 0
    
            CWWindowSwitch = False
            CWGameRun = True   '游戏状态设置为运行

    ReDim CWJoystick(0 To joyGetNumDevs() - 1)
Exit Sub
CWIniEHD:
    MsgBox "引擎初始化失败，请确认您完整安装了Directx 9.0c最新版。如仍然出现该问题请确认您的显卡是否兼容DirectX 9.0c。", vbInformation, "初始化错误"
    End
Exit Sub

CWFIniEDH:
    MsgBox "窗体初始化失败，作为CoolWind引擎初始化的窗体必须是具有句柄和客户区的对象。", vbInformation, "初始化错误"
    End

End Sub

Public Sub CWBeginScene(Optional ByVal BackColor As CWColorConstants)
    If CWGameRun = True Then
        IsActive = GetActiveWindow() = CWHwnd
        CWMouseCheck
        If IsActive Then
            CWKeyboardCheck
            CWJoystickCheck
        End If
    End If
    
    CWD3DDevice9.Clear 0, ByVal 0, D3DCLEAR_TARGET, BackColor, 0!, 0
    CWD3DDevice9.BeginScene

End Sub

Public Sub CWPresentScene()
    CWD3DDevice9.EndScene '场景结束
        
    DoEvents    '让系统能够处理其他信息（如获取键盘、鼠标状态等）
    CWMediaLoopRepair
    CWD3DDevice9.Present ByVal 0, ByVal 0, 0, ByVal 0   '刷新
        
    CWGetFPS    '获取FPS
    If CWFPS > 86 Then Sleep 10  '窗口被完全遮挡时降低CPU占用
    
    If CWWindowSwitch = True Then
        CWWindowSw
    End If
End Sub

Public Sub SaveScreenShot(file_name As String, Optional ByVal Format As D3DXIMAGE_FILEFORMAT = D3DXIFF_PNG) '截屏

    Dim BackBuffer As Direct3DSurface9

    Set BackBuffer = CWD3DDevice9.GetBackBuffer(0, 0, D3DBACKBUFFER_TYPE_MONO)

    D3DXSaveSurfaceToFileW file_name, Format, BackBuffer, ByVal 0&, ByVal 0&

    Set BackBuffer = Nothing
       
'''    enum D3DXIMAGE_FILEFORMAT
'''    D3DXIFF_BMP = 0
'''    D3DXIFF_JPG = 1
'''    D3DXIFF_TGA = 2
'''    D3DXIFF_PNG = 3
'''    D3DXIFF_DDS = 4
'''    D3DXIFF_PPM = 5
'''    D3DXIFF_DIB = 6
'''    D3DXIFF_HDR = 7
'''    D3DXIFF_PFM = 8
'''     end enum
End Sub

Public Sub CWResetDevice()      '重置设备
Dim I As Long
On Error GoTo CWResetEHD

If CWD3DDevice9.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then  '检测到设备丢失
        CWSprite.OnLostDevice
        CWSpriteSP.OnLostDevice
        CWLine.OnLostDevice
    If CWFontNum > 0 Then
    For I = 1 To CWFontNum
        CWFontList(I).OnLostDevice
    Next
    End If
    CWD3DDevice9.Reset CWDpp9
    
    CWD3DDevice9.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, True
    CWD3DDevice9.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA         '描绘时开启透明色
    CWD3DDevice9.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, True
    CWD3DDevice9.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
    CWD3DDevice9.SetRenderState D3DRS_ALPHAREF, 1

    CWD3DDevice9.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
    CWD3DDevice9.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
    CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_BORDER
    CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_BORDER
    CWD3DDevice9.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

    CWD3DDevice9.SetRenderState D3DRS_LIGHTING, False  '描绘时不必使用光线
    CWD3DDevice9.SetRenderState D3DRS_ZWRITEENABLE, False   '描绘时不必使用Z-Buffer
    CWSprite.OnResetDevice           '由于纹理、字体已经设置为系统托管，暂时用不上
    CWSpriteSP.OnResetDevice
    CWLine.OnResetDevice
    For I = 1 To CWFontNum
      CWFontList(I).OnResetDevice
    Next
End If

    Sleep 1
    DoEvents

Exit Sub
CWResetEHD:
MsgBox "修复设备失败", vbInformation, "设备丢失"
End Sub


Public Sub CWWindowSw()
Dim I As Long

On Error GoTo CWResetEHD
    Select Case CWDpp9.Windowed
    Case CW_Windowed
    CWDpp9.Windowed = CW_FullScreen
    CWDModelW = CW_FullScreen
    CWFrm.BorderStyle = 0
    CWFrm.Caption = CWFrm.Caption
    
    Case CW_FullScreen
    CWDpp9.Windowed = CW_Windowed
    CWDModelW = CW_Windowed
    CWFrm.BorderStyle = 1
    CWFrm.Caption = CWFrm.Caption
    
    End Select

        CWSprite.OnLostDevice
        CWSpriteSP.OnLostDevice
        CWLine.OnLostDevice
    If CWFontNum > 0 Then
    For I = 1 To CWFontNum
        CWFontList(I).OnLostDevice
    Next
    End If
    CWD3DDevice9.Reset CWDpp9
    
    CWD3DDevice9.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, True
    CWD3DDevice9.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA         '描绘时开启透明色
    CWD3DDevice9.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, True
    CWD3DDevice9.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
    CWD3DDevice9.SetRenderState D3DRS_ALPHAREF, 1

    CWD3DDevice9.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
    CWD3DDevice9.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
    CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_BORDER
    CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_BORDER
    CWD3DDevice9.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

    CWD3DDevice9.SetRenderState D3DRS_LIGHTING, False  '描绘时不必使用光线
    CWD3DDevice9.SetRenderState D3DRS_ZWRITEENABLE, False   '描绘时不必使用Z-Buffer
    CWSprite.OnResetDevice           '由于纹理、字体已经设置为系统托管，暂时用不上
    CWSpriteSP.OnResetDevice
    CWLine.OnResetDevice
    For I = 1 To CWFontNum
      CWFontList(I).OnResetDevice
    Next
   
   If CWDModelW = CW_Windowed Then
     CWFrm.Left = CWFrm.Left + 300
     CWFrm.Top = CWFrm.Top + 300
     CWFrm.Height = CWFrmHei
     CWFrm.Width = CWFrmWid
   End If
   
CWWindowSwitch = False
Exit Sub
CWResetEHD:
MsgBox "全屏/窗口模式切换失败", vbInformation, "设备丢失"
CWWindowSwitch = False
End Sub

Public Sub CWWinFullScrSwitch()
CWWindowSwitch = True
End Sub

Public Sub CWVBDX9Destory()
    Dim I As Long
    'mciSendStringW StrPtr("close all")
    Set CWSprite = Nothing
    Set CWSpriteSP = Nothing
        If CWFontNum > 0 Then
        For I = 1 To CWFontNum
            Set CWFontList(I) = Nothing
        Next
        End If
    Set CWD3DDevice9 = Nothing
    Set CWD3D9 = Nothing
    timeEndPeriod 1
End Sub

Public Sub CWLoadPic(Pic As CWPic, ByVal PicPath As String, Optional ByVal SColor As CWColorConstants)
    Dim DXInfo As D3DXIMAGE_INFO
    On Error GoTo CWLPEHD
    D3DXCreateTextureFromFileExW CWD3DDevice9, PicPath, D3DX_DEFAULT_NONPOW2, D3DX_DEFAULT_NONPOW2, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, SColor, DXInfo, ByVal 0, Pic.Tex
    Pic.PICSize.X1 = 0
    Pic.PICSize.Y1 = 0
    Pic.PICSize.X2 = DXInfo.Width
    Pic.PICSize.Y2 = DXInfo.Height
    
    If Pic.PICSize.X2 <> 0 And Pic.PICSize.Y2 <> 0 Then
        Exit Sub
    End If
    
CWLPEHD:
    MsgBox "找不到图片或不支持的图片格式", vbInformation, "纹理初始化失败"
    End
End Sub

Public Sub CWMapPic(Pic As CWPic, PicSrc As CWPic, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcW As Long, ByVal SrcH As Long)
    Pic.PICSize.X1 = PicSrc.PICSize.X1 + SrcX
    Pic.PICSize.Y1 = PicSrc.PICSize.Y1 + SrcY
    Pic.PICSize.X2 = Pic.PICSize.X1 + SrcW
    Pic.PICSize.Y2 = Pic.PICSize.Y1 + SrcH
    Set Pic.Tex = PicSrc.Tex
End Sub

Public Sub CWLoadPicFromGDIP(Pic As CWPic, ByVal GpBmp As Long)
    Dim GpBmpDat As GpBitmapData, lrc As D3DLOCKED_RECT
    On Error GoTo CWLPEHD
    With GpBmpDat
        If GdipGetImageWidth(GpBmp, .Width) Then GoTo CWLPEHD
        If GdipGetImageHeight(GpBmp, .Height) Then GoTo CWLPEHD
        D3DXCreateTexture CWD3DDevice9, .Width, .Height, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, Pic.Tex
        Pic.Tex.LockRect 0, lrc, ByVal 0&, D3DLOCK_DISCARD Or D3DLOCK_DONOTWAIT Or D3DLOCK_NO_DIRTY_UPDATE Or D3DLOCK_NOSYSLOCK
        .Stride = lrc.Pitch
        .Scan0 = lrc.pBits
        .PixelFormat = GpPixelFormat32bppARGB
        If GdipBitmapLockBits(GpBmp, ByVal 0&, GpImageLockModeRead Or GpImageLockModeUserInputBuf, .PixelFormat, GpBmpDat) Then GoTo CWLPEHD
        GdipBitmapUnlockBits GpBmp, GpBmpDat
        Pic.Tex.UnlockRect 0
        Pic.PICSize.X1 = 0
        Pic.PICSize.Y1 = 0
        Pic.PICSize.X2 = .Width
        Pic.PICSize.Y2 = .Height
    End With
    
    If Pic.PICSize.X2 <> 0 And Pic.PICSize.Y2 <> 0 Then
        Exit Sub
    End If
    
CWLPEHD:
    MsgBox "非GDI+位图或不支持的像素格式", vbInformation, "纹理初始化失败"
    End
End Sub

Public Property Get CWPicGetPixel(Pic As CWPic, ByVal X As Long, ByVal Y As Long) As CWColorConstants
    X = X + Pic.PICSize.X1
    Y = Y + Pic.PICSize.Y1
    If 0 = PtInRect(Pic.PICSize, X, Y) Then Exit Property
    
    Dim rc As D3DRECT, lrc As D3DLOCKED_RECT
    rc.X1 = X: rc.Y1 = Y
    rc.X2 = X: rc.Y2 = Y
    Pic.Tex.LockRect 0, lrc, rc, D3DLOCK_READONLY Or D3DLOCK_DONOTWAIT Or D3DLOCK_NO_DIRTY_UPDATE Or D3DLOCK_NOSYSLOCK
    GetMem4 lrc.pBits, CWPicGetPixel
    Pic.Tex.UnlockRect 0
End Property

Public Property Let CWPicSetPixel(Pic As CWPic, ByVal X As Long, ByVal Y As Long, ByVal Value As CWColorConstants)
    X = X + Pic.PICSize.X1
    Y = Y + Pic.PICSize.Y1
    If 0 = PtInRect(Pic.PICSize, X, Y) Then Exit Property
    
    Dim rc As D3DRECT, lrc As D3DLOCKED_RECT
    rc.X1 = X: rc.Y1 = Y
    rc.X2 = X: rc.Y2 = Y
    Pic.Tex.LockRect 0, lrc, rc, D3DLOCK_DISCARD Or D3DLOCK_DONOTWAIT Or D3DLOCK_NO_DIRTY_UPDATE Or D3DLOCK_NOSYSLOCK
    PutMem4 lrc.pBits, Value
    Pic.Tex.UnlockRect 0
End Property

' 绘制图片（不缩放、不裁剪、不旋转）
Public Sub CWPaintPic(Pic As CWPic, ByVal PaintX As Single, ByVal PaintY As Single, Optional ByVal HColor As CWColorConstants = CWWhite)
    If Pic.Tex Is Nothing Then Exit Sub
    Dim TexMatrix As D3DMATRIX
    
    With TexMatrix
        .m11 = WorldTransform.m11
        .m12 = WorldTransform.m12
        .m21 = WorldTransform.m21
        .m22 = WorldTransform.m22
        .m33 = 1!
        .m44 = 1!
        TransformationEx .m41, .m42, PaintX, PaintY, WorldTransform
    End With
    
    CWSprite.SetTransform TexMatrix
    CWSprite.Draw Pic.Tex, Pic.PICSize, CWP_PubRollCD, CWP_PubRollCD, HColor
    CWSpState = Drawed
End Sub

' 绘制图片（缩放、不裁剪、不旋转）
Public Sub CWPaintPic2(Pic As CWPic, ByVal PaintX As Single, ByVal PaintY As Single, ByVal PaintWidth As Single, ByVal PaintHeight As Single, Optional ByVal HColor As CWColorConstants = CWWhite)
    If Pic.Tex Is Nothing Then Exit Sub
    Dim TexMatrix As D3DMATRIX
    
    With TexMatrix
        .m11 = PaintWidth / (Pic.PICSize.X2 - Pic.PICSize.X1)
        .m22 = PaintHeight / (Pic.PICSize.Y2 - Pic.PICSize.Y1)
        .m33 = 1!
        .m44 = 1!
        .m12 = .m11 * WorldTransform.m12
        .m21 = .m22 * WorldTransform.m21
        .m11 = .m11 * WorldTransform.m11
        .m22 = .m22 * WorldTransform.m22
        TransformationEx .m41, .m42, PaintX, PaintY, WorldTransform
    End With
    
    CWSprite.SetTransform TexMatrix
    CWSprite.Draw Pic.Tex, Pic.PICSize, CWP_PubRollCD, CWP_PubRollCD, HColor
    CWSpState = Drawed
End Sub

' 绘制图片（不缩放、裁剪、不旋转）
Public Sub CWPaintPicEx(Pic As CWPic, ByVal PaintX As Single, ByVal PaintY As Single, ByVal CutX As Integer, ByVal CutY As Integer, ByVal CutWidth As Integer, ByVal CutHeight As Integer, Optional ByVal HColor As CWColorConstants = CWWhite)
    If Pic.Tex Is Nothing Then Exit Sub
    Dim TexMatrix As D3DMATRIX, TexCut As D3DRECT
    
    With TexMatrix
        .m11 = WorldTransform.m11
        .m12 = WorldTransform.m12
        .m21 = WorldTransform.m21
        .m22 = WorldTransform.m22
        .m33 = 1!
        .m44 = 1!
        TransformationEx .m41, .m42, PaintX, PaintY, WorldTransform
    End With
    With TexCut
        .X1 = Pic.PICSize.X1 + CutX
        .Y1 = Pic.PICSize.Y1 + CutY
        .X2 = .X1 + CutWidth
        .Y2 = .Y1 + CutHeight
    End With
    
    CWSprite.SetTransform TexMatrix
    CWSprite.Draw Pic.Tex, TexCut, CWP_PubRollCD, CWP_PubRollCD, HColor
    CWSpState = Drawed
End Sub

' 绘制图片（缩放、裁剪、不旋转）
Public Sub CWPaintPicEx2(Pic As CWPic, ByVal PaintX As Single, ByVal PaintY As Single, ByVal PaintWidth As Single, ByVal PaintHeight As Single, ByVal CutX As Integer, ByVal CutY As Integer, ByVal CutWidth As Integer, ByVal CutHeight As Integer, Optional ByVal HColor As CWColorConstants = CWWhite)
    If Pic.Tex Is Nothing Then Exit Sub
    Dim TexMatrix As D3DMATRIX, TexCut As D3DRECT
    
    With TexMatrix
        .m11 = PaintWidth / CutWidth
        .m22 = PaintHeight / CutHeight
        .m33 = 1!
        .m44 = 1!
        .m12 = .m11 * WorldTransform.m12
        .m21 = .m22 * WorldTransform.m21
        .m11 = .m11 * WorldTransform.m11
        .m22 = .m22 * WorldTransform.m22
        TransformationEx .m41, .m42, PaintX, PaintY, WorldTransform
    End With
    With TexCut
        .X1 = Pic.PICSize.X1 + CutX
        .Y1 = Pic.PICSize.Y1 + CutY
        .X2 = .X1 + CutWidth
        .Y2 = .Y1 + CutHeight
    End With
    
    CWSprite.SetTransform TexMatrix
    CWSprite.Draw Pic.Tex, TexCut, CWP_PubRollCD, CWP_PubRollCD, HColor
    CWSpState = Drawed
End Sub

' 绘制图片（缩放、裁剪、旋转）
Public Sub CWPaintPicExEx(Pic As CWPic, ByVal PaintX As Single, ByVal PaintY As Single, ByVal CutX As Integer, ByVal CutY As Integer, ByVal CutWidth As Integer, ByVal CutHeight As Integer _
, Optional ByVal ZoomX As Single = 1, Optional ByVal ZoomY As Single = 1, Optional ByVal RollX As Single, Optional ByVal RollY As Single, Optional ByVal RollAngle As Single, Optional ByVal HColor As CWColorConstants = CWWhite)
    PaintX = RollX - PaintX
    PaintY = RollY - PaintY
    Call CWPaintPicFull(Pic, RollX, RollY, CutX, CutY, CutWidth, CutHeight, ZoomX, ZoomY, PaintX, PaintY, , PaintX, PaintY, RollAngle, HColor)
End Sub

' 绘制图片（先旋转后缩放，可裁剪）
Public Sub CWPaintPicExEx1(Pic As CWPic, ByVal PaintX As Single, ByVal PaintY As Single, ByVal CutX As Integer, ByVal CutY As Integer, ByVal CutWidth As Integer, ByVal CutHeight As Integer _
, Optional ByVal ZoomX As Single = 1, Optional ByVal ZoomY As Single = 1, Optional ByVal CenterX As Single, Optional ByVal CenterY As Single, Optional ByVal RollAngle As Single, Optional ByVal HColor As CWColorConstants = CWWhite)
    
    Call CWPaintPicFull(Pic, PaintX, PaintY, CutX, CutY, CutWidth, CutHeight, ZoomX, ZoomY, CenterX, CenterY, -RollAngle, CenterX, CenterY, RollAngle, HColor)
End Sub

' 绘制图片（先缩放后旋转，可裁剪）
Public Sub CWPaintPicExEx2(Pic As CWPic, ByVal PaintX As Single, ByVal PaintY As Single, ByVal CutX As Integer, ByVal CutY As Integer, ByVal CutWidth As Integer, ByVal CutHeight As Integer _
, Optional ByVal ZoomX As Single = 1, Optional ByVal ZoomY As Single = 1, Optional ByVal CenterX As Single, Optional ByVal CenterY As Single, Optional ByVal RollAngle As Single, Optional ByVal HColor As CWColorConstants = CWWhite)
    
    Call CWPaintPicFull(Pic, PaintX, PaintY, CutX, CutY, CutWidth, CutHeight, ZoomX, ZoomY, CenterX, CenterY, , CenterX, CenterY, RollAngle, HColor)
End Sub

' 绘制图片（完整版：缩放和旋转独立，不存在先后关系，可裁剪）
Public Sub CWPaintPicFull(Pic As CWPic, ByVal PaintX As Single, ByVal PaintY As Single, ByVal CutX As Integer, ByVal CutY As Integer, ByVal CutWidth As Integer, ByVal CutHeight As Integer _
, Optional ByVal ZoomX As Single = 1, Optional ByVal ZoomY As Single = 1, Optional ByVal ZoomCX As Single, Optional ByVal ZoomCY As Single, Optional ByVal ZoomAngle As Single _
, Optional ByVal RollX As Single, Optional ByVal RollY As Single, Optional ByVal RollAngle As Single, Optional ByVal HColor As CWColorConstants = CWWhite)
    If Pic.Tex Is Nothing Then Exit Sub
    Dim TexMatrix As D3DMATRIX, TexCut As D3DRECT
    Dim TexCoordinate As D3DXVECTOR2, ScaleCoordinate As D3DXVECTOR2, ScaleCenterCoordinate As D3DXVECTOR2, RollCoordinate As D3DXVECTOR2
    
    With TexCut
        .X1 = Pic.PICSize.X1 + CutX
        .Y1 = Pic.PICSize.Y1 + CutY
        .X2 = .X1 + CutWidth
        .Y2 = .Y1 + CutHeight
    End With
    With TexCoordinate
        .X = PaintX - RollX
        .Y = PaintY - RollY
    End With
    With ScaleCoordinate
        .X = ZoomX
        .Y = ZoomY
    End With
    With ScaleCenterCoordinate
        .X = ZoomCX
        .Y = ZoomCY
    End With
    With RollCoordinate '设置图片的转动轴坐标
        .X = RollX
        .Y = RollY
    End With
    D3DXMatrixTransformation2D TexMatrix, ScaleCenterCoordinate, ZoomAngle, ScaleCoordinate, RollCoordinate, RollAngle, TexCoordinate
    With TexMatrix
        ZoomX = .m11: RollX = .m12
        RollY = .m21: ZoomY = .m22
        PaintX = .m41: PaintY = .m42
        .m11 = ZoomX * WorldTransform.m11 + RollX * WorldTransform.m21
        .m12 = ZoomX * WorldTransform.m12 + RollX * WorldTransform.m22
        .m21 = RollY * WorldTransform.m11 + ZoomY * WorldTransform.m21
        .m22 = RollY * WorldTransform.m12 + ZoomY * WorldTransform.m22
        TransformationEx .m41, .m42, PaintX, PaintY, WorldTransform
    End With
    
    CWSprite.SetTransform TexMatrix
    CWSprite.Draw Pic.Tex, TexCut, CWP_PubRollCD, CWP_PubRollCD, HColor
    CWSpState = Drawed
End Sub

Public Sub CWDrawPoint(ByVal OX As Single, ByVal OY As Single, ByVal CColor As CWColorConstants)
    Dim Vector2D As D2DVector
    With Vector2D
        TransformationEx .X, .Y, OX, OY, WorldTransform
        .Rhw = 1!
        .Color = CColor
    End With

    CWPaintPicFlush ' 要先提交精灵的绘制，再画图形（否则层级会有问题）
    CWD3DDevice9.SetTexture 0, Nothing
    CWD3DDevice9.SetFVF CWP_FVFConst
    CWD3DDevice9.DrawPrimitiveUP D3DPT_POINTLIST, 1, Vector2D, LenB(Vector2D)
End Sub

Public Sub CWDrawLine(ByVal OX As Single, ByVal OY As Single, ByVal DX As Single, ByVal DY As Single, ByVal CColor As CWColorConstants)
    Call CWDrawLineEx(OX, OY, DX, DY, 1!, 1!, CColor, CColor)
End Sub

Public Sub CWDrawLineEx(ByVal OX As Single, ByVal OY As Single, ByVal DX As Single, ByVal DY As Single, ByVal ORHW As Single, ByVal DRHW As Single, ByVal OColor As CWColorConstants, ByVal DColor As CWColorConstants)
    Dim X!, Y!, S!, Vector2D(0 To 1) As D2DVector
    Transformation OX, OY, WorldTransform
    Transformation DX, DY, WorldTransform
    X = (OX - DX): Y = (OY - DY)
    S = Abs2(X, Y)
    X = 0.5! * X / S: Y = 0.5! * Y / S
    With Vector2D(0)
        .X = OX + X
        .Y = OY + Y
        .Rhw = ORHW
        .Color = OColor
    End With
    With Vector2D(1)
        .X = DX - X
        .Y = DY - Y
        .Rhw = DRHW
        .Color = DColor
    End With

    CWPaintPicFlush ' 要先提交精灵的绘制，再画图形（否则层级会有问题）
    CWD3DDevice9.SetTexture 0, Nothing
    CWD3DDevice9.SetFVF CWP_FVFConst
    CWD3DDevice9.DrawPrimitiveUP D3DPT_LINELIST, 1, Vector2D(0), LenB(Vector2D(0))
End Sub

Public Sub CWDrawLine2(ByVal OX As Single, ByVal OY As Single, ByVal DX As Single, ByVal DY As Single, ByVal CColor As CWColorConstants, Optional ByVal Antialias As Boolean, Optional ByVal Pattern As CWLinePattern = CWLP_Solid, Optional ByVal Width As Single = 1!, Optional ByVal PatternScale As Single = 1!)
    Dim Vector2D(0 To 1) As D3DXVECTOR2
    With Vector2D(0)
        TransformationEx .X, .Y, OX, OY, WorldTransform
        .X = .X - 0.5!: .Y = .Y - 0.5!
    End With
    With Vector2D(1)
        TransformationEx .X, .Y, DX, DY, WorldTransform
        .X = .X - 0.5!: .Y = .Y - 0.5!
    End With
    CWPaintPicFlush ' 要先提交精灵的绘制，再画图形（否则层级会有问题）
    With CWLine
        .SetAntialias Antialias
        .SetPattern Pattern
        .SetWidth Width
        .SetPatternScale PatternScale
        .Begin
        .Draw Vector2D(0), 2, CColor
        .End
    End With
End Sub

Public Sub CWDrawLine2Ex(PointList() As D3DXVECTOR2, ByVal CColor As CWColorConstants, Optional ByVal Antialias As Boolean, Optional ByVal Pattern As CWLinePattern = CWLP_Solid, Optional ByVal Width As Single = 1!, Optional ByVal PatternScale As Single = 1!)
    Dim I&, Vector2D() As D3DXVECTOR2
    Vector2D = PointList
    For I = LBound(Vector2D) To UBound(Vector2D)
      With Vector2D(I)
        Transformation .X, .Y, WorldTransform
        .X = .X - 0.5!: .Y = .Y - 0.5!
      End With
    Next
    CWPaintPicFlush ' 要先提交精灵的绘制，再画图形（否则层级会有问题）
    With CWLine
        .SetAntialias Antialias
        .SetPattern Pattern
        .SetWidth Width
        .SetPatternScale PatternScale
        .Begin
        .Draw Vector2D(LBound(Vector2D)), UBound(Vector2D) - LBound(Vector2D) + 1, CColor
        .End
    End With
End Sub

Public Sub CWDrawHTriangle(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal X3 As Single, ByVal Y3 As Single, ByVal Color As CWColorConstants)
    Dim Vector2D(0 To 3) As D2DVector
    With Vector2D(0)
        TransformationEx .X, .Y, X1, Y1, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = 1!
        .Color = Color
    End With
    With Vector2D(1)
        TransformationEx .X, .Y, X2, Y2, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = 1!
        .Color = Color
    End With
    With Vector2D(2)
        TransformationEx .X, .Y, X3, Y3, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = 1!
        .Color = Color
    End With
    Vector2D(3) = Vector2D(0)

    CWPaintPicFlush ' 要先提交精灵的绘制，再画图形（否则层级会有问题）
    CWD3DDevice9.SetTexture 0, Nothing
    CWD3DDevice9.SetFVF CWP_FVFConst
    CWD3DDevice9.DrawPrimitiveUP D3DPT_LINESTRIP, 3, Vector2D(0), LenB(Vector2D(0))
End Sub

Public Sub CWDrawSTriangle(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal X3 As Single, ByVal Y3 As Single, ByVal Color As CWColorConstants)
    Call CWDrawSTriangleEx(X1, Y1, X2, Y2, X3, Y3, 1!, 1!, 1!, Color, Color, Color)
End Sub

Public Sub CWDrawSTriangleEx(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal X3 As Single, ByVal Y3 As Single, ByVal RHW1 As Single, ByVal RHW2 As Single, ByVal RHW3 As Single, ByVal Color1 As CWColorConstants, ByVal Color2 As CWColorConstants, ByVal Color3 As CWColorConstants)
    Dim Vector2D(0 To 2) As D2DVector
    With Vector2D(0)
        TransformationEx .X, .Y, X1, Y1, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = RHW1
        .Color = Color1
    End With
    With Vector2D(1)
        TransformationEx .X, .Y, X2, Y2, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = RHW2
        .Color = Color2
    End With
    With Vector2D(2)
        TransformationEx .X, .Y, X3, Y3, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = RHW3
        .Color = Color3
    End With

    CWPaintPicFlush ' 要先提交精灵的绘制，再画图形（否则层级会有问题）
    CWD3DDevice9.SetTexture 0, Nothing
    CWD3DDevice9.SetFVF CWP_FVFConst
    CWD3DDevice9.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 1, Vector2D(0), LenB(Vector2D(0))
End Sub

Public Sub CWDrawHRect(ByVal SX As Single, ByVal SY As Single, ByVal SWidth As Single, ByVal SHeight As Single, ByVal CColor As CWColorConstants)
    Call CWDrawHPlg(SX, SY, SX + SWidth, SY, SX, SY + SHeight, CColor)
End Sub

Public Sub CWDrawSRect(ByVal SX As Single, ByVal SY As Single, ByVal SWidth As Single, ByVal SHeight As Single, ByVal CColor As CWColorConstants)
    Call CWDrawSPlgEx(SX, SY, SX + SWidth, SY, SX, SY + SHeight, 1, 1, 1, CColor, CColor, CColor)
End Sub

Public Sub CWDrawSRectXGC(ByVal SX As Single, ByVal SY As Single, ByVal SWidth As Single, ByVal SHeight As Single, ByVal LRHW As Single, ByVal RRHW As Single, ByVal LColor As CWColorConstants, ByVal RColor As CWColorConstants)
    Call CWDrawSPlgEx(SX, SY, SX + SWidth, SY, SX, SY + SHeight, LRHW, RRHW, LRHW, LColor, RColor, LColor)
End Sub

Public Sub CWDrawSRectYGC(ByVal SX As Single, ByVal SY As Single, ByVal SWidth As Single, ByVal SHeight As Single, ByVal TRHW As Single, ByVal BRHW As Single, ByVal TColor As CWColorConstants, ByVal BColor As CWColorConstants)
    Call CWDrawSPlgEx(SX, SY, SX + SWidth, SY, SX, SY + SHeight, TRHW, TRHW, BRHW, TColor, TColor, BColor)
End Sub

Public Sub CWDrawHPlg(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal X3 As Single, ByVal Y3 As Single, ByVal Color As CWColorConstants)
    Dim Vector2D(0 To 4) As D2DVector
    With Vector2D(1)
        TransformationEx .X, .Y, X1, Y1, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = 1!
        .Color = Color
    End With
    With Vector2D(2)
        TransformationEx .X, .Y, X2, Y2, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = 1!
        .Color = Color
    End With
    With Vector2D(0)
        TransformationEx .X, .Y, X3, Y3, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = 1!
        .Color = Color
    End With
    With Vector2D(3)
        .X = Vector2D(0).X + Vector2D(2).X - Vector2D(1).X
        .Y = Vector2D(0).Y + Vector2D(2).Y - Vector2D(1).Y
        .Rhw = 1!
        .Color = Color
    End With
    Vector2D(4) = Vector2D(0)

    CWPaintPicFlush ' 要先提交精灵的绘制，再画图形（否则层级会有问题）
    CWD3DDevice9.SetTexture 0, Nothing
    CWD3DDevice9.SetFVF CWP_FVFConst
    CWD3DDevice9.DrawPrimitiveUP D3DPT_LINESTRIP, 4, Vector2D(0), LenB(Vector2D(0))
End Sub

Public Sub CWDrawSPlg(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal X3 As Single, ByVal Y3 As Single, ByVal Color As CWColorConstants)
    Call CWDrawSPlgEx(X1, Y1, X2, Y2, X3, Y3, 1!, 1!, 1!, Color, Color, Color)
End Sub

Public Sub CWDrawSPlgEx(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal X3 As Single, ByVal Y3 As Single, ByVal RHW1 As Single, ByVal RHW2 As Single, ByVal RHW3 As Single, ByVal Color1 As CWColorConstants, ByVal Color2 As CWColorConstants, ByVal Color3 As CWColorConstants)
    Dim Vector2D(0 To 3) As D2DVector
    With Vector2D(1)
        TransformationEx .X, .Y, X1, Y1, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = RHW1
        .Color = Color1
    End With
    With Vector2D(2)
        TransformationEx .X, .Y, X2, Y2, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = RHW2
        .Color = Color2
    End With
    With Vector2D(0)
        TransformationEx .X, .Y, X3, Y3, WorldTransform
        .X = .X - 0.5!
        .Y = .Y - 0.5!
        .Rhw = RHW3
        .Color = Color3
    End With
    With Vector2D(3)
        .X = Vector2D(0).X + Vector2D(2).X - Vector2D(1).X
        .Y = Vector2D(0).Y + Vector2D(2).Y - Vector2D(1).Y
        .Rhw = RHW2 + RHW3 - RHW1
        If Color1 = Color2 Then
            .Color = Color3
        ElseIf Color1 = Color3 Then
            .Color = Color2
        Else
            Dim T&, c1 As CWColor, c2 As CWColor, c3 As CWColor, c4 As CWColor
            c1 = CWSplitColor(Color1)
            c2 = CWSplitColor(Color2)
            c3 = CWSplitColor(Color3)
            T = CLng(c2.Alpha) + CLng(c3.Alpha) - CLng(c1.Alpha)
            If T > 255 Then c4.Alpha = 255 Else If T < 0 Then c4.Alpha = 0 Else c4.Alpha = T
            T = CLng(c2.Red) + CLng(c3.Red) - CLng(c1.Red)
            If T > 255 Then c4.Red = 255 Else If T < 0 Then c4.Red = 0 Else c4.Red = T
            T = CLng(c2.Green) + CLng(c3.Green) - CLng(c1.Green)
            If T > 255 Then c4.Green = 255 Else If T < 0 Then c4.Green = 0 Else c4.Green = T
            T = CLng(c2.Blue) + CLng(c3.Blue) - CLng(c1.Blue)
            If T > 255 Then c4.Blue = 255 Else If T < 0 Then c4.Blue = 0 Else c4.Blue = T
            GetMem4 VarPtr(c4), .Color
        End If
    End With

    CWPaintPicFlush ' 要先提交精灵的绘制，再画图形（否则层级会有问题）
    CWD3DDevice9.SetTexture 0, Nothing
    CWD3DDevice9.SetFVF CWP_FVFConst
    CWD3DDevice9.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 2, Vector2D(0), LenB(Vector2D(0))
End Sub

Public Sub CWDrawSRectXCGC(ByVal SX As Single, ByVal SY As Single, ByVal SWidth As Single, ByVal SHeight As Single, ByVal ORHW As Single, ByVal CRHW As Single, ByVal OColor As CWColorConstants, ByVal CColor As CWColorConstants)
    Call CWDrawSRectCGCEx(SX, SY, SWidth, SHeight, ORHW, CRHW, ORHW, CRHW, OColor, CColor, OColor, CColor)
End Sub

Public Sub CWDrawSRectYCGC(ByVal SX As Single, ByVal SY As Single, ByVal SWidth As Single, ByVal SHeight As Single, ByVal ORHW As Single, ByVal CRHW As Single, ByVal OColor As CWColorConstants, ByVal CColor As CWColorConstants)
    Call CWDrawSRectCGCEx(SX, SY, SWidth, SHeight, ORHW, CRHW, CRHW, ORHW, OColor, CColor, CColor, OColor)
End Sub

Public Sub CWDrawSRectCGC(ByVal SX As Single, ByVal SY As Single, ByVal SWidth As Single, ByVal SHeight As Single, ByVal ORHW As Single, ByVal CRHW As Single, ByVal OColor As CWColorConstants, ByVal CColor As CWColorConstants)
    Dim HAW As Single, HAH As Single, HARHW As Single
    HAW = 0.5! * SWidth
    HAH = 0.5! * SHeight
    HARHW = Abs(CRHW - ORHW) / Abs2(HAW, HAH)       '+ CRHW
    Call CWDrawSRectCGCEx(SX, SY, SWidth, SHeight, ORHW, CRHW, HAW * HARHW, HAH * HARHW, OColor, CColor, CColor, CColor)
End Sub

Public Sub CWDrawSRectCGCEx(ByVal SX As Single, ByVal SY As Single, ByVal SWidth As Single, ByVal SHeight As Single, ByVal ORHW As Single, ByVal CRHW As Single, ByVal XRHW As Single, ByVal YRHW As Single, ByVal OColor As CWColorConstants, ByVal CColor As CWColorConstants, ByVal XColor As CWColorConstants, ByVal YColor As CWColorConstants)
    Dim I&, Vector2D(0 To 9) As D2DVector, HAW As Single, HAH As Single
    HAW = 0.5 * SWidth
    HAH = 0.5 * SHeight
    Vector2D(0).X = SX + HAW
    Vector2D(0).Y = SY + HAH
    Vector2D(0).Color = OColor
    Vector2D(0).Rhw = ORHW
    Vector2D(1).X = SX
    Vector2D(1).Y = SY
    Vector2D(1).Color = CColor
    Vector2D(1).Rhw = CRHW
    Vector2D(2).X = SX + HAW
    Vector2D(2).Y = SY
    Vector2D(2).Color = XColor
    Vector2D(2).Rhw = XRHW
    Vector2D(3).X = SX + SWidth
    Vector2D(3).Y = SY
    Vector2D(3).Color = CColor
    Vector2D(3).Rhw = CRHW
    Vector2D(4).X = SX + SWidth
    Vector2D(4).Y = SY + HAH
    Vector2D(4).Color = YColor
    Vector2D(4).Rhw = YRHW
    Vector2D(5).X = SX + SWidth
    Vector2D(5).Y = SY + SHeight
    Vector2D(5).Color = CColor
    Vector2D(5).Rhw = CRHW
    Vector2D(6).X = SX + HAW
    Vector2D(6).Y = SY + SHeight
    Vector2D(6).Color = XColor
    Vector2D(6).Rhw = XRHW
    Vector2D(7).X = SX
    Vector2D(7).Y = SY + SHeight
    Vector2D(7).Color = CColor
    Vector2D(7).Rhw = CRHW
    Vector2D(8).X = SX
    Vector2D(8).Y = SY + HAH
    Vector2D(8).Color = YColor
    Vector2D(8).Rhw = YRHW
    For I = 0 To 8
      With Vector2D(I)
        Transformation .X, .Y, WorldTransform
        .X = .X - 0.5!: .Y = .Y - 0.5!
      End With
    Next
    Vector2D(9) = Vector2D(1)
    CWPaintPicFlush ' 要先提交精灵的绘制，再画图形（否则层级会有问题）
    CWD3DDevice9.SetTexture 0, Nothing
    CWD3DDevice9.SetFVF CWP_FVFConst
    CWD3DDevice9.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 8, Vector2D(0), LenB(Vector2D(0))
End Sub

Public Sub CWDrawSCircle(ByVal CX As Single, ByVal CY As Single, ByVal RR As Single, ByVal CColor As CWColorConstants)
    Call CWDrawSEllipseEx(CX, CY, RR, RR, 1, 1, CColor, CColor)
End Sub

Public Sub CWDrawSCircleEx(ByVal CX As Single, ByVal CY As Single, ByVal RR As Single, ByVal ORHW As Single, ByVal CRHW As Single, ByVal OColor As CWColorConstants, ByVal CColor As CWColorConstants)
    Call CWDrawSEllipseEx(CX, CY, RR, RR, ORHW, CRHW, OColor, CColor)
End Sub

Public Sub CWDrawHCircle(ByVal CX As Single, ByVal CY As Single, ByVal RR As Single, ByVal CColor As CWColorConstants)
    Call CWDrawHEllipse(CX, CY, RR, RR, CColor)
End Sub

Public Sub CWDrawSEllipse(ByVal CX As Single, ByVal CY As Single, ByVal RX As Single, ByVal RY As Single, ByVal CColor As CWColorConstants, Optional ByVal RollAngle As Single)
    Call CWDrawSEllipseEx(CX, CY, RX, RY, 1, 1, CColor, CColor, RollAngle)
End Sub

Public Sub CWDrawSEllipseEx(ByVal CX As Single, ByVal CY As Single, ByVal RX As Single, ByVal RY As Single, ByVal ORHW As Single, ByVal CRHW As Single, ByVal OColor As CWColorConstants, ByVal CColor As CWColorConstants, Optional ByVal RollAngle As Single)
    Dim Vector2D() As D2DVector, I&, C&, R!
    Dim X!, Y!, S1!, S2!: S1 = Cos(RollAngle): S2 = Sin(RollAngle)
    With WorldTransform
        X = RX * .m11 + RY * .m21
        Y = RX * .m12 + RY * .m22
        If X <= 0 Or Y <= 0 Then Exit Sub
        C = Abs2(X, Y)
        If C < 4& Then C = 4&
    End With
    ReDim Vector2D(0 To C + 1)
    With Vector2D(0)
        TransformationEx X, Y, CX, CY, WorldTransform
        .X = X - 0.5!
        .Y = Y - 0.5!
        .Rhw = ORHW
        .Color = OColor
    End With
    For I = 0 To C
      With Vector2D(I + 1)
        R = 2 * Pi * I / C
        X = RX * Cos(R)
        Y = RY * Sin(R)
        .X = X * S1 - Y * S2 + CX
        .Y = X * S2 + Y * S1 + CY
        TransformationEx X, Y, .X, .Y, WorldTransform
        .X = X - 0.5!
        .Y = Y - 0.5!
        .Rhw = CRHW
        .Color = CColor
      End With
    Next
    CWPaintPicFlush ' 要先提交精灵的绘制，再画图形（否则层级会有问题）
    CWD3DDevice9.SetTexture 0, Nothing
    CWD3DDevice9.SetFVF CWP_FVFConst
    CWD3DDevice9.DrawPrimitiveUP D3DPT_TRIANGLEFAN, C, Vector2D(0), LenB(Vector2D(0))
End Sub

Public Sub CWDrawHEllipse(ByVal CX As Single, ByVal CY As Single, ByVal RX As Single, ByVal RY As Single, ByVal CColor As CWColorConstants, Optional ByVal RollAngle As Single)
    Dim Vector2D() As D2DVector, I&, C&, R!
    Dim X!, Y!, S1!, S2!: S1 = Cos(RollAngle): S2 = Sin(RollAngle)
    With WorldTransform
        X = RX * .m11 + RY * .m21
        Y = RX * .m12 + RY * .m22
        If X <= 0 Or Y <= 0 Then Exit Sub
        C = Abs2(X, Y)
        If C < 4& Then C = 4&
    End With
    ReDim Vector2D(0 To C)
    For I = 0 To C
      With Vector2D(I)
        R = 2 * Pi * I / C
        X = RX * Cos(R)
        Y = RY * Sin(R)
        .X = X * S1 - Y * S2 + CX
        .Y = X * S2 + Y * S1 + CY
        TransformationEx X, Y, .X, .Y, WorldTransform
        .X = X - 0.5!
        .Y = Y - 0.5!
        .Rhw = 1
        .Color = CColor
      End With
    Next
    CWPaintPicFlush ' 要先提交精灵的绘制，再画图形（否则层级会有问题）
    CWD3DDevice9.SetTexture 0, Nothing
    CWD3DDevice9.SetFVF CWP_FVFConst
    CWD3DDevice9.DrawPrimitiveUP D3DPT_LINESTRIP, C, Vector2D(0), LenB(Vector2D(0))
End Sub

Public Sub CWLoadFont(CFont As CWFont, ByVal FName As String, ByVal FSize As Long, Optional ByVal FBold As CWFBStyle = CWF_Normal, Optional ByVal FItalic As Boolean)
    CWFontNum = CWFontNum + 1
    CFont.SNum = CWFontNum
    ReDim Preserve CWFontList(CWFontNum)
    D3DXCreateFontW CWD3DDevice9, FSize, 0, FBold, 0, FItalic, 1, 0, 4, 0, FName, CWFontList(CWFontNum)
End Sub

Public Sub CWCalcrFont(CFont As CWFont, ByVal Text As String, ByRef OutWidth As Long, ByRef OutHeight As Long, Optional ByVal SingleLine As Boolean)
    Dim rc As D3DRECT, txtl As Long
    txtl = Len(Text)
    
    If txtl <> 0 Then
        Dim fmt&: fmt = DT_CALCRECT Or DT_NOCLIP
        If SingleLine Then fmt = fmt Or DT_SINGLELINE
        CWFontList(CFont.SNum).DrawTextW Nothing, ByVal Text, txtl, rc, fmt, CWColorNone
        OutWidth = rc.X2 - rc.X1
        OutHeight = rc.Y2 - rc.Y1
    ElseIf SingleLine Then
        CWFontList(CFont.SNum).DrawTextW Nothing, vbNullChar, 1, rc, DT_CALCRECT Or DT_SINGLELINE, CWColorNone
        OutWidth = 0
        OutHeight = rc.Y2 - rc.Y1
    Else
        OutWidth = 0
        OutHeight = 0
    End If
End Sub

Public Sub CWPrintFont(CFont As CWFont, ByVal Text As String, ByVal PrintX As Long, ByVal PrintY As Long, ByVal FBOXWidth As Long, ByVal FBOXHeight As Long, ByVal CColor As CWColorConstants, Optional ByVal FAlign As CWFAlign)
    Dim rc As D3DRECT, txtl As Long, TexMatrix As D3DMATRIX
    txtl = Len(Text)
    If txtl = 0 Then Exit Sub
    rc.X1 = PrintX
    rc.Y1 = PrintY
    rc.X2 = PrintX + FBOXWidth
    rc.Y2 = PrintY + FBOXHeight
    If CWSpState = Ended Then
        CWFontList(CFont.SNum).DrawTextW Nothing, ByVal Text, txtl, rc, FAlign, CColor
        Exit Sub
    End If
    With TexMatrix
        .m11 = WorldTransform.m11
        .m12 = WorldTransform.m12
        .m21 = WorldTransform.m21
        .m22 = WorldTransform.m22
        .m33 = 1!
        .m44 = 1!
        .m41 = WorldTransform.mdx
        .m42 = WorldTransform.mdy
    End With
    CWSprite.SetTransform TexMatrix
    If CWFontList(CFont.SNum).DrawTextW(CWSprite, ByVal Text, txtl, rc, FAlign, CColor) > 0 Then
        CWSpState = Drawed
    End If
End Sub

Public Sub CWPrintFontTop(CFont As CWFont, ByVal Text As String, ByVal PrintX As Long, ByVal PrintY As Long, ByVal FBOXWidth As Long, ByVal FBOXHeight As Long, ByVal CColor As CWColorConstants, Optional ByVal FAlign As CWFAlign)
    Dim rc As D3DRECT, txtl As Long, TexMatrix As D3DMATRIX
    txtl = Len(Text)
    If txtl = 0 Then Exit Sub
    rc.X1 = PrintX
    rc.Y1 = PrintY
    rc.X2 = PrintX + FBOXWidth
    rc.Y2 = PrintY + FBOXHeight
    If CWSpState = Ended Then
        CWFontList(CFont.SNum).DrawTextW Nothing, ByVal Text, txtl, rc, FAlign, CColor
        Exit Sub
    End If
    With TexMatrix
        .m11 = WorldTransform.m11
        .m12 = WorldTransform.m12
        .m21 = WorldTransform.m21
        .m22 = WorldTransform.m22
        .m33 = 1!
        .m44 = 1!
        .m41 = WorldTransform.mdx
        .m42 = WorldTransform.mdy
    End With
    CWSpriteSP.SetTransform TexMatrix
    CWFontList(CFont.SNum).DrawTextW CWSpriteSP, ByVal Text, txtl, rc, FAlign, CColor
End Sub

Public Sub CWLoadMusic(Music As CWMusic, ByVal MPath As String, Optional Effect As CWMusicEffect)
    Dim PathS As String, dshow As IGraphBuilder
    Dim fx3d As DSoundRender
    PathS = GetShortName(MPath)
    Set dshow = New FilterGraph ' NoThread
    If Effect And CWME_3DFX Then
        Set fx3d = New DSoundRender
        dshow.AddFilter fx3d
    End If
    dshow.RenderFile PathS
    With Music
        If .ID = 0 Then
            CWMusicNum = CWMusicNum + 1
            .ID = CWMusicNum
            ReDim Preserve CWMusicList(1 To CWMusicNum)
        End If
    End With
    With CWMusicList(Music.ID)
        Set .mc = dshow
        Set .mp = dshow
        Set .ba = dshow
        Set .vw = dshow
        Set .evt = dshow
        Set .fx3d = fx3d
On Error Resume Next
        With .vw
            .AutoShow = False
            .Owner = CWHwnd
        End With
    End With
End Sub

Public Sub CWPlayMusic(Music As CWMusic, ByVal MPState As CWMPModel)
  With CWMusicList(Music.ID)
    Select Case MPState
    Case CWM_Once
        .IsLoop = False
    Case CWM_Repeat
        .IsLoop = True
        If .mp.CurrentPosition > .mp.Duration - 0.000001 Then
          .mp.CurrentPosition = 0
        End If
    Case CWM_Restart
        .IsLoop = False
        .mp.CurrentPosition = 0
    End Select
    .mc.Run
  End With
End Sub

Public Sub CWPauseMusic(Music As CWMusic)
    With CWMusicList(Music.ID)
        .mc.Pause
    End With
End Sub

Public Sub CWStopMusic(Music As CWMusic)
    With CWMusicList(Music.ID)
        .mc.Stop
        .mp.CurrentPosition = 0
        .mc.StopWhenReady
    End With
End Sub

Public Sub CWDelMusic(Music As CWMusic)
    With CWMusicList(Music.ID)
        If Not (.mc Is Nothing) Then
            .mc.Stop
            Set .mc = Nothing
        End If
        Set .mp = Nothing
        Set .ba = Nothing
        Set .vw = Nothing
    End With
End Sub

' 设置音量
Public Sub CWSetMusicVol(Music As CWMusic, ByVal mVolume As Single)
    With CWMusicList(Music.ID)
        .ba.Volume = VolumeToDecibel(mVolume)
    End With
End Sub

' 设置声道平衡
Public Sub CWSetMusicPan(Music As CWMusic, ByVal mBalance As Single)
    With CWMusicList(Music.ID)
      If mBalance > 0 Then
        .ba.Balance = -VolumeToDecibel(1000 - mBalance)
      Else
        .ba.Balance = VolumeToDecibel(1000 + mBalance)
      End If
    End With
End Sub

' 设置速度（音高）
Public Sub CWSetMusicRate(Music As CWMusic, ByVal mPitch As Single)
    With CWMusicList(Music.ID)
        .mp.Rate = mPitch
    End With
End Sub

' 获取播放时长（整数部分为秒数，小数部分用于微调）
Public Property Get CWMusicDuration(Music As CWMusic) As Double
    With CWMusicList(Music.ID)
        CWMusicDuration = .mp.Duration
    End With
End Property

' 获取播放进度
Public Property Get CWMusicPosition(Music As CWMusic) As Double
    With CWMusicList(Music.ID)
        CWMusicPosition = .mp.CurrentPosition
    End With
End Property

' 设置播放进度
Public Property Let CWMusicPosition(Music As CWMusic, ByVal mPos As Double)
    With CWMusicList(Music.ID)
        .mp.CurrentPosition = mPos
    End With
End Property

' 是否支持3D空间音效
Public Property Get CWMusicIs3D(Music As CWMusic) As Boolean
    With CWMusicList(Music.ID)
        CWMusicIs3D = Not (.fx3d Is Nothing)
    End With
End Property

' 设置音源的空间位置（对于2D游戏来说，忽略Y轴或Z轴的其中一个值就可以了）
' X 为负数表示左，正数表示右。
' Y 为负数表示下，正数表示上。
' Z 为负数表示后，正数表示前。
' 像渣渣辉那样的前后左右运动的2D游戏就建议忽略Z轴
' 像冒险岛那样的上下左右运动的2D游戏就建议忽略Y轴
Public Sub CWMusicSet3DPosition(Music As CWMusic, ByVal X As Single, Optional ByVal Y As Single, Optional ByVal Z As Single)
    With CWMusicList(Music.ID)
        .fx3d.SetPosition X, Y, Z
    End With
End Sub

' PicMode 为Ture表示PS的线性减淡模式（用于贴图作为光源）
' PicMode 为False表示原来的光照模式（用于填充图形作为光源）
Public Sub LightEFOpen(Optional ByVal PicMode As Boolean)
    CWPaintPicFlush
    If PicMode Then
        CWD3DDevice9.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        CWD3DDevice9.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    Else
        CWD3DDevice9.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTCOLOR
        CWD3DDevice9.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
    End If
End Sub

Public Sub LightEFClose()
    CWPaintPicFlush
    CWD3DDevice9.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    CWD3DDevice9.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub

Public Sub CWClipperOpen(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
    CWPaintPicFlush
    Dim rc As D3DRECT
    With rc
        .X1 = X: .X2 = X + Width
        .Y1 = Y: .Y2 = Y + Height
    End With
    With CWD3DDevice9
        .SetRenderState D3DRS_SCISSORTESTENABLE, True
        .SetScissorRect rc
    End With
End Sub

Public Sub CWClipperClose()
    CWPaintPicFlush
    CWD3DDevice9.SetRenderState D3DRS_SCISSORTESTENABLE, False
End Sub

Public Sub CWCheckKey(ByVal CWKeyAscii As Integer, CWKeySpecies As CWKeyState)

    Select Case GetAsyncKeyState(CWKeyAscii) < 0
        
        Case CWKU
            If CWKeySpecies.PUP Then CWKeySpecies.PUPMoment = False
            
            If CWKeySpecies.PDown Then CWKeySpecies.PUPMoment = True
            
        CWKeySpecies.PDownMoment = False
        CWKeySpecies.PUP = True
        CWKeySpecies.PDown = False
        
        Case CWKD
            If CWKeySpecies.PDown Then CWKeySpecies.PDownMoment = False
            
            If CWKeySpecies.PUP Then CWKeySpecies.PDownMoment = True
        
        CWKeySpecies.PUPMoment = False
        CWKeySpecies.PDown = True
        CWKeySpecies.PUP = False
        
    End Select
End Sub

Public Sub CWCheckKeySP(ByVal CWKeyAscii As Integer, CWKeySpecies As CWKeyStateSP)

    Select Case GetAsyncKeyState(CWKeyAscii) < 0
        
        Case CWKU
            If CWKeySpecies.PUP Then CWKeySpecies.PUPMoment = False
            
            If CWKeySpecies.PDown Then CWKeySpecies.PUPMoment = True
            
        CWKeySpecies.PDownMoment = False
        CWKeySpecies.PUP = True
        CWKeySpecies.PDown = False
        
        Case CWKD
            If CWKeySpecies.PDown Then CWKeySpecies.PDownMoment = False
            
            If CWKeySpecies.PUP Then CWKeySpecies.PDownMoment = True
        
        CWKeySpecies.PUPMoment = False
        CWKeySpecies.PDown = True
        CWKeySpecies.PUP = False
        
    End Select
End Sub

Public Sub CWCheckKeyMul(ByVal CWKeyAscii1 As Integer, ByVal CWKeyAscii2 As Integer, CWKeySpecies As CWKeyState)

    Select Case (GetAsyncKeyState(CWKeyAscii1) < 0) Or (GetAsyncKeyState(CWKeyAscii2) < 0)
        
        Case CWKU
            If CWKeySpecies.PUP Then CWKeySpecies.PUPMoment = False
            
            If CWKeySpecies.PDown Then CWKeySpecies.PUPMoment = True
            
        CWKeySpecies.PDownMoment = False
        CWKeySpecies.PUP = True
        CWKeySpecies.PDown = False
        
        Case CWKD
            If CWKeySpecies.PDown Then CWKeySpecies.PDownMoment = False
            
            If CWKeySpecies.PUP Then CWKeySpecies.PDownMoment = True
        
        CWKeySpecies.PUPMoment = False
        CWKeySpecies.PDown = True
        CWKeySpecies.PUP = False
        
    End Select
End Sub

Public Sub CWKeyboardCheck()
    CWCheckKey vbKeyA, CWKeyboard.A
    CWCheckKey vbKeyB, CWKeyboard.B
    CWCheckKey vbKeyC, CWKeyboard.C
    CWCheckKey vbKeyD, CWKeyboard.D
    CWCheckKey vbKeyE, CWKeyboard.E
    CWCheckKey vbKeyF, CWKeyboard.F
    CWCheckKey vbKeyG, CWKeyboard.G
    CWCheckKey vbKeyH, CWKeyboard.H
    CWCheckKey vbKeyI, CWKeyboard.I
    CWCheckKey vbKeyJ, CWKeyboard.j
    CWCheckKey vbKeyK, CWKeyboard.K
    CWCheckKey vbKeyL, CWKeyboard.L
    CWCheckKey vbKeyM, CWKeyboard.M
    CWCheckKey vbKeyN, CWKeyboard.N
    CWCheckKey vbKeyO, CWKeyboard.O
    CWCheckKey vbKeyP, CWKeyboard.P
    CWCheckKey vbKeyQ, CWKeyboard.Q
    CWCheckKey vbKeyR, CWKeyboard.R
    CWCheckKey vbKeyS, CWKeyboard.S
    CWCheckKey vbKeyT, CWKeyboard.T
    CWCheckKey vbKeyU, CWKeyboard.U
    CWCheckKey vbKeyV, CWKeyboard.V
    CWCheckKey vbKeyW, CWKeyboard.W
    CWCheckKey vbKeyX, CWKeyboard.X
    CWCheckKey vbKeyY, CWKeyboard.Y
    CWCheckKey vbKeyZ, CWKeyboard.Z
    CWCheckKey vbKeyUp, CWKeyboard.UP
    CWCheckKey vbKeyDown, CWKeyboard.Down
    CWCheckKey vbKeyLeft, CWKeyboard.Left
    CWCheckKey vbKeyRight, CWKeyboard.Right
    CWCheckKey vbKeySpace, CWKeyboard.Space
    CWCheckKey 13, CWKeyboard.Enter
    'CWCheckKey vbKeyShift, CWKeyboard.Shift
    CWCheckKeyMul &HA0, &HA1, CWKeyboard.Shift
    'CWCheckKey vbKeyControl, CWKeyboard.Ctrl
    CWCheckKeyMul &HA2, &HA3, CWKeyboard.Ctrl
    'CWCheckKey vbKeyMenu, CWKeyboard.Alt        '需要修正
    CWCheckKeyMul &HA4, &HA5, CWKeyboard.Alt
    CWCheckKey vbKeyTab, CWKeyboard.Tab
    CWCheckKey vbKeyBack, CWKeyboard.BackSpace
    CWCheckKey vbKeyF1, CWKeyboard.F1
    CWCheckKey vbKeyF2, CWKeyboard.F2
    CWCheckKey vbKeyF3, CWKeyboard.F3
    CWCheckKey vbKeyF4, CWKeyboard.F4
    CWCheckKey vbKeyF5, CWKeyboard.F5
    CWCheckKey vbKeyF6, CWKeyboard.F6
    CWCheckKey vbKeyF7, CWKeyboard.F7
    CWCheckKey vbKeyF8, CWKeyboard.F8
    CWCheckKey vbKeyF9, CWKeyboard.F9
    CWCheckKey vbKeyF10, CWKeyboard.F10
    CWCheckKey vbKeyF11, CWKeyboard.F11
    CWCheckKey vbKeyF12, CWKeyboard.F12
    CWCheckKey vbKeyInsert, CWKeyboard.Insert
    CWCheckKey vbKeyDelete, CWKeyboard.Delete
    CWCheckKey vbKeyPageUp, CWKeyboard.PageUp
    CWCheckKey vbKeyPageDown, CWKeyboard.PageDown
    CWCheckKey vbKeyHome, CWKeyboard.Home
    CWCheckKey vbKeyEnd, CWKeyboard.End
    CWCheckKeyMul vbKey0, vbKeyNumpad0, CWKeyboard.Num0
    CWCheckKeyMul vbKey1, vbKeyNumpad1, CWKeyboard.Num1
    CWCheckKeyMul vbKey2, vbKeyNumpad2, CWKeyboard.Num2
    CWCheckKeyMul vbKey3, vbKeyNumpad3, CWKeyboard.Num3
    CWCheckKeyMul vbKey4, vbKeyNumpad4, CWKeyboard.Num4
    CWCheckKeyMul vbKey5, vbKeyNumpad5, CWKeyboard.Num5
    CWCheckKeyMul vbKey6, vbKeyNumpad6, CWKeyboard.Num6
    CWCheckKeyMul vbKey7, vbKeyNumpad7, CWKeyboard.Num7
    CWCheckKeyMul vbKey8, vbKeyNumpad8, CWKeyboard.Num8
    CWCheckKeyMul vbKey9, vbKeyNumpad9, CWKeyboard.Num9
    CWCheckKey vbKeyEscape, CWKeyboard.ESC
End Sub

Public Sub CWMouseCheck()
    Dim CWMXY As POINTAPI
    GetCursorPos CWMXY

    If CWDModelW = CW_FullScreen Then
        CWMouse.X = CWMXY.X
        CWMouse.Y = CWMXY.Y
        If IsActive Then
            CWCheckKey vbKeyLButton, CWMouse.LeftKey
            CWCheckKey vbKeyRButton, CWMouse.RightKey
            CWCheckKeySP vbKeyMButton, CWMouse.MidKey
            CWCheckKey 5, CWMouse.BackKey
            CWCheckKey 6, CWMouse.ForwardKey
        End If
    Else
        Dim rc As D3DRECT
        IsHitWnd = WindowFromPoint(CWMXY.X, CWMXY.Y) = CWHwnd
        GetClientRect CWHwnd, rc
        ScreenToClient CWHwnd, CWMXY
        If IsHitWnd Then IsHitWnd = PtInRect(rc, CWMXY.X, CWMXY.Y)
        CWMouse.X = (CWMXY.X - rc.X1) * CWDModelX / (rc.X2 - rc.X1)
        CWMouse.Y = (CWMXY.Y - rc.Y1) * CWDModelY / (rc.Y2 - rc.Y1)
        If IsActive And IsHitWnd Then
            CWCheckKey vbKeyLButton, CWMouse.LeftKey
            CWCheckKey vbKeyRButton, CWMouse.RightKey
            CWCheckKeySP vbKeyMButton, CWMouse.MidKey
            CWCheckKey 5, CWMouse.BackKey
            CWCheckKey 6, CWMouse.ForwardKey
        End If
    End If
End Sub

Public Property Get CWCheckJoyAxis(ByVal jixPos As Long, ByVal Flag As JOYFALGS) As Single
    If Flag Then
        jixPos = jixPos \ 256
        CWCheckJoyAxis = (jixPos - 127 + (jixPos >= &H80&)) / 127!
    End If
End Property

Public Sub CWCheckJoyButton(State As CWKeyState, ByVal IsDown As Boolean)
    If IsDown Then
        If State.PDown Then State.PDownMoment = False
        If State.PUP Then State.PDownMoment = True
        
        State.PUPMoment = False
        State.PDown = True
        State.PUP = False
    Else
        If State.PUP Then State.PUPMoment = False
        If State.PDown Then State.PUPMoment = True
            
        State.PDownMoment = False
        State.PUP = True
        State.PDown = False
    End If
End Sub

Public Sub CWJoystickCheck()
    Dim I&, j&, jix As JOYINFOEX
    jix.dwSize = LenB(jix)
    For I = LBound(CWJoystick) To UBound(CWJoystick)
        jix.dwFlags = JOY_RETURNALL
        With CWJoystick(I)
            .IsConnected = 0 = joyGetPosEx(I, jix)
            If .IsConnected Then
                ' 读取手柄成功
                .IsPov = jix.dwPOV >= 0
                .X = CWCheckJoyAxis(jix.dwXpos, jix.dwFlags And JOY_RETURNX)
                .Y = CWCheckJoyAxis(jix.dwYpos, jix.dwFlags And JOY_RETURNY)
                .Z = CWCheckJoyAxis(jix.dwZpos, jix.dwFlags And JOY_RETURNZ)
                .R = CWCheckJoyAxis(jix.dwRpos, jix.dwFlags And JOY_RETURNR)
                If Not .IsPov Then .Pov = 0 Else If jix.dwPOV > 0 Then .Pov = jix.dwPOV * 0.01! Else .Pov = 360!
                Dim jb&: jb = 1
                For j = LBound(.Btn) To UBound(.Btn)
                    CWCheckJoyButton .Btn(j), jix.dwButtons And jb
                    jb = jb * 2
                Next
            Else
                ' 读取手柄失败 (可能是未插入的原因)
                ZeroMemory CWJoystick(I), LenB(CWJoystick(I))
            End If
        End With
    Next
End Sub

Public Sub AltBUGRepair(ByVal KeyCode As Integer)
If KeyCode = 18 Then
mouse_event RightButtonUpDown, 0, 0, 0, 0
SendMessage CWHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Public Sub MediaBUGRepair()
    Dim Music As CWMusic
    For Music.ID = 1 To CWMusicNum
        CWDelMusic Music
    Next
End Sub

Public Sub CWMediaLoopRepair()
    On Error Resume Next
    Dim I&, evc&
    For I = 1 To CWMusicNum
      With CWMusicList(I)
        If .IsLoop And Not (.mp Is Nothing) Then
            .evt.WaitForCompletion 0, evc
            If 1 = evc Then .mp.CurrentPosition = 0
        End If
      End With
    Next
End Sub

Public Property Get CWColorARGB(ByVal A As Byte, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As CWColorConstants
    Dim bgra As CWColor
    With bgra
        .Alpha = A
        .Red = R
        .Green = G
        .Blue = B
    End With
    GetMem4 VarPtr(bgra), CWColorARGB
End Property

Public Property Get CWColorRGBA(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255) As CWColorConstants
    Dim bgra As CWColor
    With bgra
        .Alpha = A
        .Red = R
        .Green = G
        .Blue = B
    End With
    GetMem4 VarPtr(bgra), CWColorRGBA
End Property

Public Property Get GetShortName(ByVal sLongPath As String) As String
    Dim sShortPath As String, L As Long
    sShortPath = String$(259, vbNullChar)
    L = GetShortPathNameW(StrPtr(sLongPath), StrPtr(sShortPath), 260)
    GetShortName = Left$(sShortPath, L)
End Property

Public Sub CWGetFPS()
    If timeGetTime - CWLongTime > 1000 Then
    CWLongTime = timeGetTime
    CWFPS = CWFrameCount
    CWFrameCount = 0
    Else
    CWFrameCount = CWFrameCount + 1
    End If
End Sub

Public Sub CWSetFPS()
Dim S As Long
S = (CLng(1000 / 40) - (timeGetTime - CWTimeNow))
    If S > 0 Then Sleep S
    CWTimeNow = timeGetTime
End Sub

Public Sub CWPaintPicBegin()
    If CWSpState <> Ended Then Exit Sub
    CWSpriteSP.Begin (CWP_SpriteConst Or D3DXSPRITE_SORT_TEXTURE)
    CWSprite.Begin (CWP_SpriteConst)
    CWSpState = Begined
End Sub

Public Sub CWPaintPicEnd()
    If CWSpState = Ended Then Exit Sub
    CWSpState = Ended
    CWSprite.End
    CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, True
    CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, True
    CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_BORDER
    CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_BORDER
    CWSpriteSP.End
End Sub

Public Sub CWPaintPicFlush()
    If CWSpState <> Drawed Then Exit Sub
    CWSpState = Begined
    CWSprite.Flush
End Sub

Public Property Get CRad(ByVal Angle As Single) As Single
    CRad = Angle * Pi / 180!
End Property

Public Property Get Abs2(ByVal X As Single, ByVal Y As Single) As Single
    Abs2 = Sqr(X * X + Y * Y)
End Property

Public Property Get VolumeToDecibel(ByVal Volume As Single) As Long
    If Volume <= 0.01 Then
        VolumeToDecibel = -10000
    ElseIf Volume >= 1000# Then
        VolumeToDecibel = 0
    Else
        VolumeToDecibel = CLng(2000# * Log(Volume * 0.001) / Log(10#))
    End If
End Property

Public Sub Transformation(ByRef X!, ByRef Y!, Matrix As CWMatrix)
    Call TransformationEx(X, Y, X, Y, Matrix)
End Sub

Public Sub TransformationEx(ByRef OutX!, ByRef OutY!, ByVal InX!, ByVal InY!, Matrix As CWMatrix)
    With Matrix
        OutX = InX * .m11 + InY * .m21 + .mdx
        OutY = InX * .m12 + InY * .m22 + .mdy
    End With
End Sub

' X = Pt.X - Rect.X: Y = Pt.Y - Rect.Y
' Width = Rect.Width: Height = Rect.Height
Public Property Get CWPtInRect(ByVal X!, ByVal Y!, ByVal Width!, ByVal Height!) As Boolean
    CWPtInRect = X >= 0! And Y >= 0! And X <= Width And Y <= Height
End Property

' X = Rect1.X - Rect2.X: Y = Rect1.Y - Rect2.Y
' Width1 = Rect1.Width: Height1 = Rect1.Height
' Width2 = Rect2.Width: Height2 = Rect2.Height
Public Property Get CWRectInRect(ByVal X!, ByVal Y!, ByVal Width1!, ByVal Height1!, ByVal Width2!, ByVal Height2!) As Boolean
    CWRectInRect = 0! <= X + Width1 And 0! <= Y + Height1 And X <= Width2 And Y <= Height2
End Property

' X = Pt.X - Circle.X: Y = Pt.Y - Circle.Y: R = Circle.R
' X = Circle1.X - Circle2.X: Y = Circle1.Y - Circle2.Y: R = Circle1.R + Circle2.R
Public Property Get CWPtInCircle(ByVal X!, ByVal Y!, ByVal R!) As Boolean
    CWPtInCircle = X * X + Y * Y <= R * R
End Property
