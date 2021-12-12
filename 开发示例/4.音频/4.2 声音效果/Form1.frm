VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "动态变声"
   ClientHeight    =   3000
   ClientLeft      =   3090
   ClientTop       =   4350
   ClientWidth     =   5070
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5070
   StartUpPosition =   2  '屏幕中心
   Begin 动态变声.TabPage TabPages 
      Height          =   2535
      Index           =   3
      Left            =   65
      TabIndex        =   3
      Top             =   315
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4471
      ButtonCheck     =   -1  'True
      PitchBase       =   3
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1920
      Top             =   1320
   End
   Begin 动态变声.TabPage TabPages 
      Height          =   2535
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   315
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4471
   End
   Begin 动态变声.TabPage TabPages 
      Height          =   2535
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   315
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4471
      ButtonCheck     =   -1  'True
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   5292
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "音乐"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "音效"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "人声"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     '变量使用前必须声明
'新手注意：使用前请确保 DX9VB.TLB组件已被引用，VBDX9BAS.bas模块已被添加

Dim MusicDemo1 As CWMusic, MusicDemo2(0 To 7) As CWMusic, MusicDemo3 As CWMusic  '定义CoolWind引擎音频变量
Dim TabPage As Long, CurEffect As Long

Private Sub Form_Load()

'新手注意：游戏编程中，
'通常将窗体的 BorderStyle 设置为“Fixed single”即不允许改变窗体大小
'通常将窗体的 MinButton 设置为“True”即允许最小化
'通常将窗体的 MaxButton 设置为“False”即禁止最大化
    
    With TabStrip1
        .Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
        For TabPage = TabPages.LBound To TabPages.UBound
            TabPages.Item(TabPage).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
        Next
        TabPages.Item(.SelectedItem.Index).Visible = True
        TabPage = 0
    End With
    
    CWLoadMusic MusicDemo1, App.Path & "\Music\bgm.mp3"    ' 加载音乐（一般加载一次循环播放就够了）
    For CurEffect = LBound(MusicDemo2) To UBound(MusicDemo2)    ' 加载音效（如果需要重叠播放的话，需要加载多次）
        CWLoadMusic MusicDemo2(CurEffect), App.Path & "\Music\" & "attack.mp3"
    Next
    CWLoadMusic MusicDemo3, App.Path & "\Music\voice.mp3"
    
    '本引擎尚未完善处:音频载入时间长,且载入时窗口会卡住.正式版将至少使编译后的文件载入更流畅.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CWGameRun = False       '窗体被关闭时关闭引擎
    MediaBUGRepair          '修复非正常退出时，正在播放的音乐不停止的临时解决函数（IDE下无效）
End Sub

Private Sub TabPages_ButtonClick(Index As Integer, ByVal Check As Boolean)
  With TabPages.Item(Index)
    If 1 = Index Then
        ' 播放或暂停音乐
        If Check Then
            CWPlayMusic MusicDemo1, CWM_Repeat
        Else
            CWPauseMusic MusicDemo1
        End If
    ElseIf 2 = Index Then
        ' 播放音效
        If CurEffect > UBound(MusicDemo2) Then
            CurEffect = LBound(MusicDemo2)
        End If
        CWSetMusicVol MusicDemo2(CurEffect), .Volume
        CWSetMusicPan MusicDemo2(CurEffect), .Balance
        CWSetMusicRate MusicDemo2(CurEffect), .Pitch
        CWPlayMusic MusicDemo2(CurEffect), CWM_Restart
        CurEffect = CurEffect + 1
    ElseIf 3 = Index Then
        If Check Then
            CWPlayMusic MusicDemo3, CWM_Repeat
        Else
            CWPauseMusic MusicDemo3
        End If
    End If
  End With
End Sub

Private Sub TabPages_SliderChange(Index As Integer, ByVal SndOpt As SoundOption, ByVal Value As Single)
    If 1 = Index Then
        Select Case SndOpt
        Case soVolume
            CWSetMusicVol MusicDemo1, Value
        Case soBalance
            CWSetMusicPan MusicDemo1, Value
        Case soPitch
            CWSetMusicRate MusicDemo1, Value
        End Select
    ElseIf 3 = Index Then
        Select Case SndOpt
        Case soVolume
            CWSetMusicVol MusicDemo3, Value
        Case soBalance
            CWSetMusicPan MusicDemo3, Value
        Case soPitch
            CWSetMusicRate MusicDemo3, Value
        End Select
    End If
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    TabPage = TabStrip1.SelectedItem.Index
End Sub

Private Sub TabStrip1_Click()
    If TabPage > 0 Then
        TabPages.Item(TabPage).Visible = False
        TabPage = 0
    End If
    TabPages.Item(TabStrip1.SelectedItem.Index).Visible = True
End Sub

Private Sub Timer1_Timer()
    CWMediaLoopRepair   ' 修复循环播放功能（如果启用了引擎渲染就不用了）
End Sub
