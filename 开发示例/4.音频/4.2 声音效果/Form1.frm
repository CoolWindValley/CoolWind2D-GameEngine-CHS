VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��̬����"
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
   StartUpPosition =   2  '��Ļ����
   Begin ��̬����.TabPage TabPages 
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
   Begin ��̬����.TabPage TabPages 
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
   Begin ��̬����.TabPage TabPages 
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
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ч"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
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
Option Explicit     '����ʹ��ǰ��������
'����ע�⣺ʹ��ǰ��ȷ�� DX9VB.TLB����ѱ����ã�VBDX9BAS.basģ���ѱ����

Dim MusicDemo1 As CWMusic, MusicDemo2(0 To 7) As CWMusic, MusicDemo3 As CWMusic  '����CoolWind������Ƶ����
Dim TabPage As Long, CurEffect As Long

Private Sub Form_Load()

'����ע�⣺��Ϸ����У�
'ͨ��������� BorderStyle ����Ϊ��Fixed single����������ı䴰���С
'ͨ��������� MinButton ����Ϊ��True����������С��
'ͨ��������� MaxButton ����Ϊ��False������ֹ���
    
    With TabStrip1
        .Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
        For TabPage = TabPages.LBound To TabPages.UBound
            TabPages.Item(TabPage).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
        Next
        TabPages.Item(.SelectedItem.Index).Visible = True
        TabPage = 0
    End With
    
    CWLoadMusic MusicDemo1, App.Path & "\Music\bgm.mp3"    ' �������֣�һ�����һ��ѭ�����ž͹��ˣ�
    For CurEffect = LBound(MusicDemo2) To UBound(MusicDemo2)    ' ������Ч�������Ҫ�ص����ŵĻ�����Ҫ���ض�Σ�
        CWLoadMusic MusicDemo2(CurEffect), App.Path & "\Music\" & "attack.mp3"
    Next
    CWLoadMusic MusicDemo3, App.Path & "\Music\voice.mp3"
    
    '��������δ���ƴ�:��Ƶ����ʱ�䳤,������ʱ���ڻῨס.��ʽ�潫����ʹ�������ļ����������.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CWGameRun = False       '���屻�ر�ʱ�ر�����
    MediaBUGRepair          '�޸��������˳�ʱ�����ڲ��ŵ����ֲ�ֹͣ����ʱ���������IDE����Ч��
End Sub

Private Sub TabPages_ButtonClick(Index As Integer, ByVal Check As Boolean)
  With TabPages.Item(Index)
    If 1 = Index Then
        ' ���Ż���ͣ����
        If Check Then
            CWPlayMusic MusicDemo1, CWM_Repeat
        Else
            CWPauseMusic MusicDemo1
        End If
    ElseIf 2 = Index Then
        ' ������Ч
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
    CWMediaLoopRepair   ' �޸�ѭ�����Ź��ܣ����������������Ⱦ�Ͳ����ˣ�
End Sub
