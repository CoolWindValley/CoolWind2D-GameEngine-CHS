VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������Ч"
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
   StartUpPosition =   2  '��Ļ����
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     '����ʹ��ǰ��������
'����ע�⣺ʹ��ǰ��ȷ�� DX9VB.TLB����ѱ����ã�VBDX9BAS.basģ���ѱ����

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
AltBUGRepair KeyCode    '�޸�����ALT���»���ֹͣˢ�µ���ʱ�������
End Sub

Private Sub Form_Load()

'����ע�⣺��Ϸ����У�
'ͨ��������� BorderStyle ����Ϊ��Fixed single����������ı䴰���С
'ͨ��������� MinButton ����Ϊ��True����������С��
'ͨ��������� MaxButton ����Ϊ��False������ֹ���

  '��ʼ�����沢���������ʼ�����������ֱ��ʣ�������ǵ��Գ��õķֱ��ʱ��� 640,480 �� 800,600 �� 1024,768 �� 1366,768
CWVBDX9Initialization Me, 800, 600, CW_Windowed
  '��ʼ�����棨Ŀ�괰�壬����ֱ��ʣ�����ֱ��ʣ�����ģʽ/ȫ��ģʽ��

    Dim FontDemo As CWFont      '������ʾ�õ�����

    CWLoadFont FontDemo, "Microsoft YaHei", 32, CWF_Normal, False      '������ͨ����32������

    Dim MusicDemo1 As CWMusic, MusicDemo2 As CWMusic    '����CoolWind������Ƶ����
    
    '��������(��Ƶ��������Ƶ�ļ�·������Ч����)
    CWLoadMusic MusicDemo1, App.Path & "\Music\" & "test.mp3", CWME_3DFX
    CWLoadMusic MusicDemo2, App.Path & "\Music\" & "effect.mp3", CWME_3DFX
    
    CWPlayMusic MusicDemo1, CWM_Repeat
        
    '��������δ���ƴ�:��Ƶ����ʱ�䳤,������ʱ���ڻῨס.��ʽ�潫����ʹ�������ļ����������.
    
    Do While CWGameRun = True         '������Ϸѭ��
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '����Ƿ������Ⱦ���豸�������Ҵ���δ��С��ʱ��Ⱦ��
    CWBeginScene    '׼���û��Ƴ���
      CWPaintPicBegin
        CWPrintFont FontDemo, "���������Զ��������Χ����", 0, 0, 800, 60, CWFuchsia, CWF_CenterAl Or DT_VCENTER Or DT_NOCLIP
        CWPrintFont FontDemo, "������������ڵ��λ�ò�����Ч" & vbNewLine _
            & "��ס���������Ų��϶��ɸı���Ч�Ŀռ�����", 0, 500, 800, 100, CWViolet, CWF_CenterAl Or DT_VCENTER Or DT_NOCLIP
        
        ' ���������ͷ�����ģ��
        CWDrawHTriangle 400, 280, 380, 312, 420, 312, CWColorRGBA(0, 255, 255, 120)
        
        Const R! = 240!, PxpMi = 100
        Dim t1!: t1 = t1 + 0.02!
        Dim X1!: X1 = R * Cos(t1)
        Dim Y1!: Y1 = R * Sin(t1)
        
        ' �������Ʊ������ֵ�·��
        CWDrawHCircle 400, 300, R, CWYellow
        CWDrawLine2 400, 300, 400 + X1, 300 + Y1, CWColorARGB(100, 255, 255, 0), True, CWLP_Dot, 2
        ' ���»��Ʊ������ֵĿռ����꣨ע��ǰ����Z�ᣬ����Y�ᣩ
        CWMusicSet3DPosition MusicDemo1, X1 / PxpMi, , Y1 / PxpMi
        
        Dim t2!, a2!, X2!, Y2!
        If CWMouse.LeftKey.PDownMoment Then
            ' ���������ʱ��ʼ������Ч����ʼ������
            CWPlayMusic MusicDemo2, CWM_Restart
            t2 = 0: a2 = 200
        End If
        If CWMouse.LeftKey.PDown Then
            ' ��ס������ʱ�ƶ���Ч����Դλ��
            X2 = CWMouse.X: Y2 = CWMouse.Y
            CWMusicSet3DPosition MusicDemo2, (X2 - 400) / PxpMi, , (Y2 - 300) / PxpMi
        End If
        If a2 > 0 Then
            ' ���������Ƶ�������ɢͼ
            t2 = t2 + 3.14159265358979
            CWDrawSCircle X2, Y2, t2, CWColorARGB(CByte(a2), 20, 200, 255)
            a2 = a2 - 2.44645364561314
        End If
        
      CWPaintPicEnd
    CWPresentScene   '���ֻ��Ƶĳ���

    
'*******************************����Ϊ�̶�д������Ҫ���׸Ķ�***********************************
    Else                 '����������Ⱦ����ʱ
        CWResetDevice       '�޸��豸
    End If

    Loop

        CWVBDX9Destory     '����CoolWind����
    End '�˳�
'*******************************����Ϊ�̶�д������Ҫ���׸Ķ�***********************************


End Sub

Private Sub Form_Unload(Cancel As Integer)
CWGameRun = False       '���屻�ر�ʱ�ر�����
MediaBUGRepair          '�޸��������˳�ʱ�����ڲ��ŵ����ֲ�ֹͣ����ʱ���������IDE����Ч��
End Sub
