VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ƶ����"
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

    CWLoadFont FontDemo, "SimSun", 32, CWF_Normal, False      '������ͨ����32������

    Dim MusicDemo1 As CWMusic, MusicDemo2 As CWMusic    '����CoolWind������Ƶ����
    
            CWLoadMusic MusicDemo1, App.Path & "\Music\" & "��Ϸ���⿪��ӡ.mp3"
            CWLoadMusic MusicDemo2, App.Path & "\Music\" & "��Ϸ�����ҵľ���.mp3"
            '��������(��Ƶ����,��Ƶ�ļ�·��)
            
            CWSetMusicVol MusicDemo1, 100
            CWSetMusicVol MusicDemo2, 200
             '��������(��Ƶ����,������С)
                '��������������СĬ����1000(���ֵ)
                
            '��������δ���ƴ�:��Ƶ����ʱ�䳤,������ʱ���ڻῨס.��ʽ�潫����ʹ�������ļ����������.
    
    Do While CWGameRun = True         '������Ϸѭ��
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '����Ƿ������Ⱦ���豸�������Ҵ���δ��С��ʱ��Ⱦ��
    CWBeginScene    '׼���û��Ƴ���
         
        CWPrintFont FontDemo, "P����������1��S��ֹͣ����1��M����������2��E��ֹͣ����2��B����ͣ1��2", 0, 100, 800, 100, CWYellow, CWF_CenterAl
        If CWKeyboard.P.PDownMoment Then
        CWPlayMusic MusicDemo1, CWM_Repeat     'P�����µ�˲��ѭ����������1
        '��������(��Ƶ,����ģʽ[ѭ��/һ��])
        End If
        
        If CWKeyboard.S.PDownMoment Then
        CWStopMusic MusicDemo1              'S�����µ�˲��ֹͣ����1
        'ֹͣ����(��Ƶ)
        End If
        
        If CWKeyboard.M.PDownMoment Then
        CWPlayMusic MusicDemo2, CWM_Once    'M�����µ�˲�䲥������2
        End If
        
        If CWKeyboard.E.PDownMoment Then
        CWStopMusic MusicDemo2             'E�����µ�˲��ֹͣ����2
        End If
        
        If CWKeyboard.B.PDownMoment Then
        CWPauseMusic MusicDemo1             'B�����µ�˲����ͣ����1��2
        CWPauseMusic MusicDemo2
        '��ͣ����(��Ƶ)
        End If
        
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
