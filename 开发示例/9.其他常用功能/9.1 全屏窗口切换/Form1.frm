VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ȫ���������л�"
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

Dim FontDemo As CWFont      '����CoolWind�����������

    CWLoadFont FontDemo, "SimSun", 32, CWF_Normal, False  '����Ҫ��ʾ������

    Do While CWGameRun = True         '������Ϸѭ��
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '����Ƿ������Ⱦ���豸�������Ҵ���δ��С��ʱ��Ⱦ��
    CWBeginScene    '׼���û��Ƴ���
         
         CWPrintFont FontDemo, "����ALT+ENTER���л�ȫ���ʹ���ģʽ", 0, 0, 800, 60, CWWhite, CWF_LeftAl

         If CWKeyboard.Enter.PDown And CWKeyboard.Enter.PDownMoment Then          '���س���̧��˲��
         CWWinFullScrSwitch         'ִ��ȫ��/����ģʽ�л�
         End If
         
'ע�⣺����ģʽ�������̼�������Ƚϸߣ��ӳ�<10ms���������ͼ�����Եͣ�ϵͳ�ر�æʱ�׳��ֻ��涶����
'����ϵͳCPUʹ���ʲ����߶����ֻ��涶����ͨ����Ѹ�ס�360�������������������DX9�Ļ���ˢ�£���ʱ����رռ��ɡ��������DX9��Ϸ���д������⡣����Ҳ�ڳ��Խ������
'    ȫ��ģʽ�������̼�������Ƚϵͣ��ӳ�<30ms���������ͼ���ܽϸߡ�����ʹ�ڵ�֡��ʱ��Ҳ�ܼ��õ����ƻ��涶����
'    ������Լ���Ϸ���ص�ѡ����Ϸ����ʱĬ�ϵ���ʾģʽ

    'С���գ������Ӿ���ʱ���ӳ�Ϊ20ms���ң������ְ�����Ӧ�ӳ�Ϊ100ms���ϣ���600APMΪ���༫�ޣ�
    
         
         If CWKeyboard.ESC.PDownMoment Then   '����ESC����˲���˳�����
         CWGameRun = False           '�ر����濪��
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
