VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
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

Dim FontDemo1 As CWFont, FontDemo2 As CWFont, FontDemo3 As CWFont    '����CoolWind�����������
Const HelloWorld$ = "Hello World!"

    CWLoadFont FontDemo1, "SimSun", 32, CWF_Bold, False
    CWLoadFont FontDemo2, "@SimHei", 32, CWF_Normal, False
    CWLoadFont FontDemo3, "Microsoft YaHei", 64, CWF_Normal, False
    '��������(�������,��������,�����С,�Ƿ�Ϊ����,�Ƿ�Ϊб��)
        'ע��:������������ͱ������������и���Ϸ��ϵͳ�ϴ��ڵ�����.����Ĭ�ϰ�"����"����
        'С����:�ڴ����������ǰ����@��������������~!
    
    Do While CWGameRun = True         '������Ϸѭ��
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '����Ƿ������Ⱦ���豸�������Ҵ���δ��С��ʱ��Ⱦ��
    CWBeginScene    '׼���û��Ƴ���
        CWPaintPicBegin
        
        CWPrintFont FontDemo1, CWFPS, 0, 0, 800, 60, CWWhite, CWF_LeftAl    '��ʾ��ǰFPS
        CWPrintFont FontDemo1, "�ӳ������������ҡ���", 0, 100, 800, 60, CWCyan, CWF_LeftAl
        CWPrintFont FontDemo2, "      ��ǹ        ", 0, 100, 800, 60, CWRed, CWF_LeftAl
        '�������(����,Ŀ������,��ʾ��������,��ʾ���������,��ʾ�����,��ʾ���߶�,������ɫ,���뷽ʽ)
               'ע��:�����ܳ��ȳ�����ʾ����Ƚ�������.   �����ܳ��ȳ�����ʾ���߶Ƚ�����ü�,���������ֲ���ʾ.
        Dim cx&, cy&    '�����ı���ʾ��С���������ı��߿�
        CWCalcrFont FontDemo3, HelloWorld, cx, cy, True
        CWPrintFont FontDemo3, HelloWorld, 400 - cx * 0.5!, 300, cx, cy, CWYellow, DT_SINGLELINE Or DT_NOCLIP
        CWDrawHRect 400 - cx * 0.5!, 300, cx, cy, CWBlue
        
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
