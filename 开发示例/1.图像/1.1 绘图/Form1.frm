VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ͼ"
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
    
    Do While CWGameRun = True         '������Ϸѭ��
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '����Ƿ������Ⱦ���豸�������Ҵ���δ��С��ʱ��Ⱦ��
    CWBeginScene    '׼���û��Ƴ���
    
    '������ɫ:����ʹ��BASģ�����Ѿ����úõĳ�����ɫ,Ҳ����CWColorARGB��������ת��,A����Ϊ��͸����,RGB������Ӧ����ɫ��ο�RGB��ɫ��
    '������ɫȨ��:��ɫȨ��Խ��,�õ����ɫ�ڽ����������ռ�������Խ��
    
    CWDrawPoint 50, 50, CWRed
    '����(������,������,��ɫ)
    
    CWDrawline 80, 50, 150, 50, CWRed
    '����(��������,���������,�յ������,�յ�������,��ɫ)
         
    CWDrawlineEX 200, 50, 350, 50, 1, 2.5, CWRed, CWBlue
    '������ɫ��(��������,���������,�յ������,�յ�������,�����ɫȨ��,�ص���ɫȨ��,�����ɫ,�յ���ɫ)
    
    CWDrawHRect 400, 10, 150, 80, CWYellow
    '�����ľ���(��������,���������,���,�߶�,��ɫ)
    
    CWDrawSRect 600, 10, 150, 80, CWGreen
    '��ʵ�ľ���(��������,���������,���,�߶�,��ɫ)
    
    CWDrawSRectXGC 50, 100, 150, 150, 1.5, 1, CWPurple, CWCyan
    '�����򽥱�ɫʵ�ľ���(��������,���������,���,�߶�,�����ɫȨ��,�ұ���ɫȨ��,�����ɫ,�ұ���ɫ)
    
    CWDrawSRectYGC 250, 100, 150, 150, 1, 1.5, CWBlue, CWYellow
    '�����򽥱�ɫʵ�ľ���(��������,���������,���,�߶�,�ϱ���ɫȨ��,�±���ɫȨ��,�ϱ���ɫ,�±���ɫ)
    
    CWDrawSRectXCGC 450, 100, 150, 150, 2, 1, CWRed, CWYellow
    '���������Ľ���ɫʵ�ľ���(��������,���������,���,�߶�,������ɫȨ��,������ɫȨ��,������ɫ, ������ɫ)
    
    CWDrawSRectYCGC 650, 100, 150, 150, 2, 1, CWPurple, CWBlue
    '���������Ľ���ɫʵ�ľ���(��������,���������,���,�߶�,������ɫȨ��,������ɫȨ��,������ɫ, ������ɫ)

    CWDrawSRectCGC 30, 300, 150, 150, 1.5, 1, CWRed, CWColorARGB(255, 125, 0, 125)
    '�����Ľ���ɫʵ�ľ���(��������,���������,���,�߶�,������ɫȨ��,�ܱ���ɫȨ��,������ɫ, �ܱ���ɫ)
    
    CWDrawHCircle 300, 400, 100, CWBlue
    '������Բ(Բ�ĺ�����,Բ��������,�뾶,��ɫ)
    
    CWDrawSCircle 500, 500, 100, CWPurple
    '��ʵ��Բ(Բ�ĺ�����,Բ��������,�뾶,��ɫ)

    CWDrawSCircleEX 620, 400, 150, 3, 1, CWHA_Yellow, CWRed
    '�����Ľ���ɫʵ��Բ(Բ�ĺ�����,Բ��������,�뾶,������ɫȨ��,�ܱ���ɫȨ��,������ɫ, �ܱ���ɫ)
         
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
