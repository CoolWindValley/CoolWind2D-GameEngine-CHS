VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
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

Dim PicDemo As CWPic    '����CoolWind����ͼƬ����

    CWLoadPic PicDemo, App.Path & "\Pic\bgimg.png", CWColorNone     '�������ͼƬ

    Do While CWGameRun = True         '������Ϸѭ��
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '����Ƿ������Ⱦ���豸�������Ҵ���δ��С��ʱ��Ⱦ��
    CWBeginScene    '׼���û��Ƴ���
         
      CWPaintPicBegin     '��ͼ��ʼ
    
        ' ���ò���ģʽΪѭ�������Է���ʵ��ƽ�̺�ѭ��������
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_WRAP
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_WRAP
        ' ������ǰ�Ƚ���Alpha��ϣ����Ч�ʣ�
        CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, False
        CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, False

        CWPaintPicEx PicDemo, 0, 0, 0, 0, 800, 600, CWGrey '���ϵ�����ͼ����ϻ�ɫģ�º�ҹ

      CWPaintPicEnd       '��ͼ����
 
      ' ��ͼ���������ύ���飬ֱ�ӿ���Alpha���
      CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, True
      CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, True

      LightEFOpen       '�򿪹���Ч��
         
        '����Ч�����������е���ͼ�ͻ�ͼ���ᱻ��Ϊ��Դ����͸���ȣ�A������ǿ�ȣ���ɫ��RGB������ɫ��
        CWDrawSCircle 400, 300, 192, CWWhite                    '��Բ��Ϊ��Դ
        CWDrawSCircleEX 400, 300, 192, 2, 1, CWYellow, CWHA_Red '�ص��Ĺ�Դͼ��Խ�࣬����Խǿ
    
            'ע�⣺���ֻ��ƣ������ʾ�������ᵽ�����ܹ���Ӱ��

      LightEFClose      '�رչ���Ч��

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
