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


Dim PicDemo1 As CWPic, PicDemo2 As CWPic    '����CoolWind����ͼƬ����

    CWLoadPic PicDemo1, App.Path & "\Pic\M1.png", CWColorNone
    CWLoadPic PicDemo2, App.Path & "\Pic\M2.png", CWColorNone
    '����ͼƬ(ͼƬ����,ͼƬ·��,����ɫ)
    '��������ɫ:����˼��,���ǲ���ʾͼƬ�ϵ�ָ����ĳ����ɫ,ͨ����Ϊ��ɫ����(���������κ���ɫ)
    
    Do While CWGameRun = True         '������Ϸѭ��
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '����Ƿ������Ⱦ���豸�������Ҵ���δ��С��ʱ��Ⱦ��
    CWBeginScene CWBlue '׼���û��Ƴ���
         
    CWPaintPicBegin     '��ͼ��ʼ ��������ͼ�����ɶԳ��֣�����֪ͨ����׼����ͼ
    
    '��ͼ��ʼ����ͼ����֮�䲻Ҫʹ��CWDraw��ͼ���������򽫲���һЩ��ֵ�Ч����

         CWPaintPic PicDemo1, 100, 100
         CWPaintPic PicDemo2, 400, 100
         '��ͼ(ͼƬ,��������,���������)
         
         CWPaintPicEX PicDemo1, 100, 400, 0, 0, 120, 137, CWYellow    '��ͼ��չ����,�ܲü���ͼ���ı�ͼ��ɫ��
         '�߼���ͼ(ͼƬ,��������,���������,�ü���������,�ü����������,�ü����,�ü��߶�,ͼ��ɫ��)
            'ע��:�ü�������������ͼƬ������Ϊ����ϵ��.   ������ı�ͼ��ɫ����������Ϊ��ɫ(CWWhite)����

    CWPaintPicEXEX PicDemo2, 350, 450, 0, 0, 208, 145, 1.5, 1.5, 350, 450, CRad(-45), CWRed   '��ͼ��չ����,�ܲü�����ת��ͼ���ı�ͼ��ɫ������С
    '���߼���ͼ(ͼƬ,��������,���������,�ü���������,�ü����������,�ü����,�ü��߶�,�������ű���,�������ű���, _
                ��ת�������,��ת��������,��ת�Ƕ�,ͼ��ɫ��)
        'ע��:��ת������������Ϸ������Ϊ����ϵ��.   ���ű�������һΪ�Ŵ�,С��һΪ��С.
                '��ת�ǶȲ��û�����,����Ϊ˳ʱ����ת,����Ϊ��ʱ����ת.  ����CRad�������Ƕ�ת��Ϊ����

    CWPaintPicEnd       '��ͼ���� ��������ͼ��ʼ�ɶԳ��֣�����֪ͨ������ͼ���
   
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
