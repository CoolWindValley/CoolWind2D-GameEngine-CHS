VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ֱ�"
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
      CWPaintPicBegin
        With CWJoystick(0)
         CWPrintFont FontDemo, .X & " " & .Y & vbNewLine & .Z & " " & .R, 0, 0, 800, 60, CWCyan, CWF_LeftAl Or DT_NOCLIP    '��ʾ�ֱ�����
         CWPrintFont FontDemo, IIf(.IsConnected, "�ֱ�������", "�ֱ�δ����"), 400, 0, 800, 60, CWCyan, CWF_LeftAl Or DT_NOCLIP Or DT_VCENTER
         CWPrintFont FontDemo, "�������" & .Pov & "��", 400, 80, 800, 60, CWCyan, CWF_LeftAl Or DT_NOCLIP Or DT_VCENTER
         
         If .Btn(1).PDown Then          '��⹦�ܼ�1���ڰ���״̬
         CWPrintFont FontDemo, "�ֱ����ܼ�1���ڰ���״̬", 0, 100, 800, 50, CWRed, CWF_LeftAl Or DT_NOCLIP Or DT_VCENTER
         End If
         
         If .Btn(2).PDownMoment Then    '��⹦�ܼ�2���µ�˲��
         CWPrintFont FontDemo, "�ֱ����ܼ�2���µ�˲��", 0, 150, 800, 50, CWGreen, CWF_LeftAl Or DT_NOCLIP Or DT_VCENTER
         End If
         
         If .Btn(3).PUPMoment Then      '��⹦�ܼ�3̧���˲��
         CWPrintFont FontDemo, "�ֱ����ܼ�3̧���˲��", 0, 200, 800, 50, CWBlue, CWF_LeftAl Or DT_NOCLIP Or DT_VCENTER
         End If
         
         If .Btn(4).PUP Then            '��⹦�ܼ�4����̧��״̬
         CWPrintFont FontDemo, "�ֱ����ܼ�4����̧��״̬", 0, 250, 800, 50, CWRed, CWF_LeftAl Or DT_NOCLIP Or DT_VCENTER
         End If
         
         If .Btn(9).PDown Then          '���Select������״̬
         CWPrintFont FontDemo, "�ֱ�Select�����µ�״̬", 0, 400, 800, 60, CWGreen, CWF_LeftAl Or DT_NOCLIP Or DT_VCENTER
         End If
         
         If .Btn(10).PUP Then           '���Start��̧��״̬
         CWPrintFont FontDemo, "�ֱ�Start��̧���״̬", 0, 500, 800, 60, CWBlue, CWF_LeftAl Or DT_NOCLIP Or DT_VCENTER
         End If
         
        End With
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
