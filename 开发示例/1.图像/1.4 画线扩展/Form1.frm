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
    
    '����(��������,���������,�յ������,�յ�������,��ɫ)
    CWDrawLine 80, 50, 150, 50, CWRed
    
    '����2(��������,���������,�յ������,�յ�������,��ɫ[,�����,��ʽ,�߿�,����])
    CWDrawLine2 80, 75, 150, 75, CWRed                          ' ��ԭ��Ա�
    CWDrawLine2 80, 100, 150, 100, CWRed, , , 2!                ' ����
    CWDrawLine2 80, 125, 150, 125, CWRed, True                  ' �������
    CWDrawLine2 80, 150, 150, 150, CWRed, , CWLP_Dash           ' ���� ___ '
    CWDrawLine2 80, 175, 150, 175, CWRed, , CWLP_Dot            ' ���� _ '
    CWDrawLine2 80, 200, 150, 200, CWRed, , CWLP_DashDot        ' ���� ___ _ '
    CWDrawLine2 80, 225, 150, 225, CWRed, , CWLP_DashDotDot     ' ���� _ ___ _ '
    CWDrawLine2 80, 250, 150, 250, CWRed, , CWLP_Minus          ' ���� __ '
    CWDrawLine2 80, 275, 150, 275, CWRed, , CWLP_DashMinus      ' ���� ___ __ '
    CWDrawLine2 80, 300, 150, 300, CWRed, , CWLP_MinusDot       ' ���� __ _ '
    CWDrawLine2 80, 325, 150, 325, CWRed, , CWLP_MinusDotDot    ' ���� _ __ _ '
    CWDrawLine2 80, 350, 150, 350, CWRed, , CWLP_Point          ' ����. . . ��
    CWDrawLine2 80, 375, 150, 375, CWRed, , CWLP_InvPoint       ' ���� . . .��
    CWDrawLine2 80, 400, 150, 400, CWRed, , CWLP_DotPointPoint  ' ���� . _ .��
    
    '����ݶԱ�
    CWDrawLine 200, 50, 350, 100, CWGreen               ' ԭ�滭��
    CWDrawLine2 200, 75, 350, 125, CWGreen, False       ' �°滭�ߣ��о�ݣ�
    CWDrawLine2 200, 100, 350, 150, CWGreen, True       ' �°滭�ߣ�����ݣ�
    CWDrawLine2 200, 125, 350, 175, CWGreen, False, , 2!    ' �°滭�ߣ��Ӵ֡��о�ݣ�
    CWDrawLine2 200, 150, 350, 200, CWGreen, True, , 2!     ' �°滭�ߣ��Ӵ֡�����ݣ�
    CWDrawLine2 200, 175, 350, 225, CWGreen, False, , 5!    ' �°滭�ߣ����֡��о�ݣ�
    CWDrawLine2 200, 200, 350, 250, CWGreen, True, , 5!     ' �°滭�ߣ����֡�����ݣ�
    CWDrawLine2 200, 225, 350, 275, CWGreen, False, CWLP_DashDot, 2!        ' �°滭�ߣ��о�ݡ����ߡ��Ӵ֡�������
    CWDrawLine2 200, 250, 350, 300, CWGreen, True, CWLP_DashDot, 2!         ' �°滭�ߣ�����ݡ����ߡ��Ӵ֡�������
    CWDrawLine2 200, 275, 350, 325, CWGreen, False, CWLP_DashDot, 2!, 2!    ' �°滭�ߣ��о�ݡ����ߡ��Ӵ֡�������
    CWDrawLine2 200, 300, 350, 350, CWGreen, True, CWLP_DashDot, 2!, 2!     ' �°滭�ߣ�����ݡ����ߡ��Ӵ֡�������
    
    '������ɫ��(��������,���������,�յ������,�յ�������,�����ɫȨ��,�ص���ɫȨ��,�����ɫ,�յ���ɫ)
    CWDrawLineEx 400, 50, 550, 50, 1, 2.5, CWRed, CWBlue
    
    '�������(��������,��ɫ[,�����,��ʽ,�߿�,����])
    Dim pts(0 To 5) As D3DXVECTOR2
    MakeStar pts, 100, 500, 200
    CWDrawLine2Ex pts, CWBlue
    MakeStar pts, 100, 500, 400
    CWDrawLine2Ex pts, CWBlue, True, CWLP_Dot, 2!, 2!
    
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

Private Sub MakeStar(pts() As D3DXVECTOR2, ByVal R!, ByVal X!, ByVal Y!)
    Dim I&, L&, Ang!: L = UBound(pts) - LBound(pts)
    For I = LBound(pts) To UBound(pts)
      With pts(I)
        Ang = I * 4! * Pi / L
        .X = X + R * Sin(Ang)
        .Y = Y - R * Cos(Ang)
      End With
    Next
    
End Sub
