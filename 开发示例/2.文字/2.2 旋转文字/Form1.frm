VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ת����"
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

Const TestText$ = "Cool Wind 2D" & vbNewLine & "������ת����"

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

Dim FontDemo As CWFont    '����CoolWind�����������
Dim FontRoll As Single

    CWLoadFont FontDemo, "Microsoft YaHei", 32, CWF_Normal, False
    '��������(�������,��������,�����С,�Ƿ�Ϊ����,�Ƿ�Ϊб��)
        'ע��:������������ͱ������������и���Ϸ��ϵͳ�ϴ��ڵ�����.����Ĭ�ϰ�"����"����
        'С����:�ڴ����������ǰ����@��������������~!
    
    Do While CWGameRun = True         '������Ϸѭ��
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '����Ƿ������Ⱦ���豸�������Ҵ���δ��С��ʱ��Ⱦ��
    CWBeginScene CWColorRGBA(255, 240, 200)  '׼���û��Ƴ���
        CWPaintPicBegin
        
        Dim rcos!, rsin!
        rcos = Cos(FontRoll)
        rsin = Sin(FontRoll)
        
        ' ע�⣺�޸�WorldTransform��Խ����������л��Ʋ����ı䣬ֱ���ٴ��޸ġ�
        With WorldTransform     ' �Զ�������任��˳ʱ����ת��
            ' m11 ΪX���ţ�m22ΪY���ţ�m12��X��Y�б䣬m21��Y��X�б䡣
            ' ������Ϊ���ң�ͬ�������б�Ϊ���ң������ʱ���Ϳ���ʵ����תЧ����
            ' ע������sin�ġ���ϵ������ת����˳ʱ��Ϊ+-����ʱ��Ϊ-+��
            .m11 = rcos: .m12 = rsin
            .m21 = -rsin: .m22 = rcos
            .mdx = 200!: .mdy = 150!    ' ��ת���ƽ������ע�������������������������꣩
        End With
        PrintFontEx FontDemo, CWHA_Red, CWGreen, CWBlue
        
        With WorldTransform     ' �Զ�������任����ʱ����ת��
            .m11 = rcos: .m12 = -rsin
            .m21 = rsin: .m22 = rcos
            .mdx = 200!: .mdy = 350!    ' ��ת���ƽ������ע�������������������������꣩
        End With
        PrintFontEx FontDemo, CWHA_Blue, CWRed, CWGreen
        
        ' ��������
        With WorldTransform     ' �Զ�������任
            .m11 = rcos: .m12 = rsin
            .m21 = rsin: .m22 = rcos
            .mdx = 400!: .mdy = 150!
        End With
        PrintFontEx FontDemo, CWHA_Purple, CWYellow, CWCyan
        
        With WorldTransform     ' �Զ�������任
            .m11 = rcos: .m12 = -rsin
            .m21 = -rsin: .m22 = rcos
            .mdx = 400!: .mdy = 350!
        End With
        PrintFontEx FontDemo, CWHA_Cyan, CWPurple, CWYellow
        
        With WorldTransform     ' �Զ�������任
            .m11 = rcos: .m12 = rsin
            .m21 = rsin: .m22 = rcos
            .mdx = 600!: .mdy = 150!
            PrintFontEx FontDemo, CWHA_Orange, CWkelly, CWFuchsia
            .m21 = -.m21
            PrintFontEx FontDemo, CWHA_Violet, CWTurquoise, CWCyanine
        End With
        
        With WorldTransform     ' �Զ�������任
            .m11 = rcos: .m12 = -rsin
            .m21 = -rsin: .m22 = rcos
            .mdx = 600!: .mdy = 350!
            PrintFontEx FontDemo, CWHA_kelly, CWFuchsia, CWOrange
            .m21 = -.m21
            PrintFontEx FontDemo, CWHA_Turquoise, CWViolet, CWCyanine
        End With
        
        With WorldTransform     ' �Զ�������任
            .m11 = rcos: .m12 = 0!
            .m21 = 0!: .m22 = 1!
            .mdx = 200!: .mdy = 550!
        End With
        PrintFontEx FontDemo, CWHA_Green, CWBlue, CWRed
        
        With WorldTransform     ' �Զ�������任
            .m11 = 1!: .m12 = 0!
            .m21 = 0!: .m22 = rcos
            .mdx = 400!: .mdy = 550!
        End With
        PrintFontEx FontDemo, CWHA_Yellow, CWCyan, CWPurple
        
        With WorldTransform     ' �Զ�������任
            .m11 = rcos: .m12 = 0!
            .m21 = 0!: .m22 = rcos
            .mdx = 600!: .mdy = 550!
        End With
        PrintFontEx FontDemo, CWHA_kelly, CWCyanine, CWFuchsia
        
        FontRoll = FontRoll + 0.05!
        If FontRoll > Pi Then FontRoll = FontRoll - 2! * Pi
        
        ' ��������任Ϊ��λ��������Ϊԭʼ״̬��
        WorldTransform = MatrixIdentity
        CWPrintFontTop FontDemo, "FPS: " & CWFPS, 10, 10, 800, 60, CWViolet, CWF_LeftAl  '��ʾ��ǰFPS
        
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

Private Sub PrintFontEx(FontDemo As CWFont, ByVal BackColor As CWColorConstants, ByVal TextColor As CWColorConstants, ByVal EdgeColor As CWColorConstants)
    Dim CX&, CY&    '�����ı���ʾ��С���������ı��߿�
    CWCalcrFont FontDemo, TestText, CX, CY, False
    ' ��ʾ����ͼ������ԭʼ����Ϊ�������꣨������Ϊ��תǰ�������ڶ�λ��ת�㣩
    ' (0, 0)�����ʾ�����Ͻ���ת��(-���, -�߶�)�����ʾ�����½���ת��(-0.5 * ���, -0.5 * �߶�)�����ʾ�����ĵ���ת��
    CWPrintFontTop FontDemo, TestText, CX * -0.5!, CY * -0.5!, CX, CY, TextColor, DT_CENTER Or DT_NOCLIP
    CX = CX + 6: CY = CY + 4
    CWDrawSRect CX * -0.5!, CY * -0.5!, CX, CY, BackColor
    CWDrawHRect CX * -0.5!, CY * -0.5!, CX, CY, EdgeColor
End Sub
