VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ۺϲ���"
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

    Dim bg As CWPic, yun As CWPic, spider As CWPic, light As CWPic, fnt As CWFont, bgm As CWMusic
    Const SpiderX& = 400 - 104  ' ͼ���Xλ��
    Const SpiderY& = 410        ' ͼ���Yλ��
    
    ' ������Դ
    CWLoadFont fnt, "Microsoft YaHei", 32, CWF_Bold, False
    CWLoadPic bg, App.Path & "\Pic\bgimg.png", CWColorNone
    CWLoadPic yun, App.Path & "\Pic\yun.png", CWColorNone
    CWLoadPic spider, App.Path & "\Pic\spider.png", CWColorNone
    CWLoadPic light, App.Path & "\Pic\light.png", CWColorNone
    CWLoadMusic bgm, App.Path & "\Snd\test.mp3"
    CWPlayMusic bgm, CWM_Repeat

    Do While CWGameRun         '������Ϸѭ��
    
    If CWD3DDevice9.TestCooperativeLevel = 0 And Me.WindowState <> 1 Then  '����Ƿ������Ⱦ���豸�������Ҵ���δ��С��ʱ��Ⱦ��
    CWBeginScene    '׼���û��Ƴ���
    CWPaintPicBegin
    
        ' ���ò���ģʽΪѭ�������Է���ʵ��ƽ�̺�ѭ��������
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_WRAP
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_WRAP
        ' ������ǰ�Ƚ���Alpha��ϣ����Ч�ʣ�
        CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, False
        CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, False

        Dim l!, R!, i&, j&
        l = l + 1!: If l > 0! Then l = l - 893!
        R = R + 1!: If R > 180! Then R = R - 360!
         
        CWPrintFontTop fnt, "Hello World!", 0, 0, 480, 500, CWBlue, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP

#Const Title = False    ' ƽ�������ַ�ʽʵ��
' True Ϊ���������ʽʵ�֣�Ч�ʸߣ���ֻ��ʵ�ֵ���ͼƬ��
' False Ϊѭ����ͼ��ʽʵ�֣�Ч�ʵ�һ�㣬������ʵ�ֶ���ͼ�����л���
#If Title Then
        CWPaintPicEx bg, 0, 0, 0, 0, 800, 600
#Else
        For i = 0 To 24
            For j = 0 To 18
                CWPaintPic bg, i * 32, j * 32
        Next j, i
#End If
        CWPaintPicFlush ' ���������Ҫ���ύ���ٿ���Alpha���
        CWD3DDevice9.SetRenderState D3DRS_ALPHABLENDENABLE, True
        CWD3DDevice9.SetRenderState D3DRS_ALPHATESTENABLE, True

        CWPaintPicEx yun, 0, 0, l, 0, 800, 325
        CWPaintPicFlush ' �ı�D3D��Ⱦģʽ֮ǰ��Ҫ���ύ������Ⱦ����Ȼ�ᵼ��֮ǰ�Ļ�������Ҳ��ģʽ�޸�Ӱ�죩
        
        ' ���ò���ģʽΪ�߿򣨱߿���ɫĬ��Ϊ͸��ɫ���ʺϷǱ�������ͼ��
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_BORDER
        CWD3DDevice9.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_BORDER

        CWPaintPicExEx spider, 400 - 104, 300 - 72.5, 0, 0, 208, 145, , , 400, 300, R * Pi / 180!
        CWPaintPic spider, SpiderX, SpiderY     ' ֩��ԭͼ������������������в��Եģ�
    
        Dim cwc As CWColor      ' ��ȡ�����λ�õ���ɫֵ
        cwc = CWSplitColor(CWPicGetPixel(spider, CWMouse.X - SpiderX, CWMouse.Y - SpiderY))
        CWPrintFontTop fnt, "XY(" & CWMouse.X & ", " & CWMouse.Y & ")" & vbNewLine & "ARGB(" & cwc.Alpha & ", " & cwc.Red & ", " & cwc.Green & ", " & cwc.Blue & ")", 0, 560, 800, 0, CWGreen, DT_CENTER Or DT_VCENTER Or DT_NOCLIP

        LightEFOpen True
            ' ���������У�����һ�����ȵ�֩��ͼ����Ϊ��Դ�������棨ʵ�ֱ�����Ч����
            If IsHitWnd And cwc.Alpha >= &H80 Then CWPaintPic spider, SpiderX, SpiderY, CWGrey
            ' ����ת�����ŵ�Բ�ι�ߣ���α3D��Ч����
            CWPaintPicExEx1 light, 200, 460, 0, 0, 256, 256, 2!, 0.5!, 128, 128, -R * Pi / 180!
            ' �����ź���ת��Բ�ι�ߣ���ͳ2D��Ч����
            CWPaintPicExEx2 light, 600, 440, 0, 0, 256, 256, 2!, 0.5!, 128, 128, R * Pi / 180!
        LightEFClose
    
        CWDrawSRect 300, 200, 200, 200, &H40FFFF00
        
        CWPrintFontTop fnt, CWFPS, 0, 0, 0, 0, CWRed, DT_SINGLELINE Or DT_NOCLIP
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
