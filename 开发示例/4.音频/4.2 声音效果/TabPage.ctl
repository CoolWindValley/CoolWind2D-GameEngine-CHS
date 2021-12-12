VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TabPage 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   ScaleHeight     =   2625
   ScaleWidth      =   4950
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   20000
      SmallChange     =   1000
      Max             =   100000
      SelStart        =   100000
      TickFrequency   =   2000
      Value           =   100000
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   20000
      SmallChange     =   1000
      Min             =   -100000
      Max             =   100000
      TickFrequency   =   5000
   End
   Begin ComctlLib.Slider Slider3 
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   20000
      SmallChange     =   1000
      Min             =   -100000
      Max             =   100000
      TickFrequency   =   5000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "“Ù∏ﬂ(&P)£∫"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "∆Ω∫‚(&B)£∫"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "“Ù¡ø(&V)£∫"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "TabPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public PitchBase As Single

Const GWL_STYLE& = -16&
Const BS_AUTOCHECKBOX& = &H3&
Const BM_GETSTATE& = &HF2&
Const BST_CHECKED& = &H1&
Const BS_PUSHLIKE& = &H1000&
Const BS_FLAT& = &H8000&

Enum SoundOption
    soVolume
    soBalance
    soPitch
End Enum

Event ButtonClick(ByVal Check As Boolean)
Event SliderChange(ByVal SndOpt As SoundOption, ByVal Value As Single)

Private Sub Command1_Click()
    RaiseEvent ButtonClick(SendMessage(Command1.hWnd, BM_GETSTATE, 0, 0) And BST_CHECKED)
End Sub

Private Sub Slider1_Scroll()
    RaiseEvent SliderChange(soVolume, Me.Volume)
End Sub

Private Sub Slider2_Scroll()
    RaiseEvent SliderChange(soBalance, Me.Balance)
End Sub

Private Sub Slider3_Change()
    RaiseEvent SliderChange(soPitch, Me.Pitch)
End Sub

Public Property Get ButtonCheck() As Boolean
    ButtonCheck = GetWindowLong(Command1.hWnd, GWL_STYLE) And BS_AUTOCHECKBOX
End Property

Public Property Let ButtonCheck(ByVal vNewValue As Boolean)
    Dim Style&: Style = GetWindowLong(Command1.hWnd, GWL_STYLE)
    If vNewValue Then
        Style = Style Or BS_AUTOCHECKBOX Or BS_PUSHLIKE 'Or BS_FLAT
        Command1.Caption = "4;"
    Else
        Style = Style And Not (BS_AUTOCHECKBOX Or BS_PUSHLIKE) 'Or BS_FLAT
        Command1.Caption = "4"
    End If
    SetWindowLong Command1.hWnd, GWL_STYLE, Style
End Property

Public Property Get Volume() As Single
    Volume = Slider1.Value * 0.01!
End Property

Public Property Get Balance() As Single
    Balance = Slider2.Value * 0.01!
End Property

Public Property Get Pitch() As Single
    Pitch = PitchBase ^ (Slider3.Value * 0.00001!)
End Property

Private Sub UserControl_InitProperties()
    Call UserControl_ReadProperties(New PropertyBag)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    Me.ButtonCheck = .ReadProperty("ButtonCheck", False)
    PitchBase = .ReadProperty("PitchBase", 4!)
  End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "ButtonCheck", Me.ButtonCheck, False
    .WriteProperty "PitchBase", PitchBase, 4!
  End With
End Sub
