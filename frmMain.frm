VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Aero Color Customizer"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdToggleDWM 
      Caption         =   "T&oggle DWM"
      Height          =   375
      Left            =   1440
      TabIndex        =   26
      Top             =   4320
      Width           =   1215
   End
   Begin prjAeroCustomizer.SSTabEx ssColorTabs 
      Height          =   1965
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3466
      Tabs            =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   2
      Style           =   2
      TabHeight       =   563
      AutoTabHeight   =   -1  'True
      TabCaption(0)   =   "Color"
      Tab(0).ControlCount=   10
      Tab(0).Control(0)=   "tkColorBrightness"
      Tab(0).Control(1)=   "tkColorSaturation"
      Tab(0).Control(2)=   "tkColorHue"
      Tab(0).Control(3)=   "Image3(0)"
      Tab(0).Control(4)=   "Image2(0)"
      Tab(0).Control(5)=   "Image1(0)"
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(9)=   "spColor"
      TabCaption(1)   =   "After Glow"
      Tab(1).ControlCount=   10
      Tab(1).Control(0)=   "tkAfterGlowBrightness"
      Tab(1).Control(1)=   "tkAfterGlowSaturation"
      Tab(1).Control(2)=   "tkAfterGlowHue"
      Tab(1).Control(3)=   "Image3(1)"
      Tab(1).Control(4)=   "Image2(1)"
      Tab(1).Control(5)=   "Image1(1)"
      Tab(1).Control(6)=   "spAfterGlow"
      Tab(1).Control(7)=   "Label5"
      Tab(1).Control(8)=   "Label6"
      Tab(1).Control(9)=   "Label7"
      Begin prjAeroCustomizer.Slider tkAfterGlowBrightness 
         Height          =   330
         Left            =   -73680
         TabIndex        =   32
         Top             =   1440
         Width           =   3570
         _ExtentX        =   582
         _ExtentY        =   582
         Max             =   240
         Style           =   1
         BackColor       =   192
         TickMarks       =   0   'False
         CenterLine      =   0   'False
         Transparent     =   -1  'True
      End
      Begin prjAeroCustomizer.Slider tkAfterGlowSaturation 
         Height          =   330
         Left            =   -73680
         TabIndex        =   31
         Top             =   960
         Width           =   3570
         _ExtentX        =   582
         _ExtentY        =   582
         Max             =   240
         Style           =   1
         BackColor       =   192
         TickMarks       =   0   'False
         CenterLine      =   0   'False
         Transparent     =   -1  'True
      End
      Begin prjAeroCustomizer.Slider tkAfterGlowHue 
         Height          =   330
         Left            =   -73680
         TabIndex        =   30
         Top             =   480
         Width           =   3570
         _ExtentX        =   582
         _ExtentY        =   582
         Max             =   240
         Style           =   1
         BackColor       =   192
         TickMarks       =   0   'False
         CenterLine      =   0   'False
         Transparent     =   -1  'True
      End
      Begin prjAeroCustomizer.Slider tkColorBrightness 
         Height          =   330
         Left            =   1320
         TabIndex        =   29
         Top             =   1440
         Width           =   3570
         _ExtentX        =   582
         _ExtentY        =   582
         Max             =   240
         Style           =   1
         BackColor       =   192
         TickMarks       =   0   'False
         CenterLine      =   0   'False
         Transparent     =   -1  'True
      End
      Begin prjAeroCustomizer.Slider tkColorSaturation 
         Height          =   330
         Left            =   1320
         TabIndex        =   28
         Top             =   960
         Width           =   3570
         _ExtentX        =   582
         _ExtentY        =   582
         Max             =   240
         Style           =   1
         BackColor       =   192
         TickMarks       =   0   'False
         CenterLine      =   0   'False
         Transparent     =   -1  'True
      End
      Begin prjAeroCustomizer.Slider tkColorHue 
         Height          =   330
         Left            =   1320
         TabIndex        =   27
         Top             =   480
         Width           =   3570
         _ExtentX        =   582
         _ExtentY        =   582
         Max             =   240
         Style           =   1
         BackColor       =   192
         TickMarks       =   0   'False
         CenterLine      =   0   'False
         Transparent     =   -1  'True
      End
      Begin VB.Image Image3 
         Height          =   135
         Index           =   1
         Left            =   -73680
         Picture         =   "frmMain.frx":0442
         Top             =   1545
         Width           =   3585
      End
      Begin VB.Image Image2 
         Height          =   135
         Index           =   1
         Left            =   -73680
         Picture         =   "frmMain.frx":1DD4
         Top             =   1065
         Width           =   3585
      End
      Begin VB.Image Image1 
         Height          =   135
         Index           =   1
         Left            =   -73680
         Picture         =   "frmMain.frx":3766
         Top             =   585
         Width           =   3585
      End
      Begin VB.Image Image3 
         Height          =   135
         Index           =   0
         Left            =   1320
         Picture         =   "frmMain.frx":50F8
         Top             =   1545
         Width           =   3585
      End
      Begin VB.Image Image2 
         Height          =   135
         Index           =   0
         Left            =   1320
         Picture         =   "frmMain.frx":6A8A
         Top             =   1065
         Width           =   3585
      End
      Begin VB.Image Image1 
         Height          =   135
         Index           =   0
         Left            =   1320
         Picture         =   "frmMain.frx":841C
         Top             =   585
         Width           =   3585
      End
      Begin VB.Shape spAfterGlow 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  '투명하지 않음
         Height          =   495
         Left            =   -69840
         Shape           =   3  '원형
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "B&rightness:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   25
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "S&aturation:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "H&ue:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "&Hue:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "&Saturation:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "&Brightness:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1500
         Width           =   855
      End
      Begin VB.Shape spColor 
         BackColor       =   &H000000FF&
         BackStyle       =   1  '투명하지 않음
         Height          =   495
         Left            =   5160
         Shape           =   3  '원형
         Top             =   1320
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdRestartDWM 
      Caption         =   "Restart D&WM"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox chkShiftAnimations 
      Caption         =   "<Sh&ift> for slow animations"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox chkHideBlur 
      Caption         =   "&Disable Blur"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CheckBox chkTransparent 
      Caption         =   "E&nable Transparency"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Caption         =   "S&tripes Intensity"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   11
      Top             =   2880
      Width           =   2895
      Begin prjAeroCustomizer.Slider tkStripesIntensity 
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2610
         _ExtentX        =   582
         _ExtentY        =   582
         Max             =   50
         Style           =   1
         TickMarks       =   0   'False
         TickMarkCnt     =   10
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Blur Balanc&e"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   2895
      Begin prjAeroCustomizer.Slider tkBlurBalance 
         Height          =   330
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2610
         _ExtentX        =   582
         _ExtentY        =   582
         Max             =   127
         Style           =   1
         TickMarks       =   0   'False
         TickMarkCnt     =   10
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "A&fter Glow Balance"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   7
      Top             =   2160
      Width           =   2895
      Begin prjAeroCustomizer.Slider tkAfterGlowBalance 
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2610
         _ExtentX        =   582
         _ExtentY        =   582
         Max             =   127
         Style           =   1
         TickMarks       =   0   'False
         TickMarkCnt     =   10
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Color Balance"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2895
      Begin prjAeroCustomizer.Slider tkColorBalance 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2610
         _ExtentX        =   582
         _ExtentY        =   582
         Max             =   127
         Style           =   1
         TickMarks       =   0   'False
         TickMarkCnt     =   20
      End
   End
   Begin VB.Frame pnColor 
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   1
      Left            =   8400
      TabIndex        =   0
      Top             =   2160
      Width           =   5775
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "&Alpha:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame pnColor 
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   2
      Left            =   7080
      TabIndex        =   1
      Top             =   600
      Width           =   5775
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "A&lpha:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderStyle     =   0  '투명
      Height          =   255
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderStyle     =   0  '투명
      Height          =   255
      Left            =   3075
      Shape           =   4  '둥근 사각형
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderStyle     =   0  '투명
      Height          =   255
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   3960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape Translucent 
      BackColor       =   &H00000000&
      BackStyle       =   1  '투명하지 않음
      BorderStyle     =   0  '투명
      Height          =   975
      Left            =   100
      Top             =   5355
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Label lblDebug 
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10080
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OriginalParams As DWM_COLORIZATION_PARAMS
Dim objShell As Object
Dim DiscardChanges As Boolean
Dim osVersion As Double
Dim InitialHideBlur As Integer
Dim InitialShift As Integer
Dim clrKey As Long

Const WM_APP As Long = 32768
Const WM_DWMCOMPOSITIONCHANGED As Long = &H31E
Const DWM_EC_DISABLECOMPOSITION As Long = 0
Const DWM_EC_ENABLECOMPOSITION As Long = 1
Const GWL_WNDPROC = (-4)
Const GWL_EXSTYLE As Long = (-20&)
Const WS_EX_LAYERED As Long = &H80000
Const LWA_COLORKEY As Long = &H1
Const WM_PAINT As Long = &HF
Const WM_PRINTCLIENT As Long = &H318

Sub OnDWMChange()
    On Error Resume Next
    Dim i%
    If IsDWMEnabled() Then
        'Me.BackColor = &H0&
        Translucent.BackColor = 0
        For i = RECT.LBound To RECT.UBound
            RECT(i).BackColor = &H404040
        Next i
        ExtendDWMFrame Me, 0, 0, 49 - 50, 0
        'Me.Height = 6510
        
        Me.BackColor = clrKey
        Frame1.BackColor = clrKey
        Frame2.BackColor = clrKey
        Frame3.BackColor = clrKey
        Frame4.BackColor = clrKey
        ssColorTabs.BackColor = clrKey
        cmdRestartDWM.BackColor = clrKey
        cmdOK.BackColor = clrKey
        cmdCancel.BackColor = clrKey
        chkTransparent.BackColor = clrKey '&HFFFFFF
        'Shape3.Visible = -1
        chkHideBlur.BackColor = clrKey '&HFFFFFF
        'Shape2.Visible = -1
        chkShiftAnimations.BackColor = clrKey '&HFFFFFF
        'Shape1.Visible = -1
        cmdToggleDWM.BackColor = clrKey
        tkColorBalance.BackColor = clrKey
        tkAfterGlowBalance.BackColor = clrKey
        tkBlurBalance.BackColor = clrKey
        tkStripesIntensity.BackColor = clrKey
    Else
        Translucent.BackColor = &H8000000F
        Me.BackColor = &H8000000F
        For i = RECT.LBound To RECT.UBound
            RECT(i).BackColor = &H8000000F
        Next i
        'Me.Height = 5775
        
        Me.BackColor = &H8000000F
        Frame1.BackColor = &H8000000F
        Frame2.BackColor = &H8000000F
        Frame3.BackColor = &H8000000F
        Frame4.BackColor = &H8000000F
        ssColorTabs.BackColor = &H8000000F
        cmdRestartDWM.BackColor = &H8000000F
        cmdOK.BackColor = &H8000000F
        cmdCancel.BackColor = &H8000000F
        chkTransparent.BackColor = &H8000000F
        Shape3.Visible = 0
        chkHideBlur.BackColor = &H8000000F
        Shape2.Visible = 0
        chkShiftAnimations.BackColor = &H8000000F
        Shape1.Visible = 0
        cmdToggleDWM.BackColor = &H8000000F
        tkColorBalance.BackColor = &H8000000F
        tkAfterGlowBalance.BackColor = &H8000000F
        tkBlurBalance.BackColor = &H8000000F
        tkStripesIntensity.BackColor = &H8000000F
    End If
End Sub

Private Sub chkHideBlur_Click()
        objShell.RegWrite "HKCU\Software\Microsoft\Windows\DWM\HideBlur", chkHideBlur.value, "REG_DWORD"
End Sub

Private Sub chkShiftAnimations_Click()
    objShell.RegWrite "HKCU\Software\Microsoft\Windows\DWM\AnimationsShiftKey", chkShiftAnimations.value, "REG_DWORD"
End Sub

Private Sub chkTransparent_Click()
    SetParameters
End Sub

Private Sub cmdApply_Click()
    
End Sub

Private Sub cmdOK_Click()
    objShell.RegWrite "HKCU\Software\Microsoft\Windows\DWM\HideBlur", chkHideBlur.value, "REG_DWORD"
    objShell.RegWrite "HKCU\Software\Microsoft\Windows\DWM\AnimationsShiftKey", chkShiftAnimations.value, "REG_DWORD"
    SetParameters True
    DiscardChanges = False
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    DiscardChanges = True
    Unload Me
End Sub

Private Sub cmdRestartDWM_Click()
    If osVersion >= 6.2 Or (osVersion <= 6.1 And Build <> 6000 And Build <> 6001 And Build <> 6002 And Build <> 6003 And Build <> 5920 And Build <> 6801 And Build <> 7000 And Build <> 7100 And Build <> 7600 And Build <> 7601 And Build <> 8102 And Build <> 8250 And Build <> 8400 And Build <> 9200 And Build <> 9600) Then
        If IsDWMEnabled() Then SendKeys "^+{F9}"
        SendKeys "^+{F9}"
    Else
        DwmEnableComposition DWM_EC_DISABLECOMPOSITION
        DwmEnableComposition DWM_EC_ENABLECOMPOSITION
    End If
End Sub

Private Sub cmdToggleDWM_Click()
    SendKeys "^+{F9}"
End Sub

Private Sub Form_Load()
    Dim Build As Long
    Build = GetBuild()
    osVersion = GetVersion()
    If osVersion < 6# Or osVersion > 6.4 Or Build < 5252 Then
        Alert "Unsupported operating system! Requires Windows Vista Beta to Windows 8 Beta!!!", App.Title, Me, 16
        End
    End If

    Set objShell = CreateObject("WScript.Shell")
    DiscardChanges = True
    OnConpositionChanged Me
    SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    clrKey = &HDCDCDC - &H10000
    SetLayeredWindowAttributes Me.hWnd, clrKey, 0, LWA_COLORKEY
    
    'WM_DWMCOMPOSITIONCHANGED
    OnDWMChange
    
    Dim Params As DWM_COLORIZATION_PARAMS
    DwmGetColorizationParameters OriginalParams
    DwmGetColorizationParameters Params
    lblDebug.Caption = lblDebug.Caption & "Color: " & Params.ColorRed & "," & Params.ColorGreen & "," & Params.ColorBlue & vbCrLf
    lblDebug.Caption = lblDebug.Caption & "AfterGlow: " & Params.AfterGlowRed & "," & Params.AfterGlowGreen & "," & Params.AfterGlowBlue & vbCrLf
    lblDebug.Caption = lblDebug.Caption & "Intensity: " & Params.ColorBalance & vbCrLf
    lblDebug.Caption = lblDebug.Caption & "AfterGlowBalance: " & Params.AfterGlowBalance & vbCrLf
    lblDebug.Caption = lblDebug.Caption & "BlurBalance: " & Params.BlurBalance & vbCrLf
    lblDebug.Caption = lblDebug.Caption & "Stripes: " & Params.StripesIntensity & vbCrLf
    lblDebug.Caption = lblDebug.Caption & "Opaque: " & Params.Opaque & vbCrLf
    
    spColor.BackColor = CDec(CDec(CDec(Params.ColorBlue) * &H10000) + CDec(CDec(Params.ColorGreen) * &H100) + CDec(Params.ColorRed))
    spAfterGlow.BackColor = CDec(CDec(CDec(Params.AfterGlowBlue) * &H10000) + CDec(CDec(Params.AfterGlowGreen) * &H100) + CDec(Params.AfterGlowRed))
    
    Dim hlsColor(1 To 3) As Integer
    ColorRGBToHLS CDec(CDec(CDec(Params.ColorBlue) * &H10000) + CDec(CDec(Params.ColorGreen) * &H100) + CDec(Params.ColorRed)), hlsColor(1), hlsColor(2), hlsColor(3)
    tkColorHue.value = hlsColor(1)
    tkColorSaturation.value = hlsColor(3)
    tkColorBrightness.value = hlsColor(2)
    'tkColorAlpha.Value = Params.ColorAlpha
    
    ColorRGBToHLS CDec(CDec(CDec(Params.AfterGlowBlue) * &H10000) + CDec(CDec(Params.AfterGlowGreen) * &H100) + CDec(Params.AfterGlowRed)), hlsColor(1), hlsColor(2), hlsColor(3)
    tkAfterGlowHue.value = hlsColor(1)
    tkAfterGlowSaturation.value = hlsColor(3)
    tkAfterGlowBrightness.value = hlsColor(2)
    'tkAfterGlowAlpha.Value = Params.AfterGlowAlpha
    
    tkColorBalance.value = Params.ColorBalance
    tkAfterGlowBalance.value = Params.AfterGlowBalance
    tkBlurBalance.value = Params.BlurBalance
    tkStripesIntensity.value = Params.StripesIntensity
    If Params.Opaque Then
        chkTransparent.value = 0
    Else
        chkTransparent.value = 1
    End If
    On Error Resume Next
    chkHideBlur.value = objShell.RegRead("HKCU\Software\Microsoft\Windows\DWM\HideBlur")
    InitialHideBlur = chkHideBlur.value
    chkShiftAnimations.value = objShell.RegRead("HKCU\Software\Microsoft\Windows\DWM\AnimationsShiftKey")
    InitialShift = chkShiftAnimations.value
    If Build >= 7127 And Build <> 7600 And Build <> 7601 And Build <> 8102 And Build <> 8250 And Build <> 8400 And Build <> 9200 And Build <> 9600 And Build < 10240 Then
        chkHideBlur.Enabled = -1
    Else
        chkHideBlur.Enabled = 0
    End If
    
    If osVersion > 6.3 Or Build = 7950 Or Build = 6000 Or Build = 6001 Or Build = 6002 Or Build = 6003 Or Build = 5920 Or Build = 7000 Or Build = 6801 Or Build = 7100 Or Build = 7600 Or Build = 7601 Or Build = 8102 Or Build = 8250 Or Build = 8400 Or Build = 8888 Or Build = 9200 Or Build = 9600 Then
        cmdRestartDWM.Enabled = False
    End If
    
    If Build = 7950 Or Not (osVersion >= 6.2 Or (osVersion <= 6.1 And Build <> 6000 And Build <> 6001 And Build <> 6002 And Build <> 6003 And Build <> 5920 And Build <> 6801 And Build <> 7000 And Build <> 7100 And Build <> 7600 And Build <> 7601 And Build <> 8102 And Build <> 8250 And Build <> 8400 And Build <> 9200 And Build <> 9600)) Then
        cmdToggleDWM.Enabled = False
    End If
End Sub

Sub OnColorChange()
    spColor.BackColor = ColorHLSToRGB(tkColorHue.value, tkColorBrightness.value, tkColorSaturation.value)
End Sub

Sub OnAfterGlowChange()
    spAfterGlow.BackColor = ColorHLSToRGB(tkAfterGlowHue.value, tkAfterGlowBrightness.value, tkAfterGlowSaturation.value)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DiscardChanges Then
        objShell.RegWrite "HKCU\Software\Microsoft\Windows\DWM\HideBlur", InitialHideBlur, "REG_DWORD"
        objShell.RegWrite "HKCU\Software\Microsoft\Windows\DWM\AnimationsShiftKey", InitialShift, "REG_DWORD"
        DwmSetColorizationParameters OriginalParams, True
    End If
End Sub

Private Sub HScroll1_Scroll()
    Debug.Print HScroll1.value
End Sub

Private Sub tbTabStrip_Click()
    Dim i%
    For i = 1 To tbTabStrip.Tabs.Count
        If tbTabStrip.Tabs(i).Selected Then
            pnColor(i).Visible = -1
            pnColor(i).Top = pnColor(1).Top
            pnColor(i).Left = pnColor(1).Left
        Else
            pnColor(i).Visible = 0
        End If
    Next i
End Sub

Private Sub tkAfterGlowAlpha_Scroll()
    SetParameters
End Sub

Private Sub tkAfterGlowBalance_Scrolling()
    SetParameters
End Sub

Private Sub tkBlurBalance_Scrolling()
    SetParameters
End Sub

Private Sub tkColorAlpha_Scroll()
    SetParameters
End Sub

Private Sub tkColorBalance_Scrolling()
    SetParameters
End Sub

Private Sub tkColorHue_Scrolling()
    OnColorChange
    SetParameters
End Sub

Private Sub tkColorSaturation_Scrolling()
    OnColorChange
    SetParameters
End Sub

Private Sub tkColorBrightness_Scrolling()
    OnColorChange
    SetParameters
End Sub

Private Sub tkAfterGlowHue_Scrolling()
    OnAfterGlowChange
    SetParameters
End Sub

Private Sub tkAfterGlowSaturation_Scrolling()
    OnAfterGlowChange
    SetParameters
End Sub

Private Sub tkAfterGlowBrightness_Scrolling()
    OnAfterGlowChange
    SetParameters
End Sub

Private Sub tkAfterGlowHue_Change()
    OnAfterGlowChange
    SetParameters
End Sub

Private Sub tkAfterGlowSaturation_Change()
    OnAfterGlowChange
    SetParameters
End Sub

Private Sub tkAfterGlowBrightness_Change()
    OnAfterGlowChange
    SetParameters
End Sub

Sub SetParameters(Optional ByVal Commit As Boolean = False)
    Dim Params As DWM_COLORIZATION_PARAMS
    Dim clrColor As Long
    Dim r As Long
    Dim g As Long
    Dim B As Long
    
    clrColor = ColorHLSToRGB(tkColorHue.value, tkColorBrightness.value, tkColorSaturation.value)
    ColorRefToRGB clrColor, r, g, B
    Params.ColorRed = r
    Params.ColorGreen = g
    Params.ColorBlue = B
    'Params.ColorAlpha = tkColorAlpha.Value
    
    clrColor = ColorHLSToRGB(tkAfterGlowHue.value, tkAfterGlowBrightness.value, tkAfterGlowSaturation.value)
    ColorRefToRGB clrColor, r, g, B
    Params.AfterGlowRed = r
    Params.AfterGlowGreen = g
    Params.AfterGlowBlue = B
    'Params.AfterGlowAlpha = tkAfterGlowAlpha.Value
    
    Params.ColorBalance = tkColorBalance.value
    Params.AfterGlowBalance = tkAfterGlowBalance.value
    Params.BlurBalance = tkBlurBalance.value
    Params.StripesIntensity = tkStripesIntensity.value
    
    If chkTransparent.value Then
        Params.Opaque = False
    Else
        Params.Opaque = True
    End If
    
    If Commit Then
        DwmSetColorizationParameters Params, False
    Else
        DwmSetColorizationParameters Params, True
    End If
End Sub

Private Sub tkStripesIntensity_Scrolling()
    SetParameters
End Sub

