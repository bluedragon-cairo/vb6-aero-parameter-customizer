VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Aero Transparency Customizer"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   31
      Top             =   4995
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   30
      Top             =   4995
      Width           =   1335
   End
   Begin VB.CheckBox chkHideBlur 
      Caption         =   "&Hide Blur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   29
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CheckBox chkTransparent 
      Caption         =   "E&nable Transparency"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "S&tripes Intensity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   26
      Top             =   3600
      Width           =   2895
      Begin ComctlLib.Slider tkStripesIntensity 
         Height          =   465
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   820
         _Version        =   327682
         Max             =   50
         TickFrequency   =   2
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Blur Balanc&e"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   24
      Top             =   3600
      Width           =   2895
      Begin ComctlLib.Slider tkBlurBalance 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   327682
         Max             =   127
         TickStyle       =   3
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "A&fterGlow Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   22
      Top             =   2640
      Width           =   2895
      Begin ComctlLib.Slider tkAfterGlowBalance 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   327682
         Max             =   127
         TickStyle       =   3
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Color Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   20
      Top             =   2640
      Width           =   2895
      Begin ComctlLib.Slider tkColorBalance 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   327682
         Max             =   127
         TickStyle       =   3
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
      Height          =   1815
      Index           =   1
      Left            =   280
      TabIndex        =   0
      Top             =   570
      Width           =   5775
      Begin ComctlLib.Slider tkColorHue 
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   327682
         Max             =   240
         TickFrequency   =   10
      End
      Begin ComctlLib.Slider tkColorAlpha 
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   1920
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   327682
         Max             =   255
         TickFrequency   =   10
      End
      Begin ComctlLib.Slider tkColorSaturation 
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   327682
         Max             =   240
         TickFrequency   =   10
      End
      Begin ComctlLib.Slider tkColorBrightness 
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   327682
         Max             =   240
         TickFrequency   =   10
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "&Alpha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "&Hue:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "&Saturation:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "&Brightness:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   855
      End
      Begin VB.Shape spColor 
         BackColor       =   &H000000FF&
         BackStyle       =   1  '투명하지 않음
         Height          =   495
         Left            =   5055
         Shape           =   3  '원형
         Top             =   1200
         Width           =   495
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
      Height          =   1815
      Index           =   2
      Left            =   9600
      TabIndex        =   4
      Top             =   600
      Width           =   5775
      Begin ComctlLib.Slider tkAfterGlowAlpha 
         Height          =   375
         Left            =   960
         TabIndex        =   19
         Top             =   1920
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   327682
         Max             =   255
         TickFrequency   =   10
      End
      Begin ComctlLib.Slider tkAfterGlowHue 
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   327682
         Max             =   239
         TickFrequency   =   10
      End
      Begin ComctlLib.Slider tkAfterGlowSaturation 
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   327682
         Max             =   240
         TickFrequency   =   10
      End
      Begin ComctlLib.Slider tkAfterGlowBrightness 
         Height          =   375
         Left            =   960
         TabIndex        =   15
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   327682
         Max             =   240
         TickFrequency   =   10
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "A&lpha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Shape spAfterGlow 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  '투명하지 않음
         Height          =   495
         Left            =   5055
         Shape           =   3  '원형
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "B&rightness:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "S&aturation:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "H&ue:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin ComctlLib.TabStrip tbTabStrip 
      Height          =   2205
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3889
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Color"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "After Glow"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Translucent 
      BackColor       =   &H00000000&
      BackStyle       =   1  '투명하지 않음
      BorderStyle     =   0  '투명
      Height          =   975
      Left            =   0
      Top             =   5475
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
      Left            =   6240
      TabIndex        =   8
      Top             =   3480
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

Sub OnDWMChange()
    On Error Resume Next
    Dim i%
    If IsDWMEnabled() Then
        'Me.BackColor = &H0&
        Translucent.BackColor = 0
        For i = Rect.LBound To Rect.UBound
            Rect(i).BackColor = &H404040
        Next i
        ExtendDWMFrame Me, 0, 0, 49, 0
        Me.Height = 6630
    Else
        Translucent.BackColor = &H8000000F
        Me.BackColor = &H8000000F
        For i = Rect.LBound To Rect.UBound
            Rect(i).BackColor = &H8000000F
        Next i
        Me.Height = 5895
    End If
End Sub

Private Sub chkTransparent_Click()
    SetParameters
End Sub

Private Sub cmdApply_Click()
    
End Sub

Private Sub cmdOK_Click()
    objShell.RegWrite "HKCU\Software\Microsoft\Windows\DWM\HideBlur", chkHideBlur.Value, "REG_DWORD"
    SetParameters True
    DiscardChanges = False
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    DiscardChanges = True
    Unload Me
End Sub

Private Sub Form_Load()
    Set objShell = CreateObject("WScript.Shell")
    DiscardChanges = True
    OnConpositionChanged Me
    
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
    tkColorHue.Value = hlsColor(1)
    tkColorSaturation.Value = hlsColor(3)
    tkColorBrightness.Value = hlsColor(2)
    tkColorAlpha.Value = Params.ColorAlpha
    
    ColorRGBToHLS CDec(CDec(CDec(Params.AfterGlowBlue) * &H10000) + CDec(CDec(Params.AfterGlowGreen) * &H100) + CDec(Params.AfterGlowRed)), hlsColor(1), hlsColor(2), hlsColor(3)
    tkAfterGlowHue.Value = hlsColor(1)
    tkAfterGlowSaturation.Value = hlsColor(3)
    tkAfterGlowBrightness.Value = hlsColor(2)
    tkAfterGlowAlpha.Value = Params.AfterGlowAlpha
    
    tkColorBalance.Value = Params.ColorBalance
    tkAfterGlowBalance.Value = Params.AfterGlowBalance
    tkBlurBalance.Value = Params.BlurBalance
    tkStripesIntensity.Value = Params.StripesIntensity
    If Params.Opaque Then
        chkTransparent.Value = 0
    Else
        chkTransparent.Value = 1
    End If
    On Error Resume Next
    chkHideBlur.Value = objShell.RegRead("HKCU\Software\Microsoft\Windows\DWM\HideBlur")
    Dim build As Long
    build = GetBuild()
    If build >= 7127 And build <> 7600 And build <> 7601 And build <> 8102 And build <> 8250 And build <> 8400 And build <> 9200 And build <> 9600 And build < 10240 Then
        chkHideBlur.Enabled = -1
    Else
        chkHideBlur.Enabled = 0
    End If
End Sub

Sub OnColorChange()
    spColor.BackColor = ColorHLSToRGB(tkColorHue.Value, tkColorBrightness.Value, tkColorSaturation.Value)
End Sub

Sub OnAfterGlowChange()
    spAfterGlow.BackColor = ColorHLSToRGB(tkAfterGlowHue.Value, tkAfterGlowBrightness.Value, tkAfterGlowSaturation.Value)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DiscardChanges Then DwmSetColorizationParameters OriginalParams, True
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

Private Sub tkAfterGlowBalance_Scroll()
    SetParameters
End Sub

Private Sub tkBlurBalance_Scroll()
    SetParameters
End Sub

Private Sub tkColorAlpha_Scroll()
    SetParameters
End Sub

Private Sub tkColorBalance_Scroll()
    SetParameters
End Sub

Private Sub tkColorHue_Scroll()
    OnColorChange
    SetParameters
End Sub

Private Sub tkColorSaturation_Scroll()
    OnColorChange
    SetParameters
End Sub

Private Sub tkColorBrightness_Scroll()
    OnColorChange
    SetParameters
End Sub

Private Sub tkAfterGlowHue_Scroll()
    OnAfterGlowChange
    SetParameters
End Sub

Private Sub tkAfterGlowSaturation_Scroll()
    OnAfterGlowChange
    SetParameters
End Sub

Private Sub tkAfterGlowBrightness_Scroll()
    OnAfterGlowChange
    SetParameters
End Sub

Sub SetParameters(Optional ByVal Commit As Boolean = False)
    Dim Params As DWM_COLORIZATION_PARAMS
    Dim clrColor As Long
    Dim r As Long
    Dim g As Long
    Dim b As Long
    
    clrColor = ColorHLSToRGB(tkColorHue.Value, tkColorBrightness.Value, tkColorSaturation.Value)
    ColorRefToRGB clrColor, r, g, b
    Params.ColorRed = r
    Params.ColorGreen = g
    Params.ColorBlue = b
    Params.ColorAlpha = tkColorAlpha.Value
    
    clrColor = ColorHLSToRGB(tkAfterGlowHue.Value, tkAfterGlowBrightness.Value, tkAfterGlowSaturation.Value)
    ColorRefToRGB clrColor, r, g, b
    Params.AfterGlowRed = r
    Params.AfterGlowGreen = g
    Params.AfterGlowBlue = b
    Params.AfterGlowAlpha = tkAfterGlowAlpha.Value
    
    Params.ColorBalance = tkColorBalance.Value
    Params.AfterGlowBalance = tkAfterGlowBalance.Value
    Params.BlurBalance = tkBlurBalance.Value
    Params.StripesIntensity = tkStripesIntensity.Value
    
    If chkTransparent.Value Then
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

Private Sub tkStripesIntensity_Scroll()
    SetParameters
End Sub
