Attribute VB_Name = "Util"
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Long, Margin As MARGINS) As Long
Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef pfEnabled As Long) As Long
Declare Sub DwmGetColorizationParameters Lib "dwmapi.dll" Alias "#127" (ByRef Parameters As DWM_COLORIZATION_PARAMS)
Declare Sub DwmSetColorizationParameters Lib "dwmapi.dll" Alias "#131" (ByRef Parameters As DWM_COLORIZATION_PARAMS, ByVal BoolArg As Boolean)
Const WM_APP As Long = 32768
Const WM_DWMCOMPOSITIONCHANGED As Long = &H31E
Private Const GWL_WNDPROC = (-4)
Private procOld As Long
Private DWMhWnd As Long
Private DWMForm As Form

Declare Function CallWindowProc Lib "USER32.DLL" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "USER32.DLL" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function ColorRGBToHLS Lib "shlwapi.dll" ( _
    ByVal clrRGB As Long, _
    ByRef wHue As Integer, _
    ByRef wLuminance As Integer, _
    ByRef wSaturation As Integer) As Long
    
Declare Function ColorHLSToRGB Lib "shlwapi.dll" ( _
    ByVal wHue As Integer, _
    ByVal wLuminance As Integer, _
    ByVal wSaturation As Integer) As Long

Type DWM_COLORIZATION_PARAMS
    ColorBlue  As Byte
    ColorGreen As Byte
    ColorRed   As Byte
    ColorAlpha As Byte
    
    AfterGlowBlue  As Byte
    AfterGlowGreen As Byte
    AfterGlowRed   As Byte
    AfterGlowAlpha As Byte
    
    ColorBalance     As Long '5-55
    AfterGlowBalance As Long
    BlurBalance      As Long
    
    StripesIntensity As Long
    
    Opaque As Boolean
End Type

Type MARGINS
    cxLeftWidth    As Long
    cxRightWidth   As Long
    cyTopHeight    As Long
    cyBottomHeight As Long
End Type

Type OSVERSIONINFO
    OSVSize         As Long
    dwVerMajor      As Long
    dwVerMinor      As Long
    dwBuildNumber   As Long
    PlatformID      As Long
    szCSDVersion    As String * 128
End Type

Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

Function IsDWMEnabled() As Boolean
    Dim DwmEnabled As Long
    DwmEnabled = 0
    DwmIsCompositionEnabled DwmEnabled
    If DwmEnabled > 0 Then
        IsDWMEnabled = True
    Else
        IsDWMEnabled = False
    End If
End Function

Sub ExtendDWMFrame(ByRef frm As Form, Top As Long, Right As Long, Bottom As Long, Left As Long)
    Dim Margin As MARGINS
    Margin.cxLeftWidth = Left
    Margin.cxRightWidth = Right
    Margin.cyTopHeight = Top
    Margin.cyBottomHeight = Bottom
    DwmExtendFrameIntoClientArea frm.hWnd, Margin
End Sub

Sub OnConpositionChanged(ByRef frm As Form)
    DWMhWnd = frm.hWnd
    Set DWMForm = frm
    procOld = SetWindowLong(DWMhWnd, GWL_WNDPROC, AddressOf SubWndProc)
End Sub

Sub UnsubclassWindow(ByVal hWnd As Long)
    procOld = SetWindowLong(hWnd, GWL_WNDPROC, procOld)
End Sub

Private Function SubWndProc( _
        ByVal hWnd As Long, _
        ByVal iMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

    If hWnd = DWMhWnd Then
        If iMsg = WM_DWMCOMPOSITIONCHANGED Then
            DWMForm.OnDWMChange
            SubWndProc = True
            Exit Function
        End If
    End If

    SubWndProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)
End Function

'https://www.vbforums.com/showthread.php?889448-RESOLVED-VB6-how-get-RGB-from-ARGB-color
Sub ColorRefToRGB(ByVal Color As Long, ByRef r As Long, g As Long, b As Long)
    r = Color And &HFF
    
    g = ((Color \ &H100) And &HFF)
    g = (Color And &HFF00&) \ &H100&
    
    b = ((Color \ &H10000) And &HFF)
    b = (Color And &HFF0000) \ &H10000
End Sub

Public Function GetVersion() As Double
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        GetVersion = CDbl(osv.dwVerMajor) + CDbl(osv.dwVerMinor) * 0.1
    Else
        GetVersion = 0#
    End If
End Function

Public Function GetBuild() As Long
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        GetBuild = osv.dwBuildNumber
    Else
        GetBuild = 0#
    End If
End Function
