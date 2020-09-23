Attribute VB_Name = "Game_Misc"
' ######################################
' ##
' ##  LightQ miscaleneous functions
' ##
  
'Get/Set WindowLong Constants (only those used)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

'SetWindowPos Constants (only those used)
Private Const SWP_FRAMECHANGED = &H20 'The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

'Dialog Styles (also present in the GWL_STYLE area)
Private Const DS_ABSALIGN As Long = &H1
Private Const DS_SYSMODAL As Long = &H2
Private Const DS_3DLOOK As Long = &H4
Private Const DS_FIXEDSYS As Long = &H8
Private Const DS_NOFAILCREATE As Long = &H10
Private Const DS_LOCALEDIT As Long = &H20      'Edit items get Local storage.
Private Const DS_SETFONT As Long = &H40        'User specified font for Dlg controls
Private Const DS_MODALFRAME As Long = &H80     'Can be combined with WS_CAPTION
Private Const DS_NOIDLEMSG As Long = &H100     'WM_ENTERIDLE message will not be sent
Private Const DS_SETFOREGROUND As Long = &H200 'not in win3.1
Private Const DS_CONTROL As Long = &H400
Private Const DS_CENTER As Long = &H800
Private Const DS_CENTERMOUSE As Long = &H1000
Private Const DS_CONTEXTHELP As Long = &H2000

'Window Styles (GWL_STYLE area)
Private Const WS_OVERLAPPED As Long = &H0
Private Const WS_POPUP As Long = &H80000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_MINIMIZE As Long = &H20000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_DISABLED As Long = &H8000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_CLIPCHILDREN As Long = &H2000000
Private Const WS_MAXIMIZE As Long = &H1000000
Private Const WS_CAPTION As Long = &HC00000 'WS_BORDER | WS_DLGFRAME
Private Const WS_BORDER As Long = &H800000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_VSCROLL As Long = &H200000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_GROUP As Long = &H20000
Private Const WS_TABSTOP As Long = &H10000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_TILED As Long = WS_OVERLAPPED
Private Const WS_ICONIC As Long = WS_MINIMIZE
Private Const WS_SIZEBOX As Long = WS_THICKFRAME
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_EX_NOPARENTNOTIFY As Long = &H4
Private Const WS_EX_TOPMOST As Long = &H8
Private Const WS_EX_ACCEPTFILES As Long = &H10
Private Const WS_EX_TRANSPARENT As Long = &H20
Private Const WS_EX_MDICHILD As Long = &H40
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_WINDOWEDGE As Long = &H100
Private Const WS_EX_CLIENTEDGE As Long = &H200
Private Const WS_EX_CONTEXTHELP As Long = &H400
Private Const WS_EX_RIGHT As Long = &H1000
Private Const WS_EX_LEFT As Long = &H0
Private Const WS_EX_RTLREADING As Long = &H2000
Private Const WS_EX_LTRREADING As Long = &H0
Private Const WS_EX_LEFTSCROLLBAR As Long = &H4000
Private Const WS_EX_RIGHTSCROLLBAR As Long = &H0
Private Const WS_EX_CONTROLPARENT As Long = &H10000
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_APPWINDOW As Long = &H40000

Public Const DT_BOTTOM = &H8
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_CALCRECT = &H400
Public Const DT_WORDBREAK = &H10
Public Const DT_VCENTER = &H4
Public Const DT_TOP = &H0
Public Const DT_TABSTOP = &H80
Public Const DT_SINGLELINE = &H20
Public Const DT_RIGHT = &H2
Public Const DT_NOCLIP = &H100
Public Const DT_INTERNAL = &H1000
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_EXPANDTABS = &H40
Public Const DT_CHARSTREAM = 4
Public Const DT_NOPREFIX = &H800

Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const DM_DISPLAYFREQUENCY = &H400000
Private Const CDS_FORCE As Long = &H80000000
Private Const HORZRES As Long = 8
Private Const VERTRES As Long = 10
Private Const BITSPIXEL As Long = 12
Private Const VREFRESH As Long = 116

Public Const SW_SHOWDEFAULT = 10

Public Const MAX_PATH = 260

Public Const LF_FACESIZE = 32
  
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Const TRANSPARENT = 3

Public Const PS_DASH = 1                    '  -------
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_DOT = 2                     '  .......
Public Const PS_INSIDEFRAME = 6
Public Const PS_SOLID = 0
Public Const PS_NULL = 5

Public Const EXEC_END = 0
Public Const EXEC_MENU = 1
Public Const EXEC_GAME = 2
Public Const EXEC_SELECT = 4
Public Const EXEC_EDIT = 8
Public Const EXEC_SETTINGS = 16
Public Const EXEC_PACK = 32
Public Const EXEC_UPDATE = 64
   
' indeces for hSprite()
Public Const H_BMPBUF = 0
Public Const H_BMPBUF2 = 1
Public Const H_BACKBUF = 2
Public Const H_BACKBUF2 = 3
Public Const H_SPRITES = 4
Public Const H_MINES = 5
Public Const H_CTRLBAR = 6
Public Const H_PLAY = 7
Public Const H_EDIT = 8
Public Const H_SETTINGS = 9
Public Const H_UPDATE = 10
Public Const H_VISIT = 11
Public Const H_EXIT = 12
Public Const H_MENU = 13
Public Const H_DISABLED = 14
Public Const H_EDITBAR = 15
Public Const H_POINT = 16
Public Const H_MOVE = 17
Public Const H_TPDLOGO = 18
Public Const H_GAMELOGO = 19
Public Const H_SMALLBUTTON = 20

' width of the beams
Public Const BEAM_WIDTH = 3

' INI filename
Public Const INI_FILE As String = "LightQ.ini"

' WinXP / XML Class for WinXP style objects
Private Const ICC_INTERNET_CLASSES = &H800

Private Type INITCOMMONCONTROLSEX_TYPE
    dwSize As Long
    dwICC As Long
End Type

' device mode structure
Private Type DevMode
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

' 2d coords structure
Type POINTAPI
   X As Long
   Y As Long
End Type

' used to hold screen startingmode
Private Type ScreenInfo
   Width   As Long
   Height  As Long
   Bpp     As Long
   Refresh As Long
End Type

' font structure
Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

' square region structure
Type RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
End Type

' used to detect lostfocus on screen (for FULLSCREEN mode)
Declare Function GetActiveWindow Lib "User32" () As Integer

' get windows temp directory
Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' used for "visit" to show the webpage of the author
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' used to enable WinXP controls
Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As INITCOMMONCONTROLSEX_TYPE) As Long

' direct memory copying
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' font API's
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function DrawText Lib "User32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

' form api's
Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' display api's
Declare Function EnumDisplaySettings Lib "User32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, _
    ByVal modeIndex As Long, lpDevMode As Any) As Boolean
Declare Function ChangeDisplaySettings Lib "User32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

' square memory copy routine
Declare Function BitBlt Lib "gdi32" ( _
      ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
      ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' square memory copy routine with maskable color
Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, _
      ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, _
      ByVal nWidthDest As Long, ByVal hHeightDest As Long, _
      ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, _
      ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
      ByVal crTransparent As Long) As Boolean

Public Const STRETCHMODE = vbPaletteModeNone
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long

Declare Function StretchBlt Lib "gdi32" (ByVal hdcDest As Long, _
      ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, _
      ByVal nWidthDest As Long, ByVal hHeightDest As Long, _
      ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, _
      ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
      ByVal dwRop As Long) As Boolean

'code timer
Declare Function GetTickCount Lib "kernel32" () As Long

'creating buffers / loading sprites
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long

'loading sprites
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'cleanup
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

' lines
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

' line / color
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal fnDrawMode As Integer) As Integer

' active window handle ( returned by GetActiveWinow() )
Public ActiveHwnd  As Long

' handles for all used graphic resources
Public hSprite(20) As Long

' where is excution?
Public ExecutionState As Long

' handle for the pen
Private hPen As Long

' holds the old resolution
Private OldScreen As ScreenInfo

' font classes
Public TextEx As New clsTextEx
Public FontEx As New clsStdFontEx
Public HttpEx As New clsHTTP

' registered user
Public Reg_User  As String
Public Reg_Code  As String

' sytem fonts
Public m_aFonts() As String

' cheats
Public strCheat   As String
' 0=Freeze mines,1=Barrels doesn't explode,2=All level cheat
Public arCheats(2) As Boolean

Sub Intro()
    
    Dim tc As Long
    Dim x1 As Single
    Dim y1 As Single
    Dim w  As Single
    Dim h  As Single
    
    ' erase any garbage from our backbuffer
    BitBlt hSprite(H_BACKBUF), 0, 0, XRes, YRes, 0, 0, 0, vbBlackness
    
    x1 = XRes \ 2 - 25
    w = 50
    y1 = YRes \ 2 - 5
    h = 10
    
    PlaySound "LOGOSHOW"
    
    Do
      '[NEW 2.0.2] resizing using StretchBlt in XP now looks awsome!
      SetStretchBltMode hSprite(H_BACKBUF), STRETCHMODE
      StretchBlt hSprite(H_BACKBUF), x1, y1, w, h, hSprite(H_TPDLOGO), 0, 0, 500, 100, vbSrcCopy
      y1 = y1 - (1 / 5)
      x1 = x1 - 1
      h = h + (2 / 5)
      w = w + 2
      BitBlt frmScreen.hdc, 0, 0, XRes, YRes, hSprite(H_BACKBUF), 0, 0, vbSrcCopy
    DoEvents
    Loop Until w >= 500 And h >= 100
    
    tc = GetTickCount()
    Do
      DoEvents
    Loop Until TimeElapsed(tc, 2000)
    
    With FontEx
      .Name = "Arial"
      .Size = 18
      .Bold = False
      .Italic = False
      .Strikethrough = False
      .UnderLine = False
      .Charset = 0
      .Colour = vbWhite
    End With
       
    tc = GetTickCount()
    BitBlt hSprite(H_BACKBUF), 0, 0, XRes, YRes, 0, 0, 0, vbBlackness
    Do
      TextEx.Draw FontEx, YRes, 0, 0, XRes, hSprite(H_BACKBUF), "PROUDLY PRESENTS", F_CENTER Or F_SINGLELINE Or F_VCENTER
      BitBlt frmScreen.hdc, 0, 0, XRes, YRes, hSprite(H_BACKBUF), 0, 0, vbSrcCopy
    DoEvents
    Loop Until TimeElapsed(tc, 2000)
    
    PlaySound "COVERSHOW"
    
    tc = GetTickCount()
    Do
      BitBlt hSprite(H_BACKBUF), XRes \ 2 - 200, YRes \ 2 - 200, 400, 400, hSprite(H_GAMELOGO), 0, 0, vbSrcCopy
      BitBlt frmScreen.hdc, 0, 0, XRes, YRes, hSprite(H_BACKBUF), 0, 0, vbSrcCopy
    DoEvents
    Loop Until TimeElapsed(tc, 4000)
    
    PlaySound "EXPLOSION"
    
    With frmScreen
      For i = 255 To 0 Step -5
        .BackColor = RGB(i, i, i)
        .Refresh
        tc = GetTickCount()
        Do: DoEvents: Loop Until TimeElapsed(tc, 20)
      Next i
    End With
    
End Sub

Function LoadGraphicDC(sFileName As String) As Long
   Dim tDC As Long
   ' load a graphic into memory
   On Error GoTo LoadFailure
   'create the DC address compatible with
   'the DC of the screen
   tDC = CreateCompatibleDC(GetDC(0))
   'load the graphic file into the DC...
   SelectObject tDC, LoadPicture(App.Path & "\data\graphics\" & sFileName)
   'return the address of the file
   LoadGraphicDC = tDC
   Exit Function
LoadFailure:
   LoadGraphicDC = 0
End Function

Sub InitScreen(Optional OnlySwitchResolution As Boolean = False)
   
   ' initialize the screen for use with backbuffering
   ' set screen resolution if fullscreen mode is enabled
   ' change form settings
   
   Dim tc As Long
   
   On Error GoTo InitScreenError
   
   Static CurrentScreenMode As Byte
   Static Init              As Boolean
   
   With frmScreen
   
     ' set form properties
     .ScaleMode = vbPixels
     .AutoRedraw = False
     .ClipControls = True
          
     If Not OnlySwitchResolution Then
        ' create surface
        hSprite(H_BACKBUF) = CreateCompatibleDC(GetDC(0))
        hSprite(H_BACKBUF2) = CreateCompatibleDC(GetDC(0))
        hSprite(H_BMPBUF) = CreateCompatibleBitmap(GetDC(0), XRes, YRes)
        hSprite(H_BMPBUF2) = CreateCompatibleBitmap(GetDC(0), XRes, YRes - SP_H)
        SelectObject hSprite(H_BACKBUF), hSprite(H_BMPBUF)
        SelectObject hSprite(H_BACKBUF2), hSprite(H_BMPBUF2)
     End If
     
     If CurrentScreenMode <> .cmbScreen.ListIndex Or Not Init Then
     
        If CurrentScreenMode = 1 Then
           RestoreScreen
        End If
        
        .Width = XRes * Screen.TwipsPerPixelX + (.Width - (.ScaleWidth * Screen.TwipsPerPixelX))
        .Height = YRes * Screen.TwipsPerPixelY + (.Height - (.ScaleHeight * Screen.TwipsPerPixelY))
        
        Select Case .cmbScreen.ListIndex
        Case 0
           ChangeFormBorder frmScreen, vbFixedSingle, True, True, False, True, True, False
           .Move Screen.Width \ 2 - .Width \ 2, Screen.Height \ 2 - .Height \ 2
        Case 1
           ' switch to fullscreen with max refresh rate
           tc = GetTickCount()
           SaveCurrentScreen
           SetScreenWithMaxRefresh XRes, YRes, Depth
           ChangeFormBorder frmScreen, vbBSNone, True, False, False, False, True, False
           .Move 0, 0
           Do
             DoEvents
           Loop Until TimeElapsed(tc, 3000)
        End Select
      
        .Width = XRes * Screen.TwipsPerPixelX + (.Width - (.ScaleWidth * Screen.TwipsPerPixelX))
        .Height = YRes * Screen.TwipsPerPixelY + (.Height - (.ScaleHeight * Screen.TwipsPerPixelY))
        
        .ZOrder 0
        .Refresh
        
        CurrentScreenMode = .cmbScreen.ListIndex
        
     End If
     
   End With
   
   Init = True
   
   Exit Sub

InitScreenError:
ShowError "Screen initialisation failed"
   
End Sub

Sub KillScreen()
   ' free up all used resources
   Set TextEx = Nothing
   Set FontEx = Nothing
   Set objInfoMap = Nothing
   UnsetPen
   DeleteObject hSprite(H_BMPBUF)
   DeleteObject hSprite(H_BMPBUF2)
   For i = 1 To UBound(hSprite)
      DeleteDC hSprite(i)
   Next i
   ' when in full screen, switch back to old resolution
   If frmScreen.cmbScreen.ListIndex = 1 Then
      RestoreScreen
   End If
End Sub

Sub SetPen(color As Long, Optional Width As Long = BEAM_WIDTH, Optional BACKBUFFER As Long = H_BACKBUF)
    ' create a pen
    hPen = CreatePen(PS_SOLID, Width, color)
    SelectObject hSprite(BACKBUFFER), hPen
End Sub

Sub UnsetPen()
    ' remove the pen
    If hPen Then DeleteObject hPen
    hPen = 0
End Sub

Sub LoadFonts()

   ReDim m_aFonts(0 To Screen.FontCount - 1)
   
   For i = 0 To Screen.FontCount - 1
     m_aFonts(i) = Screen.Fonts(i)
   Next i
 
End Sub

Sub SaveCurrentScreen()
   Dim hdc As Long

   hdc = GetDC(0)
   
   ' get the screen settings
   With OldScreen
     .Width = GetDeviceCaps(hdc, HORZRES)
     .Height = GetDeviceCaps(hdc, VERTRES)
     .Bpp = GetDeviceCaps(hdc, BITSPIXEL)
     .Refresh = GetDeviceCaps(hdc, VREFRESH)
   End With
   
   DeleteDC hdc
   
End Sub

Sub RestoreScreen()
   Dim DM       As DevMode
   
   DM.dmSize = Len(DM)
   DM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
   With OldScreen
     DM.dmPelsWidth = .Width
     DM.dmPelsHeight = .Height
     DM.dmBitsPerPel = .Bpp
     DM.dmDisplayFrequency = .Refresh
   End With

   If ChangeDisplaySettings(DM, CDS_FORCE) <> 0 Then
      ShowError "Error setting display mode."
   End If
   
End Sub

Sub SetScreenWithMaxRefresh(Width As Long, Height As Long, Bpp As Long)

   Dim dCount   As Long
   Dim cRefresh As Long
   Dim iX       As Long
   Dim uX       As Long
   ReDim DM(0) As DevMode
   
   cRefresh = 0
   
   DM(0).dmSize = Len(DM(0))
   DM(0).dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
    
   ' loop trough all display modes
   Do While EnumDisplaySettings(0, dCount, DM(0)) > 0
      dCount = dCount + 1
   Loop
    
   ReDim DM(0 To dCount)
      
   For iX = 0 To dCount
      
      EnumDisplaySettings 0, iX, DM(iX)
      
      With DM(iX)
        If .dmPelsWidth = Width And .dmPelsHeight = Height And .dmBitsPerPel = Bpp Then
           If .dmDisplayFrequency > cRefresh Then
              cRefresh = .dmDisplayFrequency
              uX = iX
           End If
        End If
      End With
          
   Next iX
    
   If cRefresh > 0 Then
      If ChangeDisplaySettings(DM(uX), CDS_FORCE) <> 0 Then
         ShowError "Error setting display mode."
      End If
   End If

End Sub

Public Sub ChangeFormBorder(frmForm As Form, _
    ByVal eNewBorder As FormBorderStyleConstants, _
    Optional ByVal bClipControls As Boolean = True, _
    Optional ByVal bControlBox As Boolean = True, _
    Optional ByVal bMaxButton As Boolean = True, _
    Optional ByVal bMinButton As Boolean = True, _
    Optional ByVal bShowInTaskBar As Boolean = True, _
    Optional ByVal bWhatsThisButton As Boolean = False)
    Dim lRet As Long
    Dim lStyleFlags As Long
    Dim lStyleExFlags As Long
    
    'Initialize our flags
    lStyleFlags = 0
    lStyleExFlags = 0
    
    'If we want ClipControls then add that
    'flag and change the form property
    If bClipControls Then
        lStyleFlags = lStyleFlags Or WS_CLIPCHILDREN
        frmForm.ClipControls = True
    Else
        frmForm.ClipControls = False
    End If
    
    'If we want the control box then add the
    'flag (property is read-only)
    If bControlBox Then lStyleFlags = lStyleFlags Or WS_SYSMENU
    
    'If we want the max button then add the
    'flag (property is read-only)
    If bMaxButton Then lStyleFlags = lStyleFlags Or WS_MAXIMIZEBOX
    
    'If we want the min button then add the
    'flag (property is read-only)
    If bMinButton Then lStyleFlags = lStyleFlags Or WS_MINIMIZEBOX
    
    'If we want the form to show in taskbar
    'then add the flag (property is read-only
    If bShowInTaskBar Then lStyleExFlags = lStyleExFlags Or WS_EX_APPWINDOW
    
    'If we want the what's this button then
    'add the flag (property is read-only)
    If bWhatsThisButton Then lStyleExFlags = lStyleExFlags Or WS_EX_CONTEXTHELP
    
    'If the form is an MDI Child form then a
    'add the flag (Don't want to screw up the form)
    If frmForm.MDIChild Then lStyleExFlags = lStyleExFlags Or WS_EX_MDICHILD
        
    'Now we need to set the flags for the
    'borrder we are changing to
    Select Case eNewBorder
      Case vbBSNone
       lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS)
       'No change to extended style flags.
      Case vbFixedSingle
       lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CAPTION)
       lStyleExFlags = lStyleExFlags Or WS_EX_WINDOWEDGE
      Case vbSizable
       lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CAPTION Or WS_THICKFRAME)
       lStyleExFlags = lStyleExFlags Or WS_EX_WINDOWEDGE
      Case vbFixedDialog
       lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CAPTION Or DS_MODALFRAME)
       lStyleExFlags = lStyleExFlags Or (WS_EX_WINDOWEDGE Or WS_EX_DLGMODALFRAME)
      Case vbFixedToolWindow
       lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CAPTION)
       lStyleExFlags = lStyleExFlags Or (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW)
      Case vbSizableToolWindow
       lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CAPTION Or WS_THICKFRAME)
       lStyleExFlags = lStyleExFlags Or (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW)
    End Select
   
    'Change our styles
    lRet = SetWindowLong(frmForm.hwnd, GWL_STYLE, lStyleFlags)
    lRet = SetWindowLong(frmForm.hwnd, GWL_EXSTYLE, lStyleExFlags)
    
    'Signal that the frame has changed
    lRet = SetWindowPos(frmForm.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED)
    
    'Make that we've changed the border in
    'the form's property
    frmForm.BorderStyle = eNewBorder
End Sub
        
Sub SetPointer(Name As String)
    frmScreen.MouseIcon = LoadPicture(App.Path & "\data\graphics\" & Name & ".cur")
End Sub

Function GetWindowsTempFolder() As String

    Dim Buffer     As String
    Dim BufferLen  As Long
    
    Buffer = Space(MAX_PATH)
    BufferLen = Len(Buffer)
    
    If GetTempPath(BufferLen, Buffer) Then
       GetWindowsTempFolder = Left(Buffer, InStr(1, Buffer, Chr(0), vbBinaryCompare) - 1)
    End If
    
End Function

Sub FormOnTop(TheForm As Form, bTopMost As Boolean)

    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
    Select Case bTopMost
    Case True
        Placement = HWND_TOPMOST
    Case False
        Placement = HWND_NOTOPMOST
    End Select
    
    SetWindowPos TheForm.hwnd, Placement, 0, 0, 0, 0, wFlags
End Sub
        
Sub InitWinXPControls()

   Dim ComCtls As INITCOMMONCONTROLSEX_TYPE  ' identifies the control to register
   Dim Retval  As Long                       ' generic return value
   
   With ComCtls
    .dwSize = Len(ComCtls)
    .dwICC = ICC_INTERNET_CLASSES
   End With

   Retval = InitCommonControlsEx(ComCtls)

End Sub

Sub SaveSettings()
    
    Dim INIPath As String
    
    INIPath = App.Path & "\" & INI_FILE
    
    With frmScreen
        
         WriteINI INIPath, "GENERAL", "Name", App.ProductName
         WriteINI INIPath, "GENERAL", "Version", App.Major & "." & App.Minor & "." & App.Revision
         WriteINI INIPath, "GENERAL", "Copyright", App.LegalCopyright
                 
         If AudioPresent Then
            WriteINI INIPath, "SOUND", "SoundVolume", .sVol.Value
            WriteINI INIPath, "SOUND", "SoundEnabled", .chk_settings(0).Tag
            WriteINI INIPath, "SOUND", "MusicVolume", .mVol.Value
            WriteINI INIPath, "SOUND", "MusicEnabled", .chk_settings(1).Tag
         End If
         
         WriteINI INIPath, "SCREEN", "ScreenMode", .cmbScreen.ListIndex
         WriteINI INIPath, "SCREEN", "GridStyle", .cmbGrid.ListIndex
         WriteINI INIPath, "SCREEN", "MenuAnimation", .chk_settings(2).Tag
         
         WriteINI INIPath, "REGISTER", "UserName", Reg_User
         WriteINI INIPath, "REGISTER", "RegCode", Reg_Code
         
    End With

End Sub
        
Sub LoadSettings()

    Dim INIPath As String
    
    On Error Resume Next
    
    INIPath = App.Path & "\" & INI_FILE
    
    EngineSelect = True
    
    With frmScreen

        If AudioPresent Then
           .sVol = Val(ReadINI(INIPath, "SOUND", "SoundVolume", "-1000"))
           .chk_settings(0).Tag = Val(ReadINI(INIPath, "SOUND", "SoundEnabled", "1"))
           .mVol = Val(ReadINI(INIPath, "SOUND", "MusicVolume", "-1000"))
           .chk_settings(1).Tag = Val(ReadINI(INIPath, "SOUND", "MusicEnabled", "1"))
           
           SoundEnabled = CBool(Val(.chk_settings(0).Tag))
           MusicEnabled = CBool(Val(.chk_settings(1).Tag))
        Else
           .chk_settings(0).Tag = "2"
           .chk_settings_Show 0
           .chk_settings(1).Tag = "2"
           .chk_settings_Show 1
        End If
        
        .cmbScreen.ListIndex = Val(ReadINI(INIPath, "SCREEN", "ScreenMode", "0"))
        .cmbGrid.ListIndex = Val(ReadINI(INIPath, "SCREEN", "GridStyle", "2"))
        .chk_settings(2).Tag = Val(ReadINI(INIPath, "SCREEN", "MenuAnimation", "1"))
        
        Reg_User = ReadINI(INIPath, "REGISTER", "UserName", vbNullString)
        Reg_Code = ReadINI(INIPath, "REGISTER", "RegCode", vbNullString)
        
    End With
        
    EngineSelect = False
            
End Sub
        
Function TimeElapsed(Current As Long, Delay As Long, Optional CurrentTickCount As Long) As Boolean

    Dim tc As Long
    
    If CurrentTickCount = 0 Then
       CurrentTickCount = GetTickCount()
    End If
    
    If CurrentTickCount < 0 Then
       TimeElapsed = IIf(CurrentTickCount + Current <= Delay, True, False)
    Else
       TimeElapsed = IIf(CurrentTickCount - Current >= Delay, True, False)
    End If

End Function

Sub ShowError(Message As String, Optional CustomTitle As String = "ERROR")
   
    With frmScreen
      .Cls
      .ErrMes = Message
      .lblErrTitle = CustomTitle
      .eLayer.ZOrder 0
      .eLayer.Visible = True
      
      LockMenuMouse = True
      LockGameMouse = True
        
      Do
        DoEvents
      Loop Until .eLayer.Visible = False

      LockMenuMouse = False
      LockGameMouse = False
      
    End With
    
End Sub

Function DoInput(Title As String, Optional Default As String = vbNullString) As String
  
    With frmScreen
      .Cls
      .iLayer.ZOrder 0
      .iLayer.Visible = True
      .lbl_input = Title
      .txt_input = Default
      .txt_input.SetFocus
      .txt_input.SelStart = 0
      .txt_input.SelLength = Len(.txt_input)
      
      LockMenuMouse = True
      LockGameMouse = True
      
      Do
        DoEvents
      Loop Until .iLayer.Visible = False

      DoInput = Trim(.txt_input)

      LockMenuMouse = False
      LockGameMouse = False
      
    End With

End Function

' only OK and YESNO are supported!
Function DoQuestion(Message As String, Optional CustomTitle As String = "LIGHTQ", Optional DialogType As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult
  
    With frmScreen
      .Cls
      .qLayer.ZOrder 0
      .qLayer.Visible = True
      .QuestionMess = Message
      .eQuestion.Tag = vbNullString
      .lblQuestionTitle = CustomTitle
      .frm_question(DialogType).ZOrder 0
      .img_question(DialogType).ZOrder 0
      
      LockMenuMouse = True
      LockGameMouse = True
      
      Do
        DoEvents
      Loop Until .qLayer.Visible = False

      DoQuestion = .eQuestion.Tag

      LockMenuMouse = False
      LockGameMouse = False
      
    End With

End Function

Function GetRegisterCode(Name As String) As String

   Const base As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
   
   Dim magic1 As Long
   Dim magic2 As Long
   Dim s      As String
   
   magic1 = 5789
   magic2 = 9437
   
   ' password is 20 chars long
   For i = 1 To 20
      For p = 1 To Len(Text1)
         magic1 = magic1 + Asc(Mid(Text1, p, 1)) Xor magic2
      Next p
      magic2 = magic2 Xor 128 And magic1
      s = s & Mid(base, (magic1 Xor magic2) Mod Len(base) + 1, 1)
   Next i
    
   GetRegisterCode = s
   
End Function
