VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ######################################
' ##
' ##  LightQ Buffered Text Class
' ##

Public Enum eFormatFlags
  F_Bottom = DT_BOTTOM
  F_CENTER = DT_CENTER
  F_left = DT_LEFT
  F_CALCRECT = DT_CALCRECT
  F_WORDBREAK = DT_WORDBREAK
  F_VCENTER = DT_VCENTER
  F_TOP = DT_TOP
  F_TABSTOP = DT_TABSTOP
  F_SINGLELINE = DT_SINGLELINE
  F_RIGHT = DT_RIGHT
  F_NOCLIP = DT_NOCLIP
  F_INTERNAL = DT_INTERNAL
  F_EXTERNALLEADING = DT_EXTERNALLEADING
  F_EXPANDTABS = DT_EXPANDTABS
  F_CHARSTREAM = DT_CHARSTREAM
  F_NOPREFIX = DT_NOPREFIX
End Enum

Private m_tTextRect As RECT
Private m_oStdFontEx As clsStdFontEx
Private m_lhDc As Long
Private m_sText As String
Private m_lAllignmentFlags As eFormatFlags

Public Property Set StdFontEx(vDta As clsStdFontEx)
  Set m_oStdFontEx = vDta
End Property
Public Property Get StdFontEx() As clsStdFontEx
  If Not (m_oStdFontEx Is Nothing) Then
    Set StdFontEx = m_oStdFontEx
    End If
End Property

Public Property Let RectBottom(vDta As Long)
  m_tTextRect.Bottom = vDta
End Property
Public Property Let RectTop(vDta As Long)
  m_tTextRect.Top = vDta
End Property
Public Property Let RectLeft(vDta As Long)
  m_tTextRect.Left = vDta
End Property
Public Property Let RectRight(vDta As Long)
  m_tTextRect.Right = vDta
End Property

Public Property Let hdc(vDta As Long)
  m_lhDc = vDta
End Property
Public Property Let Text(vDta As String)
  m_sText = vDta
End Property
Public Property Let AllignmentFlags(ByVal vDta As eFormatFlags)
  m_lAllignmentFlags = vDta
End Property

Public Sub Draw(Optional oStdFontEx As clsStdFontEx = Nothing, _
                Optional RectBottom As Long = -1, _
                Optional RectTop As Long = -1, _
                Optional RectLeft As Long = -1, _
                Optional RectRight As Long = -1, _
                Optional hdc As Long = -1, _
                Optional Text As String = vbNullString, _
                Optional ByRef AllignmentFlags As eFormatFlags = -1)
                
  Dim LF As LOGFONT
  Dim sFont As String
  Dim iChar As Integer
  Dim m_lFont As Long
  Dim crOldType As Long
  
  '//Check to see whether the user passed any arguements in the .Draw procedure _
     if no arguements were passed then use the Class.* propertys
  If Not (oStdFontEx Is Nothing) Then
    Set m_oStdFontEx = oStdFontEx
    End If
    
  If Not (RectBottom = -1) Then
    m_tTextRect.Bottom = RectBottom
    End If
  If Not (RectTop = -1) Then
    m_tTextRect.Top = RectTop
    End If
  If Not (RectLeft = -1) Then
    m_tTextRect.Left = RectLeft
    End If
  If Not (RectRight = -1) Then
    m_tTextRect.Right = RectRight
    End If
    
  If Not (hdc = -1) Then
    m_lhDc = hdc
    End If
  If Not (Text = vbNullString) Then
    m_sText = Text
    End If
  If Not (AllignmentFlags = -1) Then
    m_lAllignmentFlags = AllignmentFlags
    End If
  
  '//Convert StdFontEx to the LogFont Structure so that it can be applied to the hDc
  With LF
    sFont = m_oStdFontEx.Name
    ' There is a quicker way involving StrConv and CopyMemory, but
    ' this is simpler!:
    For iChar = 1 To Len(sFont)
      .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
      Next iChar
    ' Based on the Win32SDK documentation:
    .lfHeight = -MulDiv((m_oStdFontEx.Size), (GetDeviceCaps(m_lhDc, LOGPIXELSY)), 72)
    .lfItalic = m_oStdFontEx.Italic
    If (m_oStdFontEx.Bold) Then
      .lfWeight = FW_BOLD
      Else
        .lfWeight = FW_NORMAL
      End If
    .lfUnderline = m_oStdFontEx.UnderLine
    .lfStrikeOut = m_oStdFontEx.Strikethrough
    .lfCharSet = m_oStdFontEx.Charset
  End With
  
  '//Apply the Font Formatting to the hDc
  m_lFont = CreateFontIndirect(LF)
  DeleteObject SelectObject(m_lhDc, m_lFont)
  DeleteObject m_lFont

  '//Apply the Text Colour Formatting to the hDc
  SetBkMode m_lhDc, TRANSPARENT
  SetTextColor m_lhDc, m_oStdFontEx.Colour
  
  '//OutPut the Text
  DrawText m_lhDc, m_sText, Len(m_sText), m_tTextRect, m_lAllignmentFlags
  
End Sub
