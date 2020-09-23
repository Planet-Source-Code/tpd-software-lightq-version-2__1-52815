VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStdFontEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ######################################
' ##
' ##  LightQ Font Class
' ##

Private m_oStdFont As StdFont
Private m_lColour As Long
Private m_sName As String
Private m_lSize As Long
Private m_bItalic As Boolean
Private m_bBold As Boolean
Private m_bUnderline As Boolean
Private m_bStrikethrough As Boolean
Private m_lCharset As Long

Public Property Set StdFont(vDta As StdFont)
  Set m_oStdFont = vDta
  With m_oStdFont
    m_sName = .Name
    m_lSize = .Size
    m_bItalic = .Italic
    m_bBold = .Bold
    m_bUnderline = .UnderLine
    m_bStrikethrough = .Strikethrough
    m_lCharset = .Charset
  End With
End Property
Public Property Get StdFont() As StdFont
  If Not (m_oStdFont Is Nothing) Then
    Set StdFont = m_oStdFont
    Else
    'Raise Error - Not Set
  End If
End Property

Public Property Let Colour(vDta As Long)
  m_lColour = vDta
End Property
Public Property Let Name(vDta As String)
  m_sName = vDta
End Property
Public Property Let Size(vDta As Long)
  m_lSize = vDta
End Property
Public Property Let Italic(vDta As Boolean)
  m_bItalic = vDta
End Property
Public Property Let Bold(vDta As Boolean)
  m_bBold = vDta
End Property
Public Property Let UnderLine(vDta As Boolean)
  m_bUnderline = vDta
End Property
Public Property Let Strikethrough(vDta As Boolean)
  m_bStrikethrough = vDta
End Property
Public Property Let Charset(vDta As Long)
  m_lCharset = vDta
End Property

Public Property Get Colour() As Long
  If m_lColour <> 0 Then
  Colour = m_lColour
  Else
    'Raise Error - Not Set
    End If
End Property
Public Property Get Name() As String
  If m_lColour <> 0 Then
    Name = m_sName
    Else
    'Raise Error - Not Set
    End If
End Property
Public Property Get Size() As Long
  If m_lSize <> 0 Then
    Size = m_lSize
    Else
    'Raise Error - Not Set
    End If
End Property
Public Property Get Italic() As Boolean
  Italic = m_bItalic
End Property
Public Property Get Bold() As Boolean
  Bold = m_bBold
End Property
Public Property Get UnderLine() As Boolean
  UnderLine = m_bUnderline
End Property
Public Property Get Strikethrough() As Boolean
  Strikethrough = m_bStrikethrough
End Property
Public Property Get Charset() As Long
'TO DO  - _
        Change it so that it returns the Correct Charset in the enviroment that it _
        is operating in _
 _
        You could be a lazy bastard and simply loop thru the screen.font object _
        looking for the FontName and getting its .Charaset property but i know _
        that you all are real programmers and wouldnt stoop soo low - and in _
        all likelyhood cause a dramattic slow down on Slow Machines with Lots of _
        Fonts installed
        
  If m_lCharset <> 0 Then
    Charset = m_lCharset
    Else
    Charset = 0 'Better hope that its a Standard Windows character
    End If
End Property

