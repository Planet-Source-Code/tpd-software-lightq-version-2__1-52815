Attribute VB_Name = "Game_INI"
' ######################################
' ##
' ##  LightQ INI handling
' ##

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Private iniErr As Long

Function ReadINI(sFile As String, sSection As String, sKeyword As String, Optional sDefault As String = vbNullString) As String

    Dim Arg As String
    Dim Ret As Long

    Arg = String(255, vbNullChar)
    
    Ret = GetPrivateProfileString(sSection, ByVal sKeyword, vbNullString, Arg, Len(Arg), sFile)
    
    If Ret Then
       ReadINI = Left(Arg, Ret)
    Else
       ReadINI = sDefault
    End If
    
    iniErr = Ret

End Function

Sub WriteINI(sFile As String, sSection As String, sKeyword As String, sValue As String)

    Dim Ret As Long

    Ret = WritePrivateProfileString(sSection, sKeyword, sValue, sFile)
    
    iniErr = Ret

End Sub

Function LastINIError() As Long
    
    LastINIError = IIf(iniErr = 0, iniErr, 1)
    
End Function

