Attribute VB_Name = "Game_Menu"
' ######################################
' ##
' ##  LightQ menu system
' ##

Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, _
                        lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, _
                        lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Const MAX_PATH = 260

Const INVALID_HANDLE_VALUE = -1

Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime   As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime  As FILETIME
   nFileSizeHigh    As Long
   nFileSizeLow     As Long
   dwReserved0      As Long
   dwReserved1      As Long
   cFileName        As String * MAX_PATH
   cAlternate       As String * 14
End Type

Sub MenuLoop()

  Dim Tmp      As POINTAPI
  Dim aY       As Long
  Dim SelBut   As Long
  Dim CurBut   As Long
  Dim Title    As String

  Dim Fading   As Long
  Dim FadeStep As Long
  Dim FadeTime As Long
  
  'execution is menu
  ExecutionState = EXEC_MENU
        
  'animation step
  Fading = &H7F7F7F
  FadeStep = 1
  
  'load level packs
  LoadLevelPackNames

  ' ###### MAIN LOOP ######
  
  Do
     'lock game interaction
     LockGameMouse = True
          
     'play menu song
     PlayMusic "SONG01"
       
     'set title font
     With FontEx
       .Name = "Arial Black"
       .Size = 12
       .Bold = True
       .Italic = False
       .Strikethrough = False
       .UnderLine = False
       .Charset = 0
       .Colour = vbWhite
     End With
       
     'clear out filled in objects
     With frmScreen
     
     .ed_packname = vbNullString
     .ed_password = vbNullString
     .ed_retype = vbNullString
    
     ' ###### MENU LOOP ######
     
     Do
            
            ' draw menu sprite
            BitBlt hSprite(H_BACKBUF), 0, 0, XRes, YRes, hSprite(H_MENU), 0, 0, vbSrcCopy
         
            ' is menu animation enabled ? do it
            If ExecutionState = EXEC_MENU Then
               
               If Val(.chk_settings(2).Tag) = 1 Then
                
                   ' do a slow pulsating fade on the beams
                   If TimeElapsed(FadeTime, DELAY_FADE) Then
                     Fading = Fading + (FadeStep * &H101010)
                     If Fading < &H7F7F7F Then Fading = &H7F7F7F: FadeStep = -FadeStep
                     If Fading > vbWhite Then Fading = vbWhite: FadeStep = -FadeStep
                     FadeTime = GetTickCount()
                   End If
            
                   SelBut = (ScrY - 110) \ 80
                   If SelBut < 0 Then SelBut = 0
                   If SelBut > 5 Then SelBut = 5
                   CurBut = 0
                
                   For i = 130 To 630 Step 80
                     If CurBut = SelBut Then
                        Tmp.X = 98
                        Tmp.Y = i
                        SetPen vbBlue And Fading, 2
                        MoveToEx hSprite(H_BACKBUF), Tmp.X, Tmp.Y, Tmp
                        LineTo hSprite(H_BACKBUF), XRes / 2 - 150, i
                        Tmp.X = XRes / 2 + 150
                        Tmp.Y = i
                        MoveToEx hSprite(H_BACKBUF), Tmp.X, Tmp.Y, Tmp
                        LineTo hSprite(H_BACKBUF), 741, i
                        UnsetPen
                     End If
                     CurBut = CurBut + 1
                   Next i
                
               End If
      
               ' draw menu buttons
               For i = H_PLAY To H_EXIT
                   TransparentBlt hSprite(H_BACKBUF), XRes / 2 - 150, (i - H_PLAY) * 80 + 110, 300, 40, hSprite(i), 0, 0, 300, 40, vbBlack
               Next i
               
            End If
            
            DoEvents
     
            ' show visble menu's
            .lev_sel.Visible = IIf(ExecutionState = EXEC_SELECT, True, False)
            .pack_sel.Visible = IIf(ExecutionState = EXEC_PACK, True, False)
            .mnu_settings.Visible = IIf(ExecutionState And EXEC_SETTINGS, True, False)
            .mnu_update.Visible = IIf(ExecutionState = EXEC_UPDATE, True, False)
            
            Select Case ExecutionState
            Case EXEC_SELECT
                Title = "LEVEL SELECTION"
            Case EXEC_PACK
               Title = "EDIT LEVELPACK"
            Case EXEC_SETTINGS
               Title = "SETTINGS"
            Case EXEC_MENU
               Title = "MAIN MENU"
            Case EXEC_UPDATE
               Title = "LIGHTQ UPDATE DOWNLOADER"
            End Select
            TextEx.Draw FontEx, 34, 5, 258, 560, hSprite(H_BACKBUF), Title, F_CENTER Or F_SINGLELINE Or F_VCENTER
            
            ' sync to screen
            BitBlt frmScreen.hdc, 0, 0, XRes, YRes, hSprite(H_BACKBUF), 0, 0, vbSrcCopy
  
     Loop Until ExecutionState And EXEC_GAME
     
     End With
    
     ' enable game interaction
     LockGameMouse = False
     
     ' stop music
     StopMusic
     
     With frmScreen
        ' hide level selection menu
        .lev_sel.Visible = False
        
        ' enter gameloop
        GameLoop .levPacks.List(.levPacks.ListIndex), Val(.Levels.List(.Levels.ListIndex))
        
        ' show current level preview
        .Levels_Click
     End With
          
  Loop

End Sub

Public Sub LoadLevelPackNames(Optional SetThisEditPackActive As String)

    Dim Entry   As WIN32_FIND_DATA
    Dim fHandle As Long
    Dim fName   As String
    
    With frmScreen
    
    .levPacks.Clear
    .lstEditPack.Clear
    
    ' load all levelpacks
    fHandle = FindFirstFile(App.Path & "\data\levels\*.*", Entry)
    If fHandle <> INVALID_HANDLE_VALUE Then
       Do
          fName = Left(Entry.cFileName, InStr(1, Entry.cFileName, vbNullChar) - 1)
          If Entry.dwFileAttributes And vbDirectory Then
             If Left(fName, 1) <> "." Then
                
                ' hack for DEMOMODE, don't load any other levelpacks as the standard pack
                If DEMOMODE = False Or UCase(fName) = "STANDARD" Then
                  .levPacks.AddItem UCase(fName)
                End If
                
                If UCase(fName) = "STANDARD" Then
                   .levPacks.ListIndex = .levPacks.ListCount - 1
                Else
                   .lstEditPack.AddItem UCase(fName)
                   If Len(SetThisEditPackActive) And UCase(fName) = UCase(SetThisEditPackActive) Then
                      .lstEditPack.ListIndex = .lstEditPack.ListCount - 1
                   End If
                End If
             End If
          End If
       Loop While FindNextFile(fHandle, Entry)
       FindClose fHandle
    End If

    ' no levelpacks, so set to new pack creation
    If .lstEditPack.ListCount = 0 Then
       EngineSelect = True
       .o_edit_pack(1).Value = True
       EngineSelect = False
    End If
   
    End With

End Sub

Public Sub LoadLevelNames(Packname As String)

    Dim Entry   As WIN32_FIND_DATA
    Dim fHandle As Long
    Dim fName   As String
    Dim Name    As String
    
    EngineSelect = True
  
    With frmScreen.Levels
    
    .Clear
    
    ' load all levels for the selected levelpack
    fHandle = FindFirstFile(App.Path & "\data\levels\" & Packname & "\*.lev", Entry)
    If fHandle <> INVALID_HANDLE_VALUE Then
       Do
          fName = Left(Entry.cFileName, InStr(1, Entry.cFileName, vbNullChar) - 1)
          If Entry.dwFileAttributes And Not vbDirectory Then
             Open App.Path & "\data\levels\" & Packname & "\" & fName For Input As #1
                Line Input #1, Name
                .AddItem Val(Mid(fName, InStr(1, fName, "_") + 1)) & " - " & UCase(Name)
             Close #1
          End If
       Loop While FindNextFile(fHandle, Entry)
       FindClose fHandle
    End If

    End With
    
    frmScreen.Levels.ListIndex = -1
    EngineSelect = False
  
End Sub

'
' make a small level preview for use in the SELECT LEVEL menu
'
Sub PreviewLevel(Level As Long)
   
   Dim Ani    As Animations
   Dim spGrid As String
   
   With frmScreen
     
     '[NEW 2.0.1] Set grid type
     spGrid = gGrid(.cmbGrid.ListIndex)
     
     '[FIX 2.0.1] avoid reloading a preview of the same level
     If .Preview.Tag = Str(Level) Then Exit Sub
     
     '[BUGFIX 2.0.1] set field boundaries, to avoid levels with empty inventory list
     REAL_X1 = 1
     REAL_Y1 = 1
     REAL_X2 = CL_C - 2
     REAL_Y2 = CL_R - 2
     
     LoadLevel .levPacks.List(.levPacks.ListIndex), Level
   
     ' draw all objects
     For Y = 0 To CL_R - 1
        For X = 0 To CL_C - 1
           With gMatrix(X, Y)
               If Len(.sprite) Then
                  Ani = GetAnimationInfo(.sprite)
                  DrawSprite X, Y, 0, Ani, False, H_BACKBUF2
               Else
                  Ani = GetAnimationInfo(spGrid)
                  DrawSprite X, Y, 0, Ani, False, H_BACKBUF2
               End If
           End With
        Next X
     Next Y
   
     ' draw all light beams
     For Y = 0 To CL_R - 1
        For X = 0 To CL_C - 1
           With lMatrix(X, Y)
              For i = 1 To 8
                 If .Col(i) > 0 Then
                    DrawBeam X, Y, .Col(i), (i), .partial(i), H_BACKBUF2
                 End If
              Next i
          End With
        Next X
     Next Y
   
     .Preview.Cls
     '[NEW 2.0.2] resizing using StretchBlt in XP now looks awsome!
     SetStretchBltMode .Preview.hdc, STRETCHMODE
     StretchBlt .Preview.hdc, 0, 0, .Preview.ScaleWidth, .Preview.ScaleHeight, hSprite(H_BACKBUF2), 0, 0, XRes, YRes - SP_H, vbSrcCopy
     'set "this level is already previewed" FLAG
     .Preview.Tag = Str(Level)
    
   End With
      
End Sub
