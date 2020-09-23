Attribute VB_Name = "Game_Core"
' ######################################
' ##
' ##  LightQ Game engine
' ##

' switch between DEMO and full version
' !---------------------------------------!
Public DEMOMODE As Boolean
' !---------------------------------------!

Public Const GAME_TITLE As String = "LIGHTQ %v %a - TPD Software 2003"

Public Const UPDATE_SERVER As String = "http://217.121.35.5/Updates/LightQ"

' X/Y resolution
Public Const XRes  As Long = 800
Public Const YRes  As Long = 600
Public Const Depth As Long = 32

' sprites width and height
Public Const SP_W As Long = 32
Public Const SP_H As Long = 32

' field columns and rows
Public Const CL_C As Long = 25
Public Const CL_R As Long = 18

' fade delay time in tickcounts
Public Const DELAY_FADE = 50

' animation start frame for coins
Public Const COIN_START_FRAME = 1
' animation start frames for lasers
Public Const LASER_START_FRAME = 0

' predefined objects used in gameloop
Public Const SP_BARREL = "BA_*"
Public Const SP_EXPLOSION = "EX_*"
Public Const SP_MINE = "MN_*"

' playfield boundaries
Public REAL_X1 As Long
Public REAL_X2 As Long
Public REAL_Y1 As Long
Public REAL_Y2 As Long

Public X    As Long
Public Y    As Long
Public cx   As Long
Public cy   As Long

Public DragObject As String

' lock/unlock user interaction for gamefield
Public LockGameMouse As Boolean

' lock/unlock user interaction for menu's
Public LockMenuMouse As Boolean

' doing stable check when mines in field
Public StableTime    As Long
   
' beam properties
Private Enum BeamMethods
    BEAM_END = 1
    BEAM_CONTINUE = 2
End Enum

' explosion properties
Private Enum Explosion
    EX_NONE = 0
    EX_GO = 1
    EX_WAIT = 2
End Enum

' X/Y displacement
Private Type Direction
    x_disp As Long
    y_disp As Long
End Type

' gamematrix structure
Public Type gMatrixInfo
    sprite    As String
    is_rot    As Boolean
    is_mov    As Boolean
    explode   As Explosion
    cur_frame As Byte
    Time      As Long
End Type

' lightmatrix structure
Public Type gLightInfo
    Dir(1 To 8)     As Byte
    Col(1 To 8)     As Long
    partial(1 To 8) As Boolean
End Type

Public Type EditMatrix
    is_rot     As Boolean
    is_mov     As Boolean
End Type

Private Type gInventory
    sprite    As String
    is_rot    As Boolean
End Type

' animation structure
Public Type Animations
    x_start    As Byte
    y_start    As Byte
    frames     As Byte
    beammethod As Byte
    Delay      As Long
    color      As Long
End Type

' gObject -> MultiDir
Private Type MultiDir
    Dir()      As Byte
End Type

' gObject -> MultiCol
Private Type MultiCol
    Col()      As Byte
End Type

' object structure
Private Type gObject
   fromdir() As Byte
   todir()   As MultiDir
   tocol()   As MultiCol
   onedir    As Boolean
End Type

' mine structure
Private Type gMines
   X         As Long
   Y         As Long
   ax        As Integer
   aY        As Integer
   cur_frame As Byte
   Time      As Long
End Type

' objects
Public Mirror()           As gObject
Public DoubleMirror()     As gObject
Public SlitDoubleMirror() As gObject
Public Refracter()        As gObject
Public Prism()            As gObject
Public Oneway()           As gObject
Public Filter()           As gObject
Public Polarizer()        As gObject
Public Splitter()         As gObject
Public AngledSplitter()   As gObject
Public Converter()        As gObject

' map name
Public MapName            As String

' mines
Public Mines()            As gMines

' game and light array's
Public gMatrix()          As gMatrixInfo
Public eMatrix()          As EditMatrix
Public lMatrix()          As gLightInfo
Public Inventory()        As gInventory

' grid styles [ 0=NONE,1=GR_L,2=GR_D ]
Public gGrid(2)           As String

' sprite x/y offset and behaviour settings
Public objInfoMap         As Collection

' level completed flag
Public Completed          As Boolean

' color array
Public aColors(0 To 7)    As Long

Sub GameLoop(Optional Packname As String, Optional Level As Long)
   
   On Error GoTo GameError
   
   Static sLevel         As Long
   Static sPack          As String
   
   Dim Ani               As Animations
   Dim AniDone           As Long
   Dim Animate           As Boolean
   Dim FadeStep          As Long
   Dim Fading            As Long
   Dim FadeTime          As Long
   Dim oName             As String
   Dim oCol              As Byte
   Dim oDir              As Byte
   Dim BlackCoinsStopped As Long
   Dim NumCoinsSpin      As Long
   Dim NumCoins          As Long
   Dim cTick             As Long
   Dim DoStableCheck     As Long
   Dim CompletedTextTop  As Long
   Dim NoLight           As Boolean
   Dim spGrid            As String
   
   '[BUGFIX 2.0.1] set field boundaries, to avoid levels with empty inventory list
   REAL_X1 = 1
   REAL_Y1 = 1
   REAL_X2 = CL_C - 2
   REAL_Y2 = CL_R - 2
   
   'load level / init field / do calculations on field
   If ExecutionState = EXEC_GAME Then
       If Level > 0 Then
          sLevel = Level
       End If
       If Len(Packname) > 0 Then
          sPack = Packname
       End If
       LoadLevel sPack, sLevel
       
       ' start a song if not already playing
       If Not MusicPlaying Then
          PlayMusic "SONG" & Right("00" & Trim(Str(Int(Rnd * 19) + 2)), 2)
       End If
       
   Else
       ' initialise a new level to edit
       CurrentLevel = 0
       ClearMatrices
       DrawBorder
       CalcMines
       CalcBeams
       MapName = "UNTITLED"
   End If
   
   'initialize some vars
   Fading = &H7F7F7F
   FadeStep = 1
   LockGameMouse = False
   CurX = REAL_X1
   CurY = REAL_Y1
   CompletedTextTop = SP_H
   StableTime = 0
   Completed = False
   '[NEW 2.0.1] Set grid type
   spGrid = gGrid(frmScreen.cmbGrid.ListIndex)
   
   'clear the level index indicator flag
   frmScreen.Levels.Tag = vbNullString
   
   'count number of coins in game
   NumCoins = CountCoins()
   
   'clear all sound buffers
   FreeSoundBuffers
       
   'clear screenbuffer
   BitBlt hSprite(H_BACKBUF), 0, 0, XRes, YRes, 0, 0, 0, vbBlackness
       
   ' ################# GAMELOOP ######################
                     
   Do
   
       ' ################# OBJECTS ###################################
              
       NumCoinsSpin = 0
       BlackCoinsStopped = 0
       cTick = GetTickCount()
       
       For Y = 0 To CL_R - 1
       
           For X = 0 To CL_C - 1
   
               With gMatrix(X, Y)
                  
                  If Len(.sprite) Then
                  
                     Animate = True
                     
                     Ani = GetAnimationInfo(.sprite)
   
                     GetObject .sprite, oName, oDir, oCol
   
                     ' ################## OBJECT CONTROL #########################
                     
                     Select Case oName
                     Case "LS", "LR"
                        ' lasers turned off (after explosion) ?
                        If NoLight Then
                           Animate = False
                        End If
                     
                        ' always stop animation when back on starting frame
                        If .cur_frame <> LASER_START_FRAME Then
                           Animate = True
                        End If
                        
                     ' always turn off animation of mines
                     ' because this is done in MoveMines()
                     Case "MN"
                        Animate = False
                     
                     ' beam hits barrel?
                     Case "BA"
                     
                        ' only do explosions when in game mode, not in edit mode
                        ' [NEW 2.0.1] CheatCode: LightDoesntHurt
                        If ExecutionState = EXEC_GAME And Not arCheats(1) Then
                        
                            c = 0
                            For i = 1 To 8
                                If lMatrix(X, Y).Col(i) > 0 Then c = 1: Exit For
                            Next i
                     
                            If c = 1 Or .explode = EX_GO Then
                          
                               ' get rid of the dragobject if present
                               If Len(DragObject) Then
                                  mUp 0, 0, (ScrX), (ScrY)
                               End If
                           
                               ' yes, lock game user interaction
                               LockGameMouse = True
                               
                               ' allow menu interaction
                               LockMenuMouse = False
                           
                               ' if we were in stable checking progress
                               ' abort this
                               StableTime = -1
                               
                               'play explosion sound
                               PlaySound "EXPLOSION"
                           
                               ' remove all present mines
                               ReDim Mines(0)
                                                      
                               If Not NoLight Then
                                 ' recalculate beams
                                 CalcBeams True
                                 
                                 ' turn lasers off
                                 NoLight = True
                               End If
                           
                               ' check for surrounding barrels
                               For cy = Y - 1 To Y + 1
                                    For cx = X - 1 To X + 1
                                        If InField(cx, cy, False) Then
                                            With gMatrix(cx, cy)
                                                If (cx = X And cy = Y) Then
                                                    .sprite = SP_EXPLOSION
                                                Else
                                                    If .sprite = SP_BARREL Then
                                                       .explode = EX_WAIT
                                                    Else
                                                       .sprite = SP_EXPLOSION
                                                    End If
                                                End If
                                                .cur_frame = 0
                                           End With
                                       End If
                                   Next cx
                               Next cy
                        
                           End If
                           
                        End If
                                          
                     ' check beams over colored coin
                     Case "CO"
                        
                        Animate = False
                        
                        c = 0
                        With lMatrix(X, Y)
                          For i = 1 To 8
                            If .Col(i) > 0 Then
                               If .Col(i) = aColors(oDir) Then
                                  c = c + 1
                               Else
                                  c = c - 1
                               End If
                            End If
                          Next i
                        End With
                        
                        If c = 2 Then
                           Animate = True
                           NumCoinsSpin = NumCoinsSpin + 1
                        End If
                     
                        ' only stop animation when back on starting frame
                        If .cur_frame <> COIN_START_FRAME Then
                           Animate = True
                        End If
                     
                     ' check beams over black coin
                     Case "CB"
                        
                        For i = 1 To 8
                          If lMatrix(X, Y).Col(i) > 0 Then
                             Animate = False
                             BlackCoinsStopped = BlackCoinsStopped + 1
                             Exit For
                          End If
                        Next i
                        
                        ' always stop animation when back on starting frame
                        If .cur_frame <> COIN_START_FRAME Then
                           Animate = True
                        End If
                     
                     End Select
                     
                     ' ############### DRAW AND ANIMATION CONTROL ################
                      
                     If TimeElapsed(.Time, Ani.Delay, cTick) Or Ani.frames = 1 Then
   
                       ' draw the (animated) sprite into the playing field
                       AniDone = DrawSprite(X, Y, .cur_frame, Ani, Animate)
                       
                       ' control the chain reaction explosion of barrels
                       If .sprite = SP_EXPLOSION Then
                          Select Case AniDone
                          Case 25 To 50
                             ExplodeWaitingObjects X, Y
                          Case 100
                             .sprite = vbNullString
                             .explode = EX_NONE
                          End Select
                       End If
                       
                       .Time = cTick
                       
                     End If
                  
                     ' ###########################################################
                     
                  Else
                      
                     ' empty cell, set the background on it
                     Ani = GetAnimationInfo(spGrid)
                     DrawSprite X, Y, 0, Ani, False
                  
                  End If
   
               End With
               
           Next X
         
       Next Y
   
       If NumCoins = NumCoinsSpin And BlackCoinsStopped = 0 And MouseDown = False Then
         
          If UBound(Mines) > 0 And DoStableCheck <> -1 Then
             
             If Not LockGameMouse Then
                
                ' turn music off (volume low)
                SetMusicVolume True, -10000
                
                ' play stablecheck sound
                PlaySound "STABLECHECK"
                
                'when mines in game, do a 5 second stable check
                DoStableCheck = GetTickCount()
                StableTime = 5
                
                ' we are going to show the stable countdown timer
                ' into the controlbar area, so trun off this menu
                LockMenuMouse = True
                
             End If
          
          Else
            
             If Not Completed And StableTime = 0 Then
                
                'remove any present mines
                ReDim Mines(0)
            
                ' stop music
                StopMusic
                
                ' turn music back on
                SetMusicVolume
                
                'play completed sound
                PlaySound "COMPLETED"
                
                ' enable the controlbar
                LockMenuMouse = False
                
                'set level completed flag
                Completed = True
          
                '--->>> NEW: enhance level for player
                If UCase(Packname) = "STANDARD" And Not arCheats(2) Then
                   ' update player level for standard pack
                   
                   DoUpdatePlayerLevel frmScreen.lbl_player
                Else
                   ' playing a "FREE" pack so just increment level if possible
                   With frmScreen
                     If .Levels.ListIndex < .Levels.ListCount - 1 Then
                        .Levels.ListIndex = .Levels.ListIndex + 1
                     Else
                        .Levels.Tag = "ALL_DONE"
                     End If
                   End With
                End If
          
             End If
          
          End If
          
          'no further ineraction in game possible
          LockGameMouse = True
           
       End If
       
       ' ################## MINES ##################################
       
       ' move the mines if there are any
       MoveMines
      
       ' ################# BEAMS ###################################
                        
       ' do a slow pulsating fade on the beams
       If TimeElapsed(FadeTime, DELAY_FADE) Then
          Fading = Fading + (FadeStep * &H101010)
          If Fading < &H7F7F7F Then Fading = &H7F7F7F: FadeStep = -FadeStep
          If Fading > vbWhite Then Fading = vbWhite: FadeStep = -FadeStep
          FadeTime = GetTickCount()
       End If
       
       ' draw all light beams
       For Y = 0 To CL_R - 1
         For X = 0 To CL_C - 1
            With lMatrix(X, Y)
               For i = 1 To 8
                  If .Col(i) > 0 Then
                     DrawBeam X, Y, .Col(i) And Fading, (i), .partial(i)
                  End If
               Next i
           End With
         Next X
       Next Y
       
       ' ############## SCREEN CONTROL #############################
                     
       ' draw map name text
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
       TextEx.Draw FontEx, SP_H, 0, (REAL_X1 + 1) * SP_W, REAL_X2 * SP_W, hSprite(H_BACKBUF), MapName, F_CENTER Or F_SINGLELINE Or F_VCENTER
              
       If ExecutionState = EXEC_GAME Then
          
          ' we draw the system stable countdown timer instead of the controlbar
          If StableTime > 0 Then
             
             ' countdown each second
             If TimeElapsed(DoStableCheck, 1000) Then
                
                'play timer sound
                PlaySound "TIMER"
                
                StableTime = StableTime - 1
                DoStableCheck = GetTickCount()
                
                If StableTime = 0 Then
                   ' stable check performed and ok, win level!
                   DoStableCheck = -1
                End If
             End If
             
             ' clear the controlbar area
             BitBlt hSprite(H_BACKBUF), 0, YRes - SP_H, XRes, SP_H, 0, 0, 0, vbBlackness
             
             ' draw the timer in the controlbar area
             TextEx.Draw FontEx, YRes, YRes - SP_H, 0, XRes, hSprite(H_BACKBUF), "-=[ SYSTEM STABLE IN " & StableTime & " SECONDS ]=-", F_TOP Or F_CENTER Or F_SINGLELINE Or F_VCENTER
             
          Else
             ' draw controlbar
             BitBlt hSprite(H_BACKBUF), 0, YRes - SP_H, XRes, SP_H, hSprite(H_CTRLBAR), 0, 0, vbSrcCopy
          End If
          
       Else
          ' draw editbar
          BitBlt hSprite(H_BACKBUF), 0, YRes - SP_H, XRes, SP_H, hSprite(H_EDITBAR), 0, 0, vbSrcCopy
          
          If YSelected > -1 Then
             Ani = GetAnimationInfo(EditObjects(SelectedObjectIndex()).Object)
          Else
             With Ani
                .x_start = 0
                .y_start = 0
             End With
          End If
          
          ' draw selected object
          BitBlt hSprite(H_BACKBUF), 235, YRes - SP_H, SP_W, SP_H, hSprite(H_SPRITES), Ani.x_start * SP_W, Ani.y_start * SP_H, vbSrcCopy
          
          ' Update: Avoiding the "Array Is Locked" Error
          ' ########## DETECT MATRIX SHIFTING ##########
          If DoShiftOnMatrix > 0 Then
             ShiftMatrix DoShiftOnMatrix
             DoShiftOnMatrix = 0
          End If
          
       End If
              
       ' draw completed text (after completing the level)
       If Completed Then
          
          FontEx.Size = 16
          TextEx.Draw FontEx, CompletedTextTop + SP_H, CompletedTextTop, (REAL_X1 + 1) * SP_W, REAL_X2 * SP_W, hSprite(H_BACKBUF), "CONGRATULATIONS, SYSTEM COMPLETED !", F_CENTER Or F_SINGLELINE Or F_VCENTER
         
          If CompletedTextTop < YRes \ 2 - SP_H \ 2 Then
             CompletedTextTop = CompletedTextTop + 4
          End If
       Else
          
          ' disable the continue button as long as the level isn't solved
          ' [BUGFIX 2.0.1] and we are not in stablecheck mode
          If StableTime < 1 And ExecutionState = EXEC_GAME Then
             TransparentBlt hSprite(H_BACKBUF), 449, 572, 48, 24, hSprite(H_DISABLED), 0, 0, 48, 24, vbBlack
          End If
       
       End If
       
       'draw selection box around selected object
       DrawSelectionBox
              
       'draw dragobject
       If Len(DragObject) Then
          Ani = GetAnimationInfo(DragObject)
          With Ani
             TransparentBlt hSprite(H_BACKBUF), ScrX - (SP_W \ 2), ScrY - (SP_H \ 2), SP_W, SP_H, hSprite(H_SPRITES), .x_start * SP_W, .y_start * SP_H, SP_W, SP_H, vbBlack
          End With
       End If
       
       'copy backbuffer to screen
       BitBlt frmScreen.hdc, 0, 0, XRes, YRes, hSprite(H_BACKBUF), 0, 0, vbSrcCopy
   
   DoEvents
   Loop While ExecutionState And EXEC_GAME
   
   'clear the game buffers
   ClearMatrices
       
   Exit Sub
   
   ' ################ END GAMELOOP ######################

GameError:
ShowError "Error occured in Game_Core::GameLoop" & vbCrLf & "#Error: " & Err & vbCrLf & "Description: " & Err.Description

End Sub

Function DrawSprite(cellx As Long, celly As Long, curframe As Byte, Ani As Animations, DoAnimation As Boolean, Optional BACKBUFFER As Long = H_BACKBUF) As Byte

   Dim map_x As Long
   Dim map_y As Long
   
   With Ani
   
     ' get map x/y coordinates
     map_x = SP_W * (.x_start + curframe)
     map_y = SP_H * .y_start
   
     ' draw the sprite to the backbuffer
     BitBlt hSprite(BACKBUFFER), cellx * SP_W, celly * SP_H, SP_W, SP_H, hSprite(H_SPRITES), map_x, map_y, vbSrcCopy

     ' animate the sprite
     If .frames > 1 And DoAnimation Then
        curframe = curframe + 1
        DrawSprite = (curframe / .frames * 100)
        If curframe >= .frames Then
           curframe = 0
        End If
     End If
     
   End With

End Function

Function GetAnimationInfo(ByVal Obj As String) As Animations

   Dim iX         As Long
   Dim Nm         As String
   Dim cAniVar    As Variant
   
   Nm = Left(Obj, 2)
   iX = InStr(1, Obj, ":")
   If iX > 0 And Nm = "LS" Or Nm = "LR" Then
      Obj = Left(Obj, iX - 1)
   End If
  
   cAniVar = objInfoMap.Item(IIf(Len(Obj), Obj, "EMPTY"))
      
   With GetAnimationInfo
      .x_start = cAniVar(0)
      .y_start = cAniVar(1)
      .frames = cAniVar(2)
      .Delay = cAniVar(3)
      .color = cAniVar(4)
      .beammethod = cAniVar(5)
   End With
      
End Function

Sub CalcBeams(Optional NoLight As Boolean = False)

   Dim oName  As String
   Dim oDir   As Byte
   Dim oCol   As Byte
   Dim ld     As Direction
   Dim DirIn  As Byte
   Dim DirOut As Byte

   On Error GoTo CalcError

   ' clear light matrix
   ClearMatrices False, True, False
   
   ' lasers are switched off
   ' (ie after explosion)
   If NoLight Then Exit Sub
   
   RecursionLevel = 0
   
   For cy = 0 To CL_R - 1
       
      For cx = 0 To CL_C - 1
   
          With gMatrix(cx, cy)
 
              ' read object properties
              GetObject .sprite, oName, oDir, oCol
           
              'is it a lightsource?
              If oName = "LS" Or oName = "LR" Then
              
                 ' get beam x/y displacement
                 ld = GetLightDir(oDir)
                 
                 ' draw start beam on lightsource
                 With lMatrix(cx, cy)
                    .Col(oDir) = .Col(oDir) Or aColors(oCol)
                    .partial(oDir) = True
                 End With
                
                 ' make beam system for current lightsource
                 FollowBeam cx + ld.x_disp, cy + ld.y_disp, SwapInOut(oDir), aColors(oCol), ld
                            
              End If
           
          End With
          
      Next cx
      
   Next cy
   
   Exit Sub

CalcError:
ShowError "Error occured in Game_Core::CalcBeams" & vbCrLf & "#Error: " & Err & vbCrLf & "Description: " & Err.Description

End Sub

Sub FollowBeam(ByVal ctx As Long, ByVal cty As Long, InitDir As Byte, oCol As Long, ld As Direction)

  Dim oName      As String
  Dim oDir       As Byte
  Dim iCol       As Byte
  Dim nCol       As Long
  Dim Ani        As Animations
  Dim oDirIn     As Byte
  ReDim DirIn(0) As Variant
  Dim di         As Byte
  Dim objDir     As Integer
  
  On Error GoTo CalcError
  
  ' set initial direction
  DirIn(0) = InitDir
  
  ' follow this beam
  Do
     
      With gMatrix(ctx, cty)
      
         ' get object animation info
         Ani = GetAnimationInfo(.sprite)
            
         ' object is a beamblocker
         ' beam ends here
         If Ani.beammethod = BEAM_END Then Exit Sub
             
         ' cell contains an object and object does not let beam's trough?
         If Len(.sprite) And Ani.beammethod <> BEAM_CONTINUE Then
                        
              ' get object properties
              GetObject .sprite, oName, oDir, iCol
         
              ' save our current direction
              oDirIn = DirIn(0)
            
              ' save our current color
              nCol = oCol
            
              ' clear our current direction
              ReDim DirIn(0)
         
              ' select object and adjust direction and/or color
              Select Case oName
                Case "MR": DirIn() = In2Out(Mirror(oDir), oDirIn)
                Case "RF": DirIn() = In2Out(Refracter(oDir), oDirIn)
                Case "MD": DirIn() = In2Out(DoubleMirror(oDir), oDirIn)
                Case "MS": DirIn() = In2Out(SlitDoubleMirror(oDir), oDirIn)
                Case "SP": DirIn() = In2Out(Splitter(oDir), oDirIn)
                Case "AS": DirIn() = In2Out(AngledSplitter(oDir), oDirIn)
                Case "OW": DirIn() = In2Out(Oneway(oDir), oDirIn)
                Case "PR": DirIn() = In2Out(Prism(oDir), oDirIn)
      
                Case "FT": DirIn() = In2Out(Filter(oDirIn), oDirIn):          nCol = nCol And aColors(oDir)
                Case "PO": DirIn() = In2Out(Polarizer(oDir), oDirIn):         nCol = nCol And aColors(iCol)
                Case "CC": DirIn() = In2Out(Converter(oDir), oDirIn, objDir): nCol = RolBits(nCol, objDir)
              End Select
            
              ' follow each reflacted light beam
              For i = 0 To UBound(DirIn) Step 2
            
                  With lMatrix(ctx, cty)
                 
                      ' draw end beam on object
                      di = oDirIn
                      .Col(di) = .Col(di) Or oCol
                      .partial(di) = True
                    
                      ' object obstructs beam at this state
                      ' or does not reflact at all
                      ' so this beam ends here
                      If UBound(DirIn) = 0 Then Exit Sub
                    
                      ' change color if needed
                      If DirIn(i + 1) > 0 Then
                         ' our current color is black indicating
                         ' multibeam object does not reflact this
                         ' color at this state
                         ' so exit
                         If oCol = 0 Then Exit Sub
                         
                         ' apply new color
                         nCol = oCol And aColors(DirIn(i + 1))
                         
                      End If
                    
                      ' draw start beam on object
                      di = DirIn(i)
                      .Col(di) = .Col(di) Or nCol
                      .partial(di) = True
                
                  End With
            
                  ' get new x/y displacements
                  ld = GetLightDir((DirIn(i)))
            
                  ' follow this beam recursively
                  FollowBeam ctx + ld.x_disp, cty + ld.y_disp, SwapInOut((DirIn(i))), nCol, ld
            
              Next i
            
              ' beam is ready
              Exit Sub
         
         Else
            
            ' cell is empty so apply beam to whole cell
            With lMatrix(ctx, cty)
               ' do it for the current direction,
               di = DirIn(0)
               .Col(di) = .Col(di) Or oCol
               .partial(di) = False
               ' and the reverse direction
               di = SwapInOut(di)
               .Col(di) = .Col(di) Or oCol
               .partial(di) = False
            End With
         
         End If
      
      End With
    
      ' jump to next cell
      ctx = ctx + ld.x_disp
      cty = cty + ld.y_disp
  
  Loop
  
  Exit Sub

CalcError:
ShowError "Error occured in Game_Core::FollowBeam" & vbCrLf & "#Error: " & Err & vbCrLf & "Description: " & Err.Description
  
End Sub

Sub CalcMines()
  Dim NumMines As Long

  Randomize Timer

  NumMines = 0
  
  ReDim Mines(NumMines)

  ' each mine gets it's own random direction
  For cy = REAL_Y1 To REAL_Y2
    For cx = REAL_X1 To REAL_X2
       With gMatrix(cx, cy)
          If .sprite = "MN_*" Then
             NumMines = NumMines + 1
             ReDim Preserve Mines(NumMines)
             With Mines(NumMines)
                ' set mine properties
                .X = cx * SP_W
                .Y = cy * SP_H
                .cur_frame = 0
                .Time = 0
                .ax = IIf(Rnd > 0.5, 3, -3)
                .aY = IIf(Rnd > 0.5, 3, -3)
             End With
             ' remove the mine from the gamematrix
             ' as this was for positional purposes only
             .sprite = vbNullString
          End If
       End With
    Next cx
  Next cy
  
End Sub

Sub MoveMines()

  Dim dx      As Single
  Dim dy      As Single
  Dim Bounce  As Boolean
  Dim Exp     As Boolean
  Dim BounceH As Boolean
  Dim BounceV As Boolean
  Dim Ani     As Animations
  Dim reRnd   As Long
  
  Randomize GetTickCount()
  
  Ani = GetAnimationInfo(SP_MINE)
  
  For i = 1 To UBound(Mines)
    
     With Mines(i)
        
        ' move the mine
        ' [NEW 2.0.1] CheatCode: FreezeMines
        If Not arCheats(0) Then
          .X = .X + .ax
          .Y = .Y + .aY
        End If
         
        ' get mine position
        dx = .X / SP_W
        dy = .Y / SP_H
        
        ' test for vertical collision
        cx = dx + IIf(.ax > 0, 0.49, -0.49)
        cy = dy
        GoSub CheckCollision: BounceH = Bounce
        If BounceH Then .ax = -.ax
        If Exp Then gMatrix(cx, cy).explode = EX_GO
                
        ' test for horizontal collision
        cx = dx
        cy = dy + IIf(.aY > 0, 0.49, -0.49)
        GoSub CheckCollision: BounceV = Bounce
        If BounceV Then .aY = -.aY
        If Exp Then gMatrix(cx, cy).explode = EX_GO
        
        ' draw the mine
        TransparentBlt hSprite(H_BACKBUF), .X, .Y, SP_W, SP_H, hSprite(H_MINES), .cur_frame * SP_W, 0, SP_W, SP_H, vbBlack
            
        ' animate the mine
        If TimeElapsed(.Time, Ani.Delay) Then
          .cur_frame = .cur_frame + 1
          If .cur_frame >= Ani.frames Then
             .cur_frame = 0
          End If
          .Time = GetTickCount()
        End If
           
    End With
   
  Next i
  
  Exit Sub
  
CheckCollision:

  ' initially no bounce and no explosion
  Bounce = False
  Exp = False

  With gMatrix(cx, cy)
  
    ' does mine hit a barrel?
    If .sprite = SP_BARREL Then
       Exp = True
       Return
    End If
  
    ' does mine hit other objects?
    If Len(.sprite) Then
       Bounce = True
       Return
    End If
    
    ' does mine hit a beam?
    For j = 1 To 8
       If lMatrix(cx, cy).Col(j) > 0 Then
          Bounce = True
          Return
       End If
    Next j
  
  End With
  
Return

End Sub

Sub DrawBeam(cellx As Long, celly As Long, Col As Long, Dir As Byte, partial As Boolean, Optional BACKBUFFER As Long = H_BACKBUF, Optional Width As Long = BEAM_WIDTH)

   Dim Tmp As POINTAPI
   Dim M1  As Long
   Dim M2  As Long
   Dim M3  As Long
   Dim M4  As Long
   Dim x1  As Long
   Dim y1  As Long
   Dim X2  As Long
   Dim Y2  As Long
   Dim YY  As Long
   
   Select Case Dir
   ' Horizontal beam
   Case 3, 7
     M1 = IIf(Dir = 3 And partial, SP_W \ 2, 0)
     M2 = IIf(Dir = 7 And partial, SP_W \ 2, SP_W)
     Tmp.X = cellx * SP_W + M1
     Tmp.Y = celly * SP_H + (SP_H \ 2)
     X2 = cellx * SP_W + M2
     Y2 = Tmp.Y
   ' Vertical beam
   Case 1, 5
     M1 = IIf(Dir = 5 And partial, SP_H \ 2, 0)
     M2 = IIf(Dir = 1 And partial, SP_H \ 2, SP_H)
     Tmp.X = cellx * SP_W + (SP_W \ 2)
     Tmp.Y = celly * SP_H + M1
     X2 = Tmp.X
     Y2 = celly * SP_H + M2
   'Top left to Bottom right beam
   Case 4, 8
     M1 = IIf(Dir = 4 And partial, SP_W \ 2, 0)
     M2 = IIf(Dir = 4 And partial, SP_H \ 2, 0)
     M3 = IIf(Dir = 8 And partial, SP_W \ 2, SP_W)
     M4 = IIf(Dir = 8 And partial, SP_H \ 2, SP_H)
     Tmp.X = cellx * SP_W + M1
     Tmp.Y = celly * SP_H + M2
     X2 = cellx * SP_W + M3
     Y2 = celly * SP_H + M4
   'Bottom left to top right beam
   Case 2, 6
     M1 = IIf(Dir = 2 And partial, SP_W \ 2, 0)
     M2 = IIf(Dir = 2 And partial, SP_H \ 2, SP_H)
     M3 = IIf(Dir = 6 And partial, SP_W \ 2, SP_W)
     M4 = IIf(Dir = 6 And partial, SP_H \ 2, 0)
     Tmp.X = cellx * SP_W + M1
     Tmp.Y = celly * SP_H + M2
     X2 = cellx * SP_W + M3
     Y2 = celly * SP_H + M4
   End Select
   
  ' SetROP2 hSprite(BACKBUFFER), vbMergePen
   SetPen Col, Width, BACKBUFFER
   MoveToEx hSprite(BACKBUFFER), Tmp.X, Tmp.Y, Tmp
   LineTo hSprite(BACKBUFFER), X2, Y2
   UnsetPen
  ' SetROP2 hSprite(BACKBUFFER), vbCopyPen
   
   
End Sub

Sub DrawSelectionBox()
   
   ' no selection when clicked outside field
   If InField(CurX, CurY) = False Then
      Exit Sub
   End If
   
   ' are we in game mode?
   If ExecutionState = EXEC_GAME Then
   
      With gMatrix(CurX, CurY)
   
         ' no selection if object is not movable and not rotatable
         If .is_mov = False And .is_rot = False Then
            Exit Sub
         End If
   
      End With
   
      '  no selection when ineraction disabled
      If LockGameMouse Then
         Exit Sub
      End If
   
   End If
   
   Dim Tmp As POINTAPI
   Dim x1  As Long
   Dim X2  As Long
   Dim y1  As Long
   Dim Y2  As Long
         
   x1 = CurX * SP_W
   X2 = x1 + SP_W
   y1 = CurY * SP_H
   Y2 = y1 + SP_H
   
   ' draw the selection square
   Tmp.X = x1
   Tmp.Y = y1
   MoveToEx hSprite(H_BACKBUF), x1, y1, Tmp
   SetPen vbYellow, 2
   LineTo hSprite(H_BACKBUF), x1, Y2
   LineTo hSprite(H_BACKBUF), X2, Y2
   LineTo hSprite(H_BACKBUF), X2, y1
   LineTo hSprite(H_BACKBUF), x1, y1
   UnsetPen
   
End Sub

Sub DrawBorder()
   
   ' draw the rasters
   ' horizontal
   For X = REAL_X1 To CL_C - 2
     With gMatrix(X, REAL_Y1 - 1)
         Select Case X
         Case REAL_X1:                    .sprite = "BD_TL"
         Case REAL_X1 + 1 To REAL_X2 - 1: .sprite = "BD_TM"
         Case REAL_X2:                    .sprite = "BD_TR"
         Case Else:                       .sprite = "BD_HR2"
       End Select
     End With
     gMatrix(X, REAL_Y2 + 1).sprite = "BD_HR1"
   Next X

   ' vertical
   For Y = REAL_Y1 - 1 To REAL_Y2 + 1
     gMatrix(REAL_X1 - 1, Y).sprite = "BD_VR1"
     gMatrix(CL_C - 1, Y).sprite = "BD_VR2"
     If REAL_X2 <> CL_C - 2 Then
       With gMatrix(REAL_X2 + 1, Y)
         Select Case Y
         Case REAL_Y1 - 1:        .sprite = "BD_TT"
         Case REAL_Y1 To REAL_Y2: .sprite = "BD_VR1"
         Case REAL_Y2 + 1:        .sprite = "BD_TB"
         End Select
       End With
     End If
   Next Y
   
   ' draw corners
   gMatrix(REAL_X1 - 1, REAL_Y1 - 1).sprite = "BD_LT"
   gMatrix(REAL_X1 - 1, REAL_Y2 + 1).sprite = "BD_LB"
   gMatrix(CL_C - 1, REAL_Y1 - 1).sprite = "BD_RT"
   gMatrix(CL_C - 1, REAL_Y2 + 1).sprite = "BD_RB"

End Sub

Sub LoadColor()

   ' load all colors
   aColors(1) = vbWhite
   aColors(2) = vbRed
   aColors(3) = vbGreen
   aColors(4) = vbBlue
   aColors(5) = vbYellow
   aColors(6) = vbCyan
   aColors(7) = vbMagenta
   
End Sub

Sub LoadLevel(Packname As String, Num As Long)

   Dim Row       As String
   Dim Col()     As String
   Dim InvNum    As Long
   
   If Num < 1 Then Exit Sub
   
   ClearMatrices
   
   On Error GoTo LoadError
   
   Open App.Path & "\data\levels\" & Packname & "\level_" & Right("00" & Trim(Str(Num)), 3) & ".lev" For Input As #1
   
     Line Input #1, MapName: MapName = "SYSTEM " & Num & " - " & Trim(MapName)
     
     For Y = REAL_Y1 To REAL_Y2
        
        Line Input #1, Row
        
        Col() = Split(Row, ".")
        
        For X = 0 To UBound(Col)
            
            With gMatrix(X + REAL_X1, Y)
               
               Col(X) = Trim(Col(X))
               
               .is_rot = IIf(LCase(Right(Col(X), 2)) = "=r", True, False)
                
               Col(X) = Replace(Col(X), "=r", vbNullString, , , vbTextCompare)
               
               .sprite = Col(X)
            
            End With
        
        Next X
        
     Next Y
   
     Input #1, InvNum
   
     ReDim Inventory(0)
   
     If InvNum > 0 Then
    
        ReDim Inventory(1 To InvNum)
     
        For Y = 1 To InvNum
        
           With Inventory(Y)
        
             Line Input #1, Row: Row = Trim(Row)
             
             .is_rot = IIf(LCase(Right(Row, 2)) = "=r", True, False)
                
             Row = Replace(Row, "=r", vbNullString, , , vbTextCompare)
             
             .sprite = Row
                      
           End With
           
        Next Y
       
     End If
   
   Close #1
   
   BuildInventory Inventory
   
   DrawBorder
   
   CalcMines
   
   PrepareMatrix
   
   CalcBeams
   
   Exit Sub

LoadError:
ShowError "Error occured in Game_Core::Loadlevel" & vbCrLf & "#Error: " & Err & vbCrLf & "Description: " & Err.Description
   
End Sub

Sub LoadObjects()

   ' load all object maps into the object array's

   Dim TmpRef()   As gObject
   Dim Maps()     As Variant
   Dim nObj       As String
   Dim Obj        As String
   Dim fBeams()   As String
   Dim fParts()   As String
   Dim tParts()   As String
   
   ' Format:
   '
   ' FromDir,ToDir>ToCol[;ToDir>ToCol;...],Oneway[-FromDir,ToDir>ToCol[;ToDir>ToCol;...],Oneway][-...]
      
   On Error GoTo LoadError
   
   Maps() = Array("mirror", _
                  "doublemirror", _
                  "slitdoublemirror", _
                  "prism", _
                  "refracter", _
                  "splitter", _
                  "angledsplitter", _
                  "filter", _
                  "polarizer", _
                  "oneway", _
                  "converter")
   
   For i = 0 To UBound(Maps)
   
     Open App.Path & "\data\objects\" & Maps(i) & ".map" For Input As #1
     
        Line Input #1, nObj
     
        ReDim TmpRef(1 To Val(nObj))
     
        For j = 1 To nObj
        
            Line Input #1, Obj
            
            fBeams() = Split(Obj, "-")
           
            ReDim TmpRef(j).fromdir(UBound(fBeams))
            ReDim TmpRef(j).todir(UBound(fBeams))
            ReDim TmpRef(j).tocol(UBound(fBeams))
                
            For k = 0 To UBound(fBeams)
                
                fParts() = Split(fBeams(k), ",")
                
                TmpRef(j).onedir = fParts(2)
                TmpRef(j).fromdir(k) = fParts(0)
                   
                tParts() = Split(fParts(1), ";")
                
                ReDim TmpRef(j).todir(k).Dir(UBound(tParts))
                ReDim TmpRef(j).tocol(k).Col(UBound(tParts))
                                    
                For l = 0 To UBound(tParts)
                               
                  TmpRef(j).todir(k).Dir(l) = Split(tParts(l), ">")(0)
                  TmpRef(j).tocol(k).Col(l) = Split(tParts(l), ">")(1)
            
                Next l
                
            Next k
                    
        Next j
     
     Close #1
    
     ' assign to corresponding object array
     Select Case Maps(i)
       Case "mirror":           Mirror() = TmpRef()
       Case "doublemirror":     DoubleMirror() = TmpRef()
       Case "slitdoublemirror": SlitDoubleMirror() = TmpRef()
       Case "prism":            Prism() = TmpRef()
       Case "refracter":        Refracter() = TmpRef()
       Case "splitter":         Splitter() = TmpRef()
       Case "angledsplitter":   AngledSplitter() = TmpRef
       Case "oneway":           Oneway() = TmpRef()
       Case "filter":           Filter() = TmpRef()
       Case "polarizer":        Polarizer() = TmpRef()
       Case "converter":        Converter() = TmpRef()
     End Select
   
   Next i
   
   Exit Sub

LoadError:
ShowError "Error occured in Game_Core::LoadObjects" & vbCrLf & "#Error: " & Err & vbCrLf & "Description: " & Err.Description
   
End Sub

Sub LoadCoordMap()

   Dim Total       As Long
   Dim cAniInfo    As Animations
   Dim Object      As String * 6
   Dim cAniVar(5)  As Variant
   Dim FSO         As FileSystemObject
   
   Set FSO = New FileSystemObject
   Set objInfoMap = New Collection

   If FSO.FileExists(App.Path & "\data\objects\coord.map") = False Then GoTo LoadError

   Open App.Path & "\data\objects\coord.map" For Binary As #1
     
     Get #1, , Total
     
     For i = 0 To Total
        Get #1, , Object
        Get #1, , cAniInfo
        cAniVar(0) = cAniInfo.x_start
        cAniVar(1) = cAniInfo.y_start
        cAniVar(2) = cAniInfo.frames
        cAniVar(3) = cAniInfo.Delay
        cAniVar(4) = cAniInfo.color
        cAniVar(5) = cAniInfo.beammethod
        objInfoMap.Add cAniVar, Trim(Object)
     Next i
   
   Close #1
   
   Exit Sub

LoadError:
ShowError "Error occured in Game_Core::LoadCoordMap" & vbCrLf & "'coord.map' not found."

End Sub

Function GetLightDir(Dir As Byte) As Direction

  ' select x/y displacement according to the direction
  With GetLightDir
  
    Select Case Dir
    Case 1: .x_disp = 0:  .y_disp = -1
    Case 2: .x_disp = 1:  .y_disp = -1
    Case 3: .x_disp = 1:  .y_disp = 0
    Case 4: .x_disp = 1:  .y_disp = 1
    Case 5: .x_disp = 0:  .y_disp = 1
    Case 6: .x_disp = -1: .y_disp = 1
    Case 7: .x_disp = -1: .y_disp = 0
    Case 8: .x_disp = -1: .y_disp = -1
    End Select
  
  End With
  
End Function

Function SwapInOut(Dir As Byte)
    ' do a 180 degrees direction shift
    Dim Tmp As Byte
    Tmp = Dir + 4
    If Tmp > 8 Then Tmp = Tmp - 8
    SwapInOut = Tmp
End Function

Function In2Out(Object As gObject, Dir As Byte, Optional Misc As Integer) As Variant
    
    ' get new direction(s) according to the object hit by the beam
    ' retruns array with all directions and colors
    
    ReDim Tmp(0) As Variant
    
    With Object
    
    For i = 0 To UBound(.fromdir)
       
       ' normal way
       If Dir = .fromdir(i) Then
          ReDim Tmp(UBound(.todir(i).Dir) * 2 + 1)
          For j = 0 To UBound(.todir(i).Dir)
              Tmp(j * 2) = .todir(i).Dir(j)
              Tmp(j * 2 + 1) = .tocol(i).Col(j)
          Next j
          In2Out = Tmp()
          Misc = 1
          Exit Function
       End If
       
       ' reversed way
       If .onedir = False Then
          For j = 0 To UBound(.todir(i).Dir)
             If Dir = .todir(i).Dir(j) Then
                ReDim Tmp(1)
                Tmp(0) = .fromdir(i)
                Tmp(1) = .tocol(i).Col(j)
                In2Out = Tmp()
                Misc = -1
                Exit Function
             End If
          Next j
       
       End If
    
    Next i
    
    End With
        
    In2Out = Tmp()
        
End Function

Sub GetObject(Object As String, Name As String, Optional Dir As Byte, Optional Col As Byte)

  Name = vbNullString
  Dir = 0
  Col = 0
    
  ' split up the object
  If Len(Object) = 0 Then
     Exit Sub
  End If
  
  Dim iX As Long
  
  Name = Left(Object, 2)
  Dir = Val(Mid(Object, 4))
  iX = InStr(1, Object, ":")
  If (iX > 0) Then
     Col = Val(Mid(Object, iX + 1))
  End If

End Sub

Sub PrepareMatrix()
  Dim oName As String
  
  ' prepares the matrix
  For cy = REAL_Y1 To REAL_Y2
    For cx = REAL_X1 To REAL_X2
       With gMatrix(cx, cy)
          GetObject .sprite, oName
          .cur_frame = 0
          .explode = EX_NONE
          .Time = 0
          Select Case oName
          Case "LS"
             .is_rot = False
             .cur_frame = LASER_START_FRAME
          Case "LR"
             .is_rot = True
             .cur_frame = LASER_START_FRAME
          Case "CO"
             .cur_frame = COIN_START_FRAME
          End Select
       End With
    Next cx
  Next cy

End Sub

Sub ExplodeWaitingObjects(tX As Long, tY As Long)

  ' make waiting barrels explode
  For cy = tY - 1 To tY + 1
     For cx = tX - 1 To tX + 1
        With gMatrix(cx, cy)
           If .explode = EX_WAIT Then
              .explode = EX_GO
              .cur_frame = 0
           End If
        End With
     Next cx
  Next cy
  
End Sub

Function CountCoins() As Long

  Dim cc    As Long
  Dim oName As String

  ' returns number of coins in field
  For cy = 0 To CL_R - 1
    For cx = 0 To CL_C - 1
       GetObject gMatrix(cx, cy).sprite, oName
       If oName = "CO" Then
          cc = cc + 1
       End If
    Next cx
  Next cy
  
  CountCoins = IIf(cc > 0, cc, -1)

End Function

Function GetMaxDir(Name As String) As Byte

  Select Case Name
  Case "MR", "LR", "AS", "SP", "OW", "PR", "LS", "CC"
       GetMaxDir = 8
  Case "RF", "MD", "MS", "PO":
       GetMaxDir = 4
  End Select

End Function

Function ObjectIsMovable(Name As String) As Boolean

  Select Case Name
  Case "MR", "LR", "AS", "SP", "OW", "PR", "LS", "CC", "RF", "MD", "MS", "PO", "FT"
       ObjectIsMovable = True
  End Select

End Function

Function RolBits(Value As Long, Direction As Integer) As Long

   Dim h As String
   
   h = Right("000000" & Hex(Value), 6)
   
   Select Case Direction
     Case -1
        RolBits = Val("&h" & Right(h, 4) & Left(h, 2) & "&")
     Case 1
        RolBits = Val("&h" & Right(h, 2) & Left(h, 4) & "&")
   End Select

End Function

Function InField(tX As Long, tY As Long, Optional Entire As Boolean = True) As Boolean

  ' test field boundaries
  InField = True
  
  Dim X2 As Long
  
  If Entire Then
     X2 = CL_C - 2
  Else
     X2 = REAL_X2
  End If
    
  If tX < REAL_X1 Or tX > X2 Or tY < REAL_Y1 Or tY > REAL_Y2 Then
     InField = False
  End If
  
End Function

Sub ClearMatrices(Optional gM As Boolean = True, Optional lM As Boolean = True, Optional eM As Boolean = True)
   
   On Error GoTo ClearError
   
   ' catch Error 10 and empty Matrices the harder way
   
   If gM Then
      For Y = 0 To CL_R - 1
         For X = 0 To CL_C - 1
             With gMatrix(X, Y)
                .cur_frame = 0
                .explode = EX_NONE
                .is_mov = False
                .is_rot = False
                .sprite = vbNullString
                .Time = 0
             End With
         Next X
      Next Y
   End If
   
   If lM Then
      ReDim lMatrix(CL_C - 1, CL_R - 1)
   End If
   
   If eM Then
      ReDim eMatrix(CL_C - 1, CL_R - 1)
   End If
   
   Exit Sub
   
ClearError:
ShowError "Error occured in Game_Core::ClearMatrices" & vbCrLf & "#Error: " & Err & vbCrLf & "Description: " & Err.Description
   
End Sub

Sub BuildInventory(Inventory() As gInventory)
  Dim iX As Long
 
  If UBound(Inventory) > 0 Then
  
     Y = REAL_Y1
     X = CL_C - 2

     For iX = 1 To UBound(Inventory)
    
         With gMatrix(X, Y)
            .sprite = Inventory(iX).sprite
            .is_rot = Inventory(iX).is_rot
            .is_mov = True
         End With
      
         Y = Y + 1
         If Y > REAL_Y2 And iX <> UBound(Inventory) Then
            Y = REAL_Y1
            X = X - 1
         End If
       
     Next iX
  
     REAL_X2 = X - 2
  
  End If
  
End Sub
