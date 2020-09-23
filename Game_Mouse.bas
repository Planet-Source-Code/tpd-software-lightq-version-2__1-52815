Attribute VB_Name = "Game_Mouse"
' ######################################
' ##
' ##  LightQ mouse routines
' ##

' mouse coordinates
Public ScrX As Long
Public ScrY As Long

Public CurX As Long
Public CurY As Long

Public SelectWithoutRotate As Boolean

Public MouseOutside As Boolean
Public MouseDown    As Boolean

Private ObjectPickedUp As Boolean

Sub mMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' save current screenposition
  ScrX = X
  ScrY = Y
  
  ' user interaction blocked?
  If LockGameMouse Or LockMenuMouse Then Exit Sub
  
  Dim tX As Long
  Dim tY As Long
  Dim oName As String
 
  ' current cell
  tX = X \ SP_W
  tY = Y \ SP_H
  
  ' mouse down?
  If MouseDown Then
  
     ' is position in playfield ?
     If InField(CurX, CurY) And MouseOutside = False Then
  
       ' user did move a whole cell?
       ' wanting to move an object
       If (tX <> CurX) Or (tY <> CurY) Then
  
         With gMatrix(CurX, CurY)
         
              ' we do not have a dragobject?
              If Len(DragObject) = 0 Then
         
                  ' the object is movable?
                  If .is_mov Then
                   
                     ' have we moved already?
                     ' this test is for the "clear" object, so it does not retrig the
                     ' sound evertime we move the mouse
                     If Not ObjectPickedUp Then
                     
                        ObjectPickedUp = True
                     
                        ' change mousepointer
                        SetPointer "move"
                     
                        ' play the pickup sound
                        PlaySound "PICKUP"
                        
                        ' set the dragobject to the current cellobject
                        DragObject = .sprite
                   
                        ' clear it's current position
                        ' to create a pickup
                        .sprite = vbNullString
                     
                        ' recalculate light beams
                        CalcBeams
                        
                     End If
                     
                  End If
                
              End If
            
         End With
   
       End If
  
    End If
    
 End If
  
End Sub

Sub mDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' clear cheats
  strCheat = vbNullString
  
  ' user interaction blocked?
  If LockGameMouse Or LockMenuMouse Then Exit Sub

  ' is mouse already down?
  ' indicating user pressed both buttons, exit
  If MouseDown = True Then
     Exit Sub
  End If

  ' only left mousebutton may move an object
  If (Button = vbLeftButton) Then
     MouseDown = True
  End If
  
  Dim tX As Long
  Dim tY As Long
  
  tX = X \ SP_W
  tY = Y \ SP_H
  
  ' current object cellposition
  If InField(tX, tY) Then
     
     ' when in edit mode don't rotate the object as it get's the focus
     ' but only select it at this time
     If ExecutionState = (EXEC_GAME + EXEC_EDIT) And (CurX <> tX Or CurY <> tY) Then
        SelectWithoutRotate = True
     Else
        SelectWithoutRotate = False
     End If
     
     CurX = tX
     CurY = tY
      
     MouseOutside = False
  Else
     MouseOutside = True
  End If
 
End Sub

Sub mUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' mouse not down
  MouseDown = False
  
  ' field is totally locked, no menu and no gamefield interaction
  If LockMenuMouse Then Exit Sub
  
  Dim tX    As Long
  Dim tY    As Long
  Dim oName As String
  Dim eName As String
  Dim oDir  As Byte
  Dim oCol  As Byte
  Dim lObj  As Long
  
  ' position of cell
  tX = X \ SP_W
  tY = Y \ SP_H
    
  ' be sure to reside in the matrix
  If tX < 0 Then tX = 0
  If tX > CL_C - 1 Then tX = CL_C - 1
  If tY < 0 Then tY = 0
  If tY > CL_R - 1 Then tY = CL_R - 1
    
  ' don't do menu's if a dragobject is placed on it
  If Len(DragObject) = 0 Then
    ' button selection
    Select Case ExecutionState
    Case EXEC_MENU
        If X >= 250 And X <= 550 Then
            Select Case Y
            
            ' MAINMENU->play
            Case 110 To 150
                PlaySound "BUTTON"
        
                ExecutionState = EXEC_SELECT
                YSelected = -1
            
            ' MAINMENU->edit
            Case 190 To 230
                PlaySound "BUTTON"
                
                If DEMOMODE Then
                    ShowError "Editor is disabled in DEMO version of LightQ", "DEMO"
                Else
                    ExecutionState = EXEC_PACK  'EXEC_GAME + EXEC_EDIT
                End If
            
            ' MAINMENU->settings
            Case 270 To 310
                PlaySound "BUTTON"
                
                ExecutionState = EXEC_SETTINGS
            
            ' [ADDON 2.0.1] MAINMENU->update
            Case 350 To 390
                PlaySound "BUTTON"
                
                If DEMOMODE Then
                    ShowError "Updates are disabled in DEMO version of LightQ", "DEMO"
                Else
                    ExecutionState = EXEC_UPDATE
                End If
                
            ' [ADDON 2.0.1] MAINMENU->visit site
            Case 430 To 470
                PlaySound "BUTTON"
                
                ShellExecute frmScreen.hwnd, vbNullString, "http://cp108629-b.landg1.lb.home.nl", vbNullString, vbNullString, SW_SHOWDEFAULT
                
            ' MAINMENU->exit
            Case 510 To 550
                PlaySound "BUTTON"
                
                If DoQuestion("Are you sure you want to leave LightQ?", "Exit", vbYesNo) = vbYes Then
                   frmScreen.Form_Unload (0)
                End If
            
            Case Else
                Exit Sub
            
            End Select
        End If
        
        Exit Sub
    Case EXEC_GAME
        If Y >= 572 And Y <= 600 Then
            Select Case X
            
            'GAME-> menu
            Case 125 To 217
                '[ADDON 2.0.1] leave game confirmation box
                If DoQuestion("Are you sure you want to exit to main menu?", , vbYesNo) = vbYes Then
                   'ExecutionState = EXEC_MENU
                   ExecutionState = EXEC_SELECT
                End If
                Exit Sub
            
            'GAME->cleanup [CURRENTLY UNUSED]
            Case 227 To 319
              
            'GAME-> restart
            Case 329 To 421
                GameLoop
                Exit Sub
      
            'GAME-> continue
            Case 431 To 523
                With frmScreen
                   If Completed = True Then
                      If .Levels.Tag = "ALL_DONE" Then
                         ExecutionState = EXEC_SELECT
                      Else
                         GameLoop .levPacks.List(.levPacks.ListIndex), Val(.Levels.List(.Levels.ListIndex))
                      End If
                   End If
                End With
                Exit Sub
              
            'GAME->rotate object left
            Case 593 To 686
                tX = CurX
                tY = CurY
                Button = vbLeftButton
                
            'GAME->rotate object right
            Case 693 To 786
                tX = CurX
                tY = CurY
                Button = vbRightButton
      
            End Select
        End If
    Case EXEC_GAME + EXEC_EDIT
        If Y >= 572 And Y <= 600 Then
            Select Case X
            
            ' EDIT->menu
            Case 125 To 217
                If DoQuestion("Are you sure you want to leave the editor?", , vbYesNo) = vbYes Then
                   ExecutionState = EXEC_MENU
                End If
                Exit Sub
                
            ' EDIT->select object
            Case 235 To 235 + SP_W
                ShowEditObjects
                Exit Sub
            
            ' EDIT->object settings
            Case 283 To 375
                ShowEditSettings
                Exit Sub
        
            'EDIT->save
            Case 386 To 478
                With frmScreen
                   SaveLevel .lstEditPack.List(.lstEditPack.ListIndex), .ed_ver_passw
                End With
                Exit Sub
            
            'EDIT->load
            Case 489 To 581
                With frmScreen
                   LoadBinLevel .lstEditPack.List(.lstEditPack.ListIndex), .ed_ver_passw
                End With
                Exit Sub
            
            'EDIT->rotate object left
            Case 593 To 686
                tX = CurX
                tY = CurY
                Button = vbLeftButton
                
            'EDIT->rotate object right
            Case 693 To 786
                tX = CurX
                tY = CurY
                Button = vbRightButton
            
            End Select
        End If
    End Select
    
  End If
  
  ' user interaction blocked?
  If LockGameMouse Then Exit Sub
  
  With gMatrix(tX, tY)
  
     ' are we in edit mode and selected an object?
     If ExecutionState = (EXEC_GAME + EXEC_EDIT) And YSelected > -1 Then
        
        ' get length of objects name
        lObj = Len(EditObjects(SelectedObjectIndex()).Object)
        
        ' object placed inside field?
        If InField(tX, tY) Then
        
           If Len(DragObject) = 0 Then
           
              ' is the cell empty, we don't have a drag and we have a selected object ?
              ' set the object
              If Len(.sprite) = 0 Or lObj = 0 Then
         
                   If lObj > 0 Then
                      ' we want place a new object in the game
                      ' play drop sound
                      PlaySound "DROP"
                 
                      ' all objects may be moved in editmode
                      .is_mov = True
                   Else
                      ' we want to remove an object
                      ' play remove sound
                      If Len(.sprite) Then
                         PlaySound "REMOVE"
                      End If
                 
                      ' we are removing an object so nothing to move anymore
                      .is_mov = False
                   End If
             
                   ' apply the object and set it's properties
                   .sprite = EditObjects(SelectedObjectIndex()).Object
                   .is_rot = EditObjects(SelectedObjectIndex()).is_rot
                 
                   ' apply settings to editmatrix also
                   FixEditMatrix
            
                   .Time = 0
                   .cur_frame = 0
                   .explode = EX_NONE
            
                   ' calculate beams
                   CalcBeams
          
                   ' for our selection box
                   CurX = tX
                   CurY = tY
           
               End If
             
           End If
             
        End If
            
     End If
  
     ' do we have a dragobject?
     ' then we want to drop it in it's new position
     If Len(DragObject) Then
     
        ' is the cell empty
        If Len(.sprite) Then
           ' no we respot it to it's old place
           
           ' play respot sound
           PlaySound "RESPOT"
     
           ' place back the object
           ' we removed earlier
           gMatrix(CurX, CurY).sprite = DragObject
           
        Else
           ' cell is empty
           ' so we place the dragobject here
           
           ' play drop sound
           PlaySound "DROP"
     
           ' move the old position settings to the
           ' new position
           .sprite = DragObject
           .is_mov = True
           .is_rot = gMatrix(CurX, CurY).is_rot
                   
           If ExecutionState = (EXEC_GAME + EXEC_EDIT) Then
              
              ' apply settings to editmatrix also when in edit mode
              With eMatrix(tX, tY)
                 .is_rot = eMatrix(CurX, CurY).is_rot
                 .is_mov = eMatrix(CurX, CurY).is_mov
              End With
              
           End If
                   
           ' do not alter when dropped on same position
           If CurX <> tX Or CurY <> tY Then
              
              With gMatrix(CurX, CurY)
                .Time = 0
                .cur_frame = 0
                .explode = EX_NONE
                .is_rot = False
                .is_mov = False
              End With
              
              If ExecutionState = (EXEC_GAME + EXEC_EDIT) Then
                
                With eMatrix(CurX, CurY)
                   .is_rot = False
                   .is_mov = False
                End With
              
              End If
              
           End If
        
           ' apply new position as current
           CurX = tX
           CurY = tY
        
        End If
     
        ' recalculate beams
        CalcBeams
        
     Else
        
        ' we don't have a dragobject
        ' so the user hasn't moved
        ' rotate the object
        
        ' has the cell an object?
        If Len(.sprite) Then
     
             ' is it rotatable?
             If .is_rot Then
     
                 If SelectWithoutRotate = False Then
     
                    ' play the rotation sound
                    PlaySound "ROTATE"
     
                    ' get object properties
                    GetObject .sprite, oName, oDir, oCol
         
                    Select Case Button
                    Case vbLeftButton
                      ' rotate counter-clockwise
                      oDir = oDir - 1
                      If oDir = 0 Then oDir = GetMaxDir(oName)
                    Case vbRightButton
                      ' rotate clockwise
                      oDir = oDir + 1
                      If oDir > GetMaxDir(oName) Then oDir = 1
                    End Select
        
                    ' make the object with it's new direction
                    .sprite = oName & "_" & oDir & IIf(oCol > 0, ":" & oCol, vbNullString)
             
                    ' recalculate beams
                    CalcBeams
          
                 End If
          
             End If
           
             ' if we destroyed position
             ' reset it here to make sure
             CurX = tX
             CurY = tY
        
        End If
        
     End If
          
  End With
   
  ' normal pointer
  SetPointer "normal"
  
  ' no dragobjects
  ObjectPickedUp = False
  
  ' clear selection-first flag
  SelectWithoutRotate = False
  
  ' always lose the dragobject when mouse up
  DragObject = vbNullString
  
End Sub
